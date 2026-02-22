"""
IndexService — in-memory inverted index with Snowball stemming.

Architecture:
  - Text index: built inline on first access (~3s for 873 paras).
  - Page map:   built in background thread via execute_on_main_thread,
                in chunks of 50 paragraphs.  Foreground queries compute
                pages on-demand for result paragraphs only; those results
                also feed the shared page map so background can skip them.
  - Prewarm:    called by mcp_server after each tool call to ensure
                the active document's index is ready.

Language detected from UNO CharLocale.  Stemming via bundled snowballstemmer.
"""

import logging
import re
import threading
import time
import unicodedata
from typing import Any, Dict, List, Optional, Set

logger = logging.getLogger(__name__)

# ── Language mapping (ISO 639-1 -> snowballstemmer algorithm) ─────────

_ISO_TO_SNOWBALL = {
    "ar": "arabic",    "hy": "armenian",  "eu": "basque",
    "ca": "catalan",   "da": "danish",    "nl": "dutch",
    "en": "english",   "eo": "esperanto", "et": "estonian",
    "fi": "finnish",   "fr": "french",    "de": "german",
    "el": "greek",     "hi": "hindi",     "hu": "hungarian",
    "id": "indonesian","ga": "irish",     "it": "italian",
    "lt": "lithuanian","ne": "nepali",    "no": "norwegian",
    "nb": "norwegian", "nn": "norwegian",
    "pt": "portuguese","ro": "romanian",  "ru": "russian",
    "sr": "serbian",   "es": "spanish",   "sv": "swedish",
    "ta": "tamil",     "tr": "turkish",   "yi": "yiddish",
}

# ── Stop words per language ───────────────────────────────────────────

_STOP_WORDS = {
    "french": frozenset({
        "au", "aux", "avec", "ce", "ces", "cette", "dans", "de",
        "des", "du", "elle", "en", "est", "et", "il", "ils", "je",
        "la", "le", "les", "leur", "leurs", "lui", "ma", "mais",
        "me", "mes", "mon", "ne", "ni", "nos", "notre", "nous",
        "on", "ou", "par", "pas", "pour", "qu", "que", "qui", "sa",
        "se", "ses", "si", "son", "sur", "ta", "te", "tes", "ton",
        "tu", "un", "une", "vos", "votre", "vous",
    }),
    "english": frozenset({
        "a", "an", "and", "are", "as", "at", "be", "but", "by",
        "for", "from", "had", "has", "he", "her", "his", "if", "in",
        "is", "it", "its", "my", "no", "not", "of", "on", "or",
        "our", "she", "so", "the", "to", "up", "us", "was", "we",
    }),
    "german": frozenset({
        "aber", "als", "am", "an", "auch", "auf", "aus", "bei",
        "bin", "bis", "da", "das", "dem", "den", "der", "des",
        "die", "du", "ein", "er", "es", "fur", "hat", "ich", "ihr",
        "im", "in", "ist", "ja", "mir", "mit", "nach", "nicht",
        "noch", "nun", "nur", "ob", "oder", "sie", "so", "und",
        "uns", "vom", "von", "vor", "was", "wir", "zu", "zum",
        "zur",
    }),
    "spanish": frozenset({
        "a", "al", "con", "de", "del", "el", "en", "es", "la",
        "las", "lo", "los", "no", "por", "que", "se", "su", "un",
        "una", "y",
    }),
    "italian": frozenset({
        "a", "al", "che", "con", "da", "del", "di", "e", "il",
        "in", "la", "le", "lo", "non", "per", "si", "su", "un",
        "una",
    }),
    "portuguese": frozenset({
        "a", "ao", "com", "da", "de", "do", "e", "em", "na", "no",
        "o", "os", "por", "que", "se", "um", "uma",
    }),
}

_STOP_WORDS_FALLBACK = frozenset()

# ── Tokenisation ──────────────────────────────────────────────────────

_PUNCT_RE = re.compile(r'[^\w\s]', re.UNICODE)
_MIN_TOKEN_LEN = 2


def _deaccent(text):
    """Remove diacritics: e->e, c->c, u->u, etc."""
    nfkd = unicodedata.normalize('NFKD', text)
    return ''.join(c for c in nfkd if unicodedata.category(c) != 'Mn')


def _raw_tokens(text):
    """Lowercase, deaccent, strip punctuation, split."""
    cleaned = _PUNCT_RE.sub(' ', _deaccent(text.lower()))
    return [t for t in cleaned.split() if len(t) >= _MIN_TOKEN_LEN]


# ── Per-document index ────────────────────────────────────────────────

class _DocIndex:
    """Inverted index for one document."""

    __slots__ = ('terms', 'para_texts', 'para_elements',
                 'para_count', 'build_ms', 'language',
                 'page_map', 'page_map_complete', 'page_paras',
                 '_bg_started')

    def __init__(self):
        self.terms = {}            # stem -> set[int]
        self.para_texts = {}       # int -> str
        self.para_elements = []    # UNO paragraph elements
        self.para_count = 0
        self.build_ms = 0.0
        self.language = "english"
        # Page map — incrementally built (foreground + background)
        self.page_map = {}         # int -> int (para_index -> page)
        self.page_map_complete = False
        self.page_paras = {}       # int -> list[int] (page -> paras)
        self._bg_started = False

    # ── Query primitives ──

    def query_and(self, stem_groups):
        """AND: each group is a list of stems (OR within group)."""
        if not stem_groups:
            return set()
        sets = []
        for group in stem_groups:
            s = set()
            for stem in group:
                ps = self.terms.get(stem)
                if ps:
                    s |= ps
            if not s:
                return set()
            sets.append(s)
        sets.sort(key=len)
        result = sets[0].copy()
        for s in sets[1:]:
            result &= s
            if not result:
                return result
        return result

    def query_or(self, stems):
        """OR: paragraphs containing ANY stem."""
        result = set()
        for stem in stems:
            ps = self.terms.get(stem)
            if ps:
                result |= ps
        return result

    def query_not(self, include, exclude_stems):
        """Remove paragraphs matching exclude stems."""
        result = include.copy()
        for stem in exclude_stems:
            ps = self.terms.get(stem)
            if ps:
                result -= ps
        return result

    def query_near(self, stems_a, stems_b, distance):
        """Paragraphs where any stem_a and any stem_b within ±distance."""
        set_a = set()
        for s in stems_a:
            ps = self.terms.get(s)
            if ps:
                set_a |= ps
        set_b = set()
        for s in stems_b:
            ps = self.terms.get(s)
            if ps:
                set_b |= ps
        if not set_a or not set_b:
            return set()
        result = set()
        sorted_b = sorted(set_b)
        for pa in set_a:
            for pb in sorted_b:
                if abs(pa - pb) <= distance:
                    result.add(pa)
                    result.add(pb)
                elif pb > pa + distance:
                    break
        return result


# ── Background page map builder ───────────────────────────────────────

_PAGE_CHUNK = 50


def _build_page_chunk(doc, idx, start, end):
    """Build one chunk of the page map.  Runs on VCL main thread.

    Saves/restores the view cursor so the user doesn't see scrolling.
    """
    try:
        controller = doc.getCurrentController()
        vc = controller.getViewCursor()
        # Save cursor position
        saved_range = vc.getStart()
        for pi in range(start, min(end, len(idx.para_elements))):
            if pi in idx.page_map:
                continue  # already computed by foreground
            try:
                vc.gotoRange(idx.para_elements[pi].getStart(), False)
                page = vc.getPage()
            except Exception:
                page = idx.page_map.get(pi - 1, 1) if pi > 0 else 1
            idx.page_map[pi] = page
            if page not in idx.page_paras:
                idx.page_paras[page] = []
            idx.page_paras[page].append(pi)
        # Restore cursor position
        try:
            vc.gotoRange(saved_range, False)
        except Exception:
            pass
    except Exception as e:
        logger.warning("Page chunk %d-%d failed: %s", start, end, e)


def _background_page_map(doc, idx):
    """Background thread: build full page map in chunks."""
    from main_thread_executor import execute_on_main_thread
    t0 = time.perf_counter()
    total = len(idx.para_elements)
    for start in range(0, total, _PAGE_CHUNK):
        if idx.page_map_complete:
            return
        try:
            execute_on_main_thread(
                _build_page_chunk, doc, idx,
                start, start + _PAGE_CHUNK,
                timeout=30.0)
        except Exception as e:
            logger.warning("Background page map chunk failed: %s", e)
            # Don't block on errors — skip chunk
        # Small yield between chunks so tool calls get priority
        time.sleep(0.05)
    idx.page_map_complete = True
    elapsed = round((time.perf_counter() - t0) * 1000, 1)
    logger.info("Background page map complete: %d pages, %.1fms",
                 len(idx.page_paras), elapsed)


# ── Service ───────────────────────────────────────────────────────────

class IndexService:
    """Per-document inverted index with Snowball stemming."""

    def __init__(self, writer):
        self._writer = writer
        self._base = writer._base
        self._cache = {}           # doc_key -> _DocIndex
        self._stemmers = {}        # lang -> StemmerInstance

    def invalidate_cache(self, doc=None):
        if doc is None:
            # Mark all as complete to stop background threads
            for idx in self._cache.values():
                idx.page_map_complete = True
            self._cache.clear()
        else:
            key = self._base.doc_key(doc)
            old = self._cache.get(key)
            if old:
                old.page_map_complete = True  # stop background
            self._cache.pop(key, None)

    # ── Stemmer management ────────────────────────────────────────

    def _get_stemmer(self, lang):
        cached = self._stemmers.get(lang)
        if cached is not None:
            return cached
        try:
            import snowballstemmer
            s = snowballstemmer.stemmer(lang)
            self._stemmers[lang] = s
            return s
        except (ImportError, KeyError):
            logger.warning("No stemmer for '%s', falling back to english",
                           lang)
            if lang != "english":
                return self._get_stemmer("english")
            return None

    def _detect_language(self, doc):
        try:
            text = doc.getText()
            enum = text.createEnumeration()
            if enum.hasMoreElements():
                first_para = enum.nextElement()
                locale = first_para.getPropertyValue("CharLocale")
                iso = locale.Language
                lang = _ISO_TO_SNOWBALL.get(iso)
                if lang:
                    return lang
        except Exception as e:
            logger.debug("Language detection failed: %s", e)
        return "english"

    def _stem(self, stemmer, tokens, stop_words):
        stems = []
        for t in tokens:
            if t in stop_words:
                continue
            stems.append(stemmer.stemWord(t))
        return stems

    # ── Index build ───────────────────────────────────────────────

    def _get_index(self, doc):
        """Get or build the inverted index.  Returns (index, was_cached)."""
        key = self._base.doc_key(doc)
        cached = self._cache.get(key)
        if cached is not None:
            return cached, True

        t0 = time.perf_counter()
        lang = self._detect_language(doc)
        stemmer = self._get_stemmer(lang)
        stop_words = _STOP_WORDS.get(lang, _STOP_WORDS_FALLBACK)

        idx = _DocIndex()
        idx.language = lang
        text_obj = doc.getText()
        enum = text_obj.createEnumeration()
        para_i = 0

        while enum.hasMoreElements():
            el = enum.nextElement()
            idx.para_elements.append(el)
            if el.supportsService("com.sun.star.text.Paragraph"):
                text = el.getString()
                idx.para_texts[para_i] = text
                raw = _raw_tokens(text)
                if stemmer:
                    stems = self._stem(stemmer, raw, stop_words)
                else:
                    stems = [t for t in raw if t not in stop_words]
                for stem in stems:
                    s = idx.terms.get(stem)
                    if s is None:
                        s = set()
                        idx.terms[stem] = s
                    s.add(para_i)
            else:
                idx.para_texts[para_i] = "[Table]"
            para_i += 1

        idx.para_count = para_i
        idx.build_ms = round((time.perf_counter() - t0) * 1000, 1)
        self._cache[key] = idx
        logger.info("Index built [%s]: %d paras, %d stems, %.1fms",
                     lang, para_i, len(idx.terms), idx.build_ms)
        return idx, False

    def _start_background_page_map(self, doc, idx):
        """Kick off background page map build if not already started."""
        if idx._bg_started or idx.page_map_complete:
            return
        idx._bg_started = True
        t = threading.Thread(
            target=_background_page_map, args=(doc, idx),
            daemon=True, name="idx-pagemap")
        t.start()
        logger.info("Background page map build started")

    def _get_pages_for(self, doc, idx, para_indices):
        """Get page numbers for specific paragraphs (on-demand).

        Saves/restores cursor position to avoid visible scrolling.
        Caches results in the shared page_map.
        """
        missing = [pi for pi in para_indices
                   if pi not in idx.page_map]
        if missing:
            try:
                controller = doc.getCurrentController()
                vc = controller.getViewCursor()
                saved_range = vc.getStart()
                for pi in missing:
                    if pi < len(idx.para_elements):
                        try:
                            vc.gotoRange(
                                idx.para_elements[pi].getStart(), False)
                            page = vc.getPage()
                        except Exception:
                            page = 1
                        idx.page_map[pi] = page
                        if page not in idx.page_paras:
                            idx.page_paras[page] = []
                        idx.page_paras[page].append(pi)
                try:
                    vc.gotoRange(saved_range, False)
                except Exception:
                    pass
            except Exception:
                pass
        return {pi: idx.page_map.get(pi) for pi in para_indices}

    # ── Prewarm (called by mcp_server after each tool call) ──────

    def prewarm(self):
        """Prewarm text index for the active Writer document.

        Builds text index inline (if not cached).  Does NOT start
        the background page map — that only runs when a query needs
        page numbers (include_pages / around_page).
        Called after every tool execution — instant if already cached.
        """
        try:
            desktop = self._base.desktop
            doc = desktop.getCurrentComponent()
            if doc is None:
                return
            if not doc.supportsService(
                    "com.sun.star.text.TextDocument"):
                return
            self._get_index(doc)
        except Exception:
            pass

    # ── Query parsing ─────────────────────────────────────────────

    def _stem_query_tokens(self, text, stemmer, stop_words):
        raw = _raw_tokens(text)
        stems = []
        dropped = []
        for t in raw:
            if t in stop_words:
                dropped.append(t)
            else:
                stems.append(stemmer.stemWord(t) if stemmer else t)
        return stems, dropped

    def _parse_query(self, query, stemmer, stop_words):
        result = {
            "and_stems": [], "or_stems": [], "not_stems": [],
            "near": [], "dropped_stops": [], "mode": "and",
            "error": None,
        }

        not_split = re.split(r'\bNOT\b', query, flags=re.IGNORECASE)
        main_part = not_split[0].strip()
        for part in not_split[1:]:
            stems, dropped = self._stem_query_tokens(
                part, stemmer, stop_words)
            result["not_stems"].extend(stems)
            result["dropped_stops"].extend(dropped)

        # NEAR/N
        near_match = re.search(
            r'(.+?)\s+NEAR/(\d+)\s+(.+)',
            main_part, re.IGNORECASE)
        if near_match:
            left, dropped_l = self._stem_query_tokens(
                near_match.group(1), stemmer, stop_words)
            dist = int(near_match.group(2))
            right, dropped_r = self._stem_query_tokens(
                near_match.group(3), stemmer, stop_words)
            result["dropped_stops"].extend(dropped_l + dropped_r)
            if left and right:
                result["near"].append((left, right, dist))
                result["mode"] = "near"
            elif not left and not right:
                result["error"] = "NEAR terms are all stop words"
            return result

        has_and = bool(re.search(r'\bAND\b', main_part,
                                  re.IGNORECASE))
        has_or = bool(re.search(r'\bOR\b', main_part,
                                 re.IGNORECASE))
        if has_and and has_or:
            result["error"] = (
                "Mixed AND/OR not supported. "
                "Use one operator per query.")
            return result

        if has_or:
            chunks = re.split(r'\bOR\b', main_part,
                              flags=re.IGNORECASE)
            for chunk in chunks:
                stems, dropped = self._stem_query_tokens(
                    chunk, stemmer, stop_words)
                result["or_stems"].extend(stems)
                result["dropped_stops"].extend(dropped)
            result["mode"] = "or"
        else:
            if has_and:
                chunks = re.split(r'\bAND\b', main_part,
                                  flags=re.IGNORECASE)
            else:
                chunks = [main_part]
            for chunk in chunks:
                stems, dropped = self._stem_query_tokens(
                    chunk, stemmer, stop_words)
                for stem in stems:
                    result["and_stems"].append([stem])
                result["dropped_stops"].extend(dropped)
            result["mode"] = "and"

        return result

    # ── Public API ────────────────────────────────────────────────

    def search_boolean(self, query, max_results=20,
                       context_paragraphs=1,
                       around_page=None, page_radius=1,
                       include_pages=False,
                       file_path=None):
        """Boolean full-text search with Snowball stemming.

        around_page + page_radius: restrict to pages within range.
        include_pages: add page numbers to results (on-demand, fast).
        """
        try:
            doc = self._base.resolve_document(file_path)
            idx, was_cached = self._get_index(doc)
            self._start_background_page_map(doc, idx)

            stemmer = self._get_stemmer(idx.language)
            stop_words = _STOP_WORDS.get(idx.language,
                                          _STOP_WORDS_FALLBACK)
            parsed = self._parse_query(query, stemmer, stop_words)

            if parsed["error"]:
                return {"success": False, "error": parsed["error"]}

            mode = parsed["mode"]
            and_stems = parsed["and_stems"]
            or_stems = parsed["or_stems"]
            not_stems = parsed["not_stems"]
            near = parsed["near"]

            all_positive = []
            for group in and_stems:
                all_positive.extend(group)
            all_positive.extend(or_stems)
            if near:
                for left, right, _ in near:
                    all_positive.extend(left + right)

            # Execute query
            if mode == "near" and near:
                left, right, dist = near[0]
                hits = idx.query_near(left, right, dist)
            elif or_stems:
                hits = idx.query_or(or_stems)
            elif and_stems:
                hits = idx.query_and(and_stems)
            else:
                return {
                    "success": False,
                    "error": "No search terms after stop-word filtering",
                    "dropped_stops": parsed["dropped_stops"],
                }

            if not_stems:
                hits = idx.query_not(hits, not_stems)

            # Page filtering with around_page
            needs_pages = around_page is not None
            if needs_pages:
                # Compute pages for all hits on-demand
                self._get_pages_for(doc, idx, hits)
                lo = around_page - page_radius
                hi = around_page + page_radius
                hits = {p for p in hits
                        if lo <= idx.page_map.get(p, 0) <= hi}

            total = len(hits)
            selected = sorted(hits)[:max_results]

            # On-demand pages for result + context paragraphs
            want_pages = needs_pages or include_pages
            if want_pages and selected:
                all_paras = set(selected)
                for pi in selected:
                    ctx_lo = max(0, pi - context_paragraphs)
                    ctx_hi = min(idx.para_count,
                                 pi + context_paragraphs + 1)
                    for j in range(ctx_lo, ctx_hi):
                        all_paras.add(j)
                self._get_pages_for(doc, idx, all_paras)

            bookmark_map = self._writer.tree.get_mcp_bookmark_map(doc)

            results = []
            for para_i in selected:
                ctx_lo = max(0, para_i - context_paragraphs)
                ctx_hi = min(idx.para_count,
                             para_i + context_paragraphs + 1)
                context = [
                    {"index": j,
                     "text": idx.para_texts.get(j, "")}
                    for j in range(ctx_lo, ctx_hi)
                ]
                if want_pages:
                    for c in context:
                        c["page"] = idx.page_map.get(c["index"])

                matched = [s for s in all_positive
                           if para_i in idx.terms.get(s, set())]

                entry = {
                    "paragraph_index": para_i,
                    "text": idx.para_texts.get(para_i, ""),
                    "matched_stems": matched,
                    "context": context,
                }
                if want_pages:
                    entry["page"] = idx.page_map.get(para_i)

                nearest = (self._writer.tree
                           .find_nearest_heading_bookmark(
                               para_i, bookmark_map))
                if nearest:
                    entry["nearest_heading"] = nearest
                results.append(entry)

            resp = {
                "success": True,
                "query": query,
                "mode": mode,
                "language": idx.language,
                "total_found": total,
                "returned": len(results),
                "matches": results,
                "index": {
                    "paragraphs": idx.para_count,
                    "unique_stems": len(idx.terms),
                    "build_ms": idx.build_ms,
                    "cached": was_cached,
                    "page_map_progress": (
                        "%d/%d" % (len(idx.page_map), idx.para_count)),
                    "page_map_complete": idx.page_map_complete,
                },
            }
            if near:
                resp["near"] = {"left": near[0][0],
                                "right": near[0][1],
                                "distance": near[0][2]}
            if parsed["dropped_stops"]:
                resp["dropped_stops"] = parsed["dropped_stops"]
            return resp
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_index_stats(self, file_path=None):
        """Index statistics + top 20 most frequent stems."""
        try:
            doc = self._base.resolve_document(file_path)
            idx, was_cached = self._get_index(doc)

            top = sorted(idx.terms.items(),
                         key=lambda x: len(x[1]),
                         reverse=True)[:20]

            resp = {
                "success": True,
                "language": idx.language,
                "paragraphs": idx.para_count,
                "unique_stems": len(idx.terms),
                "build_ms": idx.build_ms,
                "cached": was_cached,
                "page_map_progress": (
                    "%d/%d" % (len(idx.page_map), idx.para_count)),
                "page_map_complete": idx.page_map_complete,
                "top_stems": [{"stem": t, "paragraphs": len(s)}
                              for t, s in top],
            }
            return resp
        except Exception as e:
            return {"success": False, "error": str(e)}
