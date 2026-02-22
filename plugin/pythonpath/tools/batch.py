"""Batch execution with human stop conditions, follow mode, and
batch variables ($last, $step.N) for chaining operations."""

import logging
import re
import time
from .base import McpTool

logger = logging.getLogger(__name__)

_STOP_KEYWORDS = {"STOP", "CANCEL"}
_BLOCKED_PHASES = {"pause", "human validation", "locked"}

# Keys in tool results that hint at a document location
_LOCATION_KEYS = (
    "paragraph_index", "para_index", "locator",
    "page", "page_number",
)

# ── Batch variable resolution ──────────────────────────────────────

_VAR_RE = re.compile(
    r'\$(?:'
    r'last\.bookmark'                 # $last.bookmark
    r'|last(?:([+-])(\d+))?'          # $last, $last+1, $last-2
    r'|step\.(\d+)\.bookmark'         # $step.1.bookmark
    r'|step\.(\d+)(?:([+-])(\d+))?'   # $step.1, $step.2+1
    r')')


def _extract_step_info(result):
    """Extract paragraph_index and bookmark from a tool result.

    Returns (para_index: int|None, bookmark: str|None).
    """
    if not isinstance(result, dict):
        return None, None
    pi = result.get("paragraph_index")
    if pi is None:
        pi = result.get("para_index")
    if pi is not None:
        pi = int(pi)
    bm = result.get("bookmark")
    return pi, bm


def _resolve_var(match, batch_vars):
    """Replace a single $var match with its resolved value.

    Returns string representation. Bookmarks resolve to
    'bookmark:_mcp_xxx' for use as locators.
    """
    full = match.group(0)

    # $last.bookmark
    if full == "$last.bookmark":
        bm = batch_vars.get("$last.bookmark")
        return ("bookmark:%s" % bm) if bm else full

    # $step.N.bookmark
    if ".bookmark" in full:
        # extract step number
        m = re.match(r'\$step\.(\d+)\.bookmark', full)
        if m:
            bm = batch_vars.get("$step.%s.bookmark" % m.group(1))
            return ("bookmark:%s" % bm) if bm else full
        return full

    # $last, $last+N, $last-N
    if full.startswith("$last"):
        base = batch_vars.get("$last")
        if base is None:
            return full
        sign = match.group(1)
        offset_str = match.group(2)
        offset = int(offset_str) if offset_str else 0
        if sign == '-':
            offset = -offset
        return str(base + offset)

    # $step.N, $step.N+M
    step_num = match.group(4)
    if step_num is not None:
        key = "$step.%s" % step_num
        base = batch_vars.get(key)
        if base is None:
            return full
        sign = match.group(5)
        offset_str = match.group(6)
        offset = int(offset_str) if offset_str else 0
        if sign == '-':
            offset = -offset
        return str(base + offset)

    return full


def _resolve_batch_vars(args, batch_vars):
    """Recursively resolve $last / $step.N in args dict.

    String values that are EXACTLY a variable (e.g. "$last") become
    integers.  Strings containing a variable within text
    (e.g. "paragraph:$last") get string substitution.
    """
    if not batch_vars:
        return args
    if isinstance(args, dict):
        return {k: _resolve_batch_vars(v, batch_vars)
                for k, v in args.items()}
    if isinstance(args, list):
        return [_resolve_batch_vars(v, batch_vars) for v in args]
    if isinstance(args, str) and '$' in args:
        # Pure variable reference → return as int
        pure = _VAR_RE.fullmatch(args)
        if pure:
            resolved = _resolve_var(pure, batch_vars)
            try:
                return int(resolved)
            except ValueError:
                return resolved
        # Embedded variable → string substitution
        return _VAR_RE.sub(
            lambda m: _resolve_var(m, batch_vars), args)
    return args


# ── Stop conditions ────────────────────────────────────────────────

def _scan_stop_conditions(services, file_path=None):
    """Lightweight check for human stop signals.

    Returns (should_stop, reason, details).
    """
    # 1. Scan comments for STOP/CANCEL
    try:
        comments_result = services.comments.list_comments(
            file_path=file_path)
        for c in comments_result.get("comments", []):
            content_upper = c.get("content", "").upper().strip()
            first_word = content_upper.split()[0] if content_upper else ""
            if first_word in _STOP_KEYWORDS:
                return (True,
                        "Human %s comment detected" % first_word,
                        {"author": c.get("author", ""),
                         "comment": c.get("content", "")})
    except Exception as e:
        logger.warning("Stop condition check (comments) failed: %s", e)

    # 2. Check workflow dashboard phase
    try:
        wf = services.comments.get_workflow_status(file_path=file_path)
        if wf.get("found"):
            phase = wf.get("status", {}).get("Phase", "").strip().lower()
            if phase in _BLOCKED_PHASES:
                return (True,
                        "Workflow phase is '%s'" % wf["status"]["Phase"],
                        {"phase": wf["status"]["Phase"]})
    except Exception as e:
        logger.warning("Stop condition check (workflow) failed: %s", e)

    return (False, None, None)


# ── Follow ─────────────────────────────────────────────────────────

def _follow_result(services, result):
    """Scroll the view to the location implied by a tool result."""
    if not isinstance(result, dict):
        return
    try:
        doc = services.base.resolve_document()
        controller = doc.getCurrentController()
        vc = controller.getViewCursor()

        # Direct page reference
        page = result.get("page") or result.get("page_number")
        if page and isinstance(page, int):
            vc.jumpToPage(page)
            return

        # Paragraph index → move view cursor there
        para_idx = result.get("paragraph_index")
        if para_idx is None:
            para_idx = result.get("para_index")
        if para_idx is not None and isinstance(para_idx, int):
            text = doc.getText()
            enum = text.createEnumeration()
            idx = 0
            while enum.hasMoreElements():
                para = enum.nextElement()
                if idx == para_idx:
                    vc.gotoStart(False)
                    vc.gotoRange(para.getStart(), False)
                    return
                idx += 1
    except Exception:
        pass


# ── ExecuteBatch tool ──────────────────────────────────────────────

class ExecuteBatch(McpTool):
    name = "execute_batch"
    description = (
        "Execute multiple tool calls in a single request (one human "
        "approval). Operations run sequentially with batch mode "
        "(caches/indexing deferred to end). "
        "Stops on first error by default. "
        "Checks for human stop signals between operations. "
        "BATCH VARIABLES: $last = paragraph_index from previous step "
        "(works for BOTH insert AND edit ops like set_paragraph_text), "
        "$last+N / $last-N = offset, "
        "$last.bookmark = _mcp_ bookmark from previous step "
        "(available after insert ops or edit of bookmarked paragraphs), "
        "$step.N = paragraph_index from step N, "
        "$step.N.bookmark = bookmark from step N. "
        "Variables resolve to integers in numeric fields, "
        "strings in text fields (e.g. locator: 'paragraph:$last+1'). "
        "LOCATORS: Use 'heading_text:Some Title' for resilient "
        "addressing by heading text (survives bookmark loss). "
        "Use 'bookmark:_mcp_xxx' for precise addressing (call "
        "get_document_tree first to refresh). "
        "Use 'paragraph:N' for absolute positioning. "
        "When Record Changes is active, edits are attributed to "
        "'MCP' author. revision_comment adds a note visible in "
        "Manage Track Changes. "
        "Cannot call execute_batch recursively."
    )
    parameters = {
        "type": "object",
        "properties": {
            "operations": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "tool": {
                            "type": "string",
                            "description": "Tool name to execute",
                        },
                        "args": {
                            "type": "object",
                            "description": (
                                "Tool arguments. Use $last, $last+1, "
                                "$step.N for paragraph chaining."
                            ),
                        },
                        "revision_comment": {
                            "type": "string",
                            "description": (
                                "Comment attached to tracked changes "
                                "created by this operation (visible in "
                                "Manage Track Changes dialog). Only "
                                "effective when Record Changes is on."
                            ),
                        },
                    },
                    "required": ["tool"],
                },
                "description": (
                    "List of {tool, args} to execute sequentially"
                ),
            },
            "stop_on_error": {
                "type": "boolean",
                "description": (
                    "Halt on first failed operation (default: true)"
                ),
            },
            "check_conditions": {
                "type": "boolean",
                "description": (
                    "Check human stop conditions between ops "
                    "(default: true)"
                ),
            },
            "follow": {
                "type": "string",
                "description": (
                    "Scroll the view to follow edits. "
                    "'off' = no scroll (default), "
                    "'each' = scroll after every operation, "
                    "'end' = scroll after the last operation only"
                ),
            },
            "revision_comment": {
                "type": "string",
                "description": (
                    "Default revision comment for all operations. "
                    "Per-operation revision_comment overrides this. "
                    "Only effective when Record Changes is on."
                ),
            },
        },
        "required": ["operations"],
    }

    def execute(self, operations, stop_on_error=True,
                check_conditions=True,
                follow="off",
                revision_comment=None, **_):
        from mcp_server import get_mcp_server
        server = get_mcp_server()

        if not operations:
            return {"success": False, "error": "No operations provided"}
        if len(operations) > 50:
            return {"success": False,
                    "error": "Maximum 50 operations per batch"}

        # ── Pre-flight validation (no UNO, pure Python) ──
        validation_errors = []
        for i, op in enumerate(operations):
            tool_name = op.get("tool", "")
            args = op.get("args") or {}

            if tool_name == "execute_batch":
                validation_errors.append({
                    "step": i + 1, "tool": tool_name,
                    "error": "Recursive execute_batch not allowed"})
                continue

            tool = server.tools.get(tool_name)
            if tool is None:
                validation_errors.append({
                    "step": i + 1, "tool": tool_name,
                    "error": "Unknown tool: %s" % tool_name})
                continue

            # Skip validation for args with $vars (can't resolve yet)
            has_vars = any(
                isinstance(v, str) and '$' in v
                for v in args.values()
            ) if isinstance(args, dict) else False
            if not has_vars:
                ok, msg = tool.validate(**args)
                if not ok:
                    validation_errors.append({
                        "step": i + 1, "tool": tool_name,
                        "error": msg})

        if validation_errors:
            return {
                "success": False,
                "error": "Validation failed — nothing was executed",
                "validation_errors": validation_errors,
                "total": len(operations),
            }

        # ── Enter batch mode (suppress caches/indexing) ──
        self.services.batch_mode = True

        # ── Execute ──
        results = []
        stopped = False
        stop_reason = None
        stop_details = None
        last_result = None
        batch_vars = {}  # $last, $step.N

        for i, op in enumerate(operations):
            tool_name = op.get("tool", "")
            args = op.get("args") or {}

            # Resolve batch variables in args
            if batch_vars:
                args = _resolve_batch_vars(args, batch_vars)

            # Pre-flight stop condition check (skip before first op)
            if check_conditions and i > 0:
                should_stop, reason, details = _scan_stop_conditions(
                    self.services)
                if should_stop:
                    stopped = True
                    stop_reason = reason
                    stop_details = details
                    logger.info("Batch stopped at step %d/%d: %s",
                                i + 1, len(operations), reason)
                    break

            # Execute the tool
            rev_comment = op.get("revision_comment") or revision_comment
            result = server.execute_tool_sync(
                tool_name, args, revision_comment=rev_comment)
            step_ok = (isinstance(result, dict)
                       and result.get("success", True))
            results.append({
                "step": i + 1,
                "tool": tool_name,
                "success": step_ok,
                "result": result,
            })
            last_result = result

            # Update batch variables from result
            if step_ok:
                pi, bm = _extract_step_info(result)
                if pi is not None:
                    batch_vars["$last"] = pi
                    batch_vars["$step.%d" % (i + 1)] = pi
                if bm:
                    batch_vars["$last.bookmark"] = bm
                    batch_vars["$step.%d.bookmark" % (i + 1)] = bm

            # Follow: scroll after each operation
            if follow == "each" and step_ok:
                _follow_result(self.services, result)

            # Stop on error
            if stop_on_error and not step_ok:
                stopped = True
                stop_reason = "Tool '%s' failed" % tool_name
                if isinstance(result, dict):
                    stop_details = {
                        "error": result.get("error", "unknown")}
                break

            # Yield — brief pause to let other threads run
            if i < len(operations) - 1:
                time.sleep(0.01)

        # Follow: scroll after last operation (discrete mode)
        if follow == "end" and last_result and not stopped:
            _follow_result(self.services, last_result)

        # ── Exit batch mode — single invalidate + prewarm ──
        self.services.batch_mode = False
        try:
            self.services.writer.invalidate_caches()
        except Exception:
            pass
        try:
            self.services.writer.index.prewarm()
        except Exception:
            pass

        all_ok = all(r["success"] for r in results) and not stopped
        resp = {
            "success": all_ok,
            "completed": len(results),
            "total": len(operations),
            "stopped": stopped,
            "results": results,
        }
        if batch_vars:
            resp["batch_vars"] = batch_vars
        if stop_reason:
            resp["stop_reason"] = stop_reason
        if stop_details:
            resp["stop_details"] = stop_details
        return resp


class CheckStopConditions(McpTool):
    name = "check_stop_conditions"
    description = (
        "Check if human stop signals are present in the document. "
        "Scans for STOP/CANCEL comments and checks workflow dashboard "
        "phase. Call between tool calls to respect human control. "
        "Returns should_stop=true if the agent should halt."
    )
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": (
                    "Document path (optional, uses active document)"
                ),
            },
        },
    }

    def execute(self, file_path=None, **_):
        should_stop, reason, _ = _scan_stop_conditions(
            self.services, file_path)

        checks = {
            "stop_comments": [],
            "workflow_phase": None,
            "workflow_blocked": False,
        }

        # Gather details for transparency
        try:
            comments_result = self.services.comments.list_comments(
                file_path=file_path)
            for c in comments_result.get("comments", []):
                content = c.get("content", "").strip()
                first_word = (content.upper().split()[0]
                              if content else "")
                if first_word in _STOP_KEYWORDS:
                    checks["stop_comments"].append({
                        "author": c.get("author", ""),
                        "content": content,
                    })
        except Exception:
            pass

        try:
            wf = self.services.comments.get_workflow_status(
                file_path=file_path)
            if wf.get("found"):
                phase = wf.get("status", {}).get("Phase", "")
                checks["workflow_phase"] = phase
                checks["workflow_blocked"] = (
                    phase.strip().lower() in _BLOCKED_PHASES)
        except Exception:
            pass

        resp = {"should_stop": should_stop, "checks": checks}
        if reason:
            resp["reason"] = reason
        return resp
