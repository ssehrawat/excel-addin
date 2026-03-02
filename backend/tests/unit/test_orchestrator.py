"""Unit tests for pure helper methods on ``LangGraphOrchestrator``.

These tests exercise the data-transformation helpers used by the ReAct loop
(JSON extraction, table rendering, MongoDB extended-JSON handling, etc.)
without performing any I/O or instantiating LLM providers.
"""

from __future__ import annotations

import json

import pytest

from app.orchestrator import LangGraphOrchestrator
from app.mcp import ToolInvocationResult
from app.schemas import (
    ChatMessage,
    ChatRequest,
    MessageKind,
    MessageRole,
    WorkbookToolResult,
)


@pytest.fixture
def orch() -> LangGraphOrchestrator:
    """Return an orchestrator with no MCP service (Excel-only mode)."""
    return LangGraphOrchestrator(mcp_service=None)


# ---- _parse_mcp_tool_name ----

class TestParseMcpToolName:
    """Parsing namespaced MCP tool names."""

    def test_valid_split(self, orch):
        server_id, tool_name = orch._parse_mcp_tool_name("mcp__srv1__find")
        assert server_id == "srv1"
        assert tool_name == "find"

    def test_tool_with_underscores(self, orch):
        server_id, tool_name = orch._parse_mcp_tool_name("mcp__abc__find_all_items")
        assert server_id == "abc"
        assert tool_name == "find_all_items"

    def test_missing_prefix_raises(self, orch):
        with pytest.raises(ValueError, match="Expected format"):
            orch._parse_mcp_tool_name("get_xl_range_as_csv")

    def test_too_few_parts(self, orch):
        with pytest.raises(ValueError):
            orch._parse_mcp_tool_name("mcp__only")

    def test_empty_server_raises(self, orch):
        with pytest.raises(ValueError):
            orch._parse_mcp_tool_name("mcp____tool")

    def test_empty_tool_raises(self, orch):
        with pytest.raises(ValueError):
            orch._parse_mcp_tool_name("mcp__srv__")


# ---- _stringify_value ----

class TestStringifyValue:
    """Value stringification with MongoDB extended JSON support."""

    def test_plain_string(self, orch):
        assert orch._stringify_value("hello") == "hello"

    def test_integer(self, orch):
        assert orch._stringify_value(42) == "42"

    def test_date(self, orch):
        assert orch._stringify_value({"$date": "2024-01-01T00:00:00Z"}) == "2024-01-01T00:00:00Z"

    def test_oid(self, orch):
        assert orch._stringify_value({"$oid": "abc123"}) == "abc123"

    def test_number_double(self, orch):
        assert orch._stringify_value({"$numberDouble": "3.14"}) == "3.14"

    def test_list(self, orch):
        result = orch._stringify_value(["a", "b", "c"])
        assert result == "a, b, c"

    def test_generic_dict(self, orch):
        result = orch._stringify_value({"key": "val"})
        parsed = json.loads(result)
        assert parsed == {"key": "val"}


# ---- _render_table ----

class TestRenderTable:
    """Markdown table rendering from row dicts."""

    def test_basic_table(self, orch):
        rows = [{"name": "Alice", "age": "30"}, {"name": "Bob", "age": "25"}]
        result = orch._render_table(rows)
        assert "| name | age |" in result
        assert "| Alice | 30 |" in result
        assert "| Bob | 25 |" in result

    def test_empty_rows(self, orch):
        assert orch._render_table([]) == ""

    def test_long_cell_truncated(self, orch):
        rows = [{"data": "x" * 300}]
        result = orch._render_table(rows)
        # Cell should be truncated to 200 chars
        lines = result.strip().split("\n")
        data_line = lines[-1]
        # The value inside the pipe should be truncated
        assert len(data_line) < 300

    def test_column_order_from_rows(self, orch):
        rows = [
            {"b": "1", "a": "2"},
            {"a": "3", "c": "4"},
        ]
        result = orch._render_table(rows)
        header = result.split("\n")[0]
        # 'b' should come before 'a' because it appears first in the first row
        assert header.index("b") < header.index("a")


# ---- _find_balanced_segment ----

class TestFindBalancedSegment:
    """Balanced JSON segment extraction."""

    def test_simple_object(self, orch):
        text = '{"key": "value"}'
        assert orch._find_balanced_segment(text, "{", "}") == text

    def test_nested(self, orch):
        text = '{"a": {"b": 1}}'
        assert orch._find_balanced_segment(text, "{", "}") == text

    def test_array(self, orch):
        text = '[1, 2, [3, 4]]'
        assert orch._find_balanced_segment(text, "[", "]") == text

    def test_escaped_quotes(self, orch):
        text = '{"key": "val\\"ue"}'
        result = orch._find_balanced_segment(text, "{", "}")
        assert result == text

    def test_no_match_returns_none(self, orch):
        assert orch._find_balanced_segment("hello", "{", "}") is None

    def test_unclosed_returns_none(self, orch):
        assert orch._find_balanced_segment('{"key": "val', "{", "}") is None


# ---- _extract_json_from_text ----

class TestExtractJsonFromText:
    """JSON extraction from potentially wrapped text."""

    def test_plain_object(self, orch):
        text = '{"answer": "hi"}'
        result = orch._extract_json_from_text(text)
        assert result is not None
        assert json.loads(result)["answer"] == "hi"

    def test_untrusted_user_data_tags(self, orch):
        text = '<untrusted-user-data type="json">[{"id": 1}]</untrusted-user-data>'
        result = orch._extract_json_from_text(text)
        assert result is not None
        assert json.loads(result) == [{"id": 1}]

    def test_embedded_in_prose(self, orch):
        text = 'Here is the data: {"x": 1} and more text.'
        result = orch._extract_json_from_text(text)
        assert result is not None
        assert json.loads(result) == {"x": 1}

    def test_array(self, orch):
        text = "[1, 2, 3]"
        result = orch._extract_json_from_text(text)
        assert result is not None
        assert json.loads(result) == [1, 2, 3]

    def test_no_json_returns_none(self, orch):
        assert orch._extract_json_from_text("no json here") is None


# ---- _extract_rows ----

class TestExtractRows:
    """Row extraction from MCP tool responses."""

    def test_structured_content_list(self, orch):
        response = {"structuredContent": [{"name": "Alice"}, {"name": "Bob"}]}
        rows = orch._extract_rows(response)
        assert len(rows) == 2
        assert rows[0]["name"] == "Alice"

    def test_content_array_with_json(self, orch):
        response = {
            "content": [
                {"text": '[{"id": 1}, {"id": 2}]'}
            ]
        }
        rows = orch._extract_rows(response)
        assert len(rows) == 2

    def test_untrusted_tags(self, orch):
        response = {
            "content": [
                {"text": '<untrusted-user-data>[{"x": 1}]</untrusted-user-data>'}
            ]
        }
        rows = orch._extract_rows(response)
        assert len(rows) == 1

    def test_empty_returns_empty(self, orch):
        assert orch._extract_rows({}) == []


# ---- _summarize_workbook_tool_result ----

class TestSummarizeWorkbookToolResult:
    """Workbook tool result formatting."""

    def test_success(self, orch):
        result = WorkbookToolResult(id="1", tool="get_xl_range_as_csv", result="a,b\n1,2")
        summary = orch._summarize_workbook_tool_result(result)
        assert "get_xl_range_as_csv" in summary
        assert "a,b" in summary

    def test_error_field(self, orch):
        result = WorkbookToolResult(id="1", tool="test", result=None, error="Range not found")
        summary = orch._summarize_workbook_tool_result(result)
        assert "error" in summary.lower()
        assert "Range not found" in summary

    def test_truncates_at_50000(self, orch):
        result = WorkbookToolResult(id="1", tool="test", result="x" * 60000)
        summary = orch._summarize_workbook_tool_result(result)
        assert len(summary) < 55000


# ---- _augment_request ----

class TestAugmentRequest:
    """Request augmentation with context messages."""

    def test_no_context_returns_same(self, orch, minimal_chat_request):
        result = orch._augment_request(minimal_chat_request, [])
        assert result is minimal_chat_request

    def test_appends_context_messages(self, orch, minimal_chat_request):
        ctx = [orch._message(MessageKind.CONTEXT, "extra context")]
        result = orch._augment_request(minimal_chat_request, ctx)
        assert len(result.messages) == len(minimal_chat_request.messages) + 1

    def test_does_not_mutate_original(self, orch, minimal_chat_request):
        original_count = len(minimal_chat_request.messages)
        ctx = [orch._message(MessageKind.CONTEXT, "extra")]
        orch._augment_request(minimal_chat_request, ctx)
        assert len(minimal_chat_request.messages) == original_count
