"""Unit tests for pure helper functions in ``app.providers``.

These functions transform raw LLM output (JSON dicts, strings, None) into
validated Pydantic models.  They are the highest-ROI tests because they
exercise the core data-transformation layer without any I/O.
"""

from __future__ import annotations

import json

import pytest

from app.providers import (
    CHART_TYPE_ALIASES,
    MockProvider,
    _build_tool_call,
    _normalize_chart_type,
    _normalize_identifier,
    assemble_messages_from_payload,
    build_cell_updates,
    build_chart_inserts,
    build_format_updates,
    build_pivot_table_inserts,
    build_prompt_payload,
    build_system_prompt,
    parse_structured_response,
    MCPToolEntry,
)
from app.schemas import (
    CellSelection,
    ChatMessage,
    ChatRequest,
    MessageKind,
    MessageRole,
)


# ---- parse_structured_response ----

class TestParseStructuredResponse:
    """Parsing raw LLM output into a dict."""

    def test_dict_passthrough(self):
        raw = {"answer": "hi"}
        assert parse_structured_response(raw) == raw

    def test_valid_json_string(self):
        raw = '{"answer": "hi", "cell_updates": []}'
        result = parse_structured_response(raw)
        assert result["answer"] == "hi"

    def test_invalid_json_fallback(self):
        raw = "This is not JSON at all."
        result = parse_structured_response(raw)
        assert result["answer"] == raw
        assert result["cell_updates"] == []

    def test_none_input(self):
        result = parse_structured_response(None)
        assert result["answer"] == "No answer produced."

    def test_needs_data_preserved(self):
        raw = {"needs_data": True, "tool_call": {"tool": "get_xl_range_as_csv", "args": {}}}
        result = parse_structured_response(raw)
        assert result["needs_data"] is True

    def test_empty_string(self):
        result = parse_structured_response("")
        assert result["answer"] == ""

    def test_integer_input(self):
        result = parse_structured_response(42)
        assert result["answer"] == "No answer produced."

    def test_json_with_extra_keys(self):
        raw = '{"answer": "x", "custom_key": 1}'
        result = parse_structured_response(raw)
        assert result["answer"] == "x"
        assert result["custom_key"] == 1


# ---- build_cell_updates ----

class TestBuildCellUpdates:
    """Converting raw LLM cell_updates into validated CellUpdate models."""

    def test_valid_replace(self):
        raw = [{"address": "A1", "values": [["hello"]], "mode": "replace"}]
        updates = build_cell_updates(raw)
        assert len(updates) == 1
        assert updates[0].address == "A1"
        assert updates[0].mode.value == "replace"

    def test_valid_append(self):
        raw = [{"address": "A1", "values": [["x"]], "mode": "append"}]
        updates = build_cell_updates(raw)
        assert updates[0].mode.value == "append"

    def test_flat_list_auto_wrap(self):
        """A flat list of values should be auto-wrapped into a 2D array."""
        raw = [{"address": "A1", "values": ["a", "b", "c"]}]
        updates = build_cell_updates(raw)
        assert updates[0].values == [["a", "b", "c"]]

    def test_missing_address_skipped(self):
        raw = [{"values": [["x"]]}]
        assert build_cell_updates(raw) == []

    def test_missing_values_skipped(self):
        raw = [{"address": "A1"}]
        assert build_cell_updates(raw) == []

    def test_invalid_mode_defaults_to_replace(self):
        raw = [{"address": "A1", "values": [["x"]], "mode": "upsert"}]
        updates = build_cell_updates(raw)
        assert updates[0].mode.value == "replace"

    def test_non_dict_skipped(self):
        raw = ["not a dict", 42, None]
        assert build_cell_updates(raw) == []

    def test_none_input(self):
        assert build_cell_updates(None) == []

    def test_multiple_valid(self):
        raw = [
            {"address": "A1", "values": [["a"]]},
            {"address": "B2", "values": [[1, 2]]},
        ]
        assert len(build_cell_updates(raw)) == 2

    def test_worksheet_preserved(self):
        raw = [{"address": "A1", "values": [["x"]], "worksheet": "Sheet2"}]
        updates = build_cell_updates(raw)
        assert updates[0].worksheet == "Sheet2"


# ---- build_chart_inserts ----

class TestBuildChartInserts:
    """Converting raw LLM chart_inserts into ChartInsert models."""

    def test_scatter_alias_normalized(self):
        raw = [{"chart_type": "scatter", "source_address": "A1:B10"}]
        inserts = build_chart_inserts(raw)
        assert len(inserts) == 1
        assert inserts[0].chart_type == "XYScatter"

    def test_camelcase_keys(self):
        raw = [{"chartType": "column", "sourceAddress": "A1:C5"}]
        inserts = build_chart_inserts(raw)
        assert inserts[0].chart_type == "ColumnClustered"
        assert inserts[0].source_address == "A1:C5"

    def test_unknown_type_skipped(self):
        raw = [{"chart_type": "", "source_address": "A1:B5"}]
        inserts = build_chart_inserts(raw)
        assert len(inserts) == 0

    def test_missing_source_skipped(self):
        raw = [{"chart_type": "line"}]
        inserts = build_chart_inserts(raw)
        assert len(inserts) == 0

    def test_axis_titles(self):
        raw = [{
            "chart_type": "line",
            "source_address": "A1:B5",
            "x_axis_title": "X",
            "y_axis_title": "Y",
        }]
        inserts = build_chart_inserts(raw)
        assert inserts[0].x_axis_title == "X"
        assert inserts[0].y_axis_title == "Y"

    def test_series_by(self):
        raw = [{"chart_type": "bar", "source_address": "A1:C5", "series_by": "columns"}]
        inserts = build_chart_inserts(raw)
        assert inserts[0].series_by.value == "columns"

    def test_camelcase_axis_titles(self):
        raw = [{
            "chartType": "line",
            "sourceAddress": "A1:B5",
            "xAxisTitle": "Time",
            "yAxisTitle": "Revenue",
        }]
        inserts = build_chart_inserts(raw)
        assert inserts[0].x_axis_title == "Time"

    def test_non_dict_skipped(self):
        raw = ["not a dict"]
        assert build_chart_inserts(raw) == []

    def test_none_input(self):
        assert build_chart_inserts(None) == []


# ---- build_format_updates ----

class TestBuildFormatUpdates:
    """Converting raw LLM format_updates into FormatUpdate models."""

    def test_fill_color(self):
        raw = [{"address": "A1", "fill_color": "#FF0000"}]
        updates = build_format_updates(raw)
        assert len(updates) == 1
        assert updates[0].fill_color == "#FF0000"

    def test_camelcase_aliases(self):
        raw = [{"address": "A1", "fillColor": "#00FF00", "fontColor": "#0000FF"}]
        updates = build_format_updates(raw)
        assert updates[0].fill_color == "#00FF00"
        assert updates[0].font_color == "#0000FF"

    def test_bold_italic(self):
        raw = [{"address": "A1", "bold": True, "italic": False}]
        updates = build_format_updates(raw)
        assert updates[0].bold is True
        assert updates[0].italic is False

    def test_missing_address_skipped(self):
        raw = [{"fill_color": "#FF0000"}]
        assert build_format_updates(raw) == []

    def test_non_bool_bold_ignored(self):
        raw = [{"address": "A1", "bold": "yes"}]
        updates = build_format_updates(raw)
        assert updates[0].bold is None

    def test_none_input(self):
        assert build_format_updates(None) == []

    def test_number_format(self):
        raw = [{"address": "A1", "numberFormat": "#,##0.00"}]
        updates = build_format_updates(raw)
        assert updates[0].number_format == "#,##0.00"

    def test_border_properties(self):
        raw = [{"address": "A1:D10", "borderColor": "#000000",
                "borderStyle": "Continuous", "borderWeight": "Thin"}]
        updates = build_format_updates(raw)
        assert updates[0].border_color == "#000000"
        assert updates[0].border_style == "Continuous"
        assert updates[0].border_weight == "Thin"

    def test_snake_case_border_keys(self):
        raw = [{"address": "A1", "border_color": "#333",
                "border_style": "Dash", "border_weight": "Medium"}]
        updates = build_format_updates(raw)
        assert updates[0].border_color == "#333"

    def test_all_fields_together(self):
        raw = [{"address": "Sheet1!A1:D1", "fillColor": "#4472C4",
                "fontColor": "#FFFFFF", "bold": True, "italic": False,
                "numberFormat": "0.00%", "borderColor": "#000000",
                "borderStyle": "Continuous", "borderWeight": "Thin"}]
        updates = build_format_updates(raw)
        assert len(updates) == 1
        u = updates[0]
        assert u.fill_color == "#4472C4"
        assert u.font_color == "#FFFFFF"
        assert u.bold is True
        assert u.italic is False
        assert u.number_format == "0.00%"
        assert u.border_color == "#000000"


# ---- _normalize_chart_type ----

class TestNormalizeChartType:
    """Chart type alias resolution and normalization."""

    def test_known_alias(self):
        assert _normalize_chart_type("scatter") == "XYScatter"

    def test_xl_prefix(self):
        assert _normalize_chart_type("xlScatter") == "XYScatter"

    def test_case_insensitive(self):
        assert _normalize_chart_type("SCATTER") == "XYScatter"

    def test_unknown_passthrough(self):
        result = _normalize_chart_type("SomeCustomType")
        assert result == "SomeCustomType"

    def test_non_string_returns_none(self):
        assert _normalize_chart_type(42) is None


# ---- _build_tool_call ----

class TestBuildToolCall:
    """Building WorkbookToolCall from raw LLM tool_call dict."""

    def test_valid_excel_tool(self):
        raw = {"tool": "get_xl_range_as_csv", "args": {"sheetName": "Sheet1"}}
        tc = _build_tool_call(raw)
        assert tc is not None
        assert tc.tool == "get_xl_range_as_csv"
        assert tc.args == {"sheetName": "Sheet1"}

    def test_mcp_tool(self):
        raw = {"tool": "mcp__srv1__find_customers", "args": {"query": "test"}}
        tc = _build_tool_call(raw)
        assert tc is not None
        assert tc.tool == "mcp__srv1__find_customers"

    def test_missing_tool_returns_none(self):
        assert _build_tool_call({"args": {}}) is None

    def test_empty_tool_returns_none(self):
        assert _build_tool_call({"tool": "", "args": {}}) is None

    def test_non_dict_returns_none(self):
        assert _build_tool_call("not a dict") is None

    def test_non_dict_args_default_to_empty(self):
        tc = _build_tool_call({"tool": "get_xl_range_as_csv", "args": "bad"})
        assert tc is not None
        assert tc.args == {}


# ---- assemble_messages_from_payload ----

class TestAssembleMessagesFromPayload:
    """Building ChatMessage objects from parsed LLM response dict."""

    def test_answer_present(self):
        payload = {"answer": "The total is 42."}
        msgs = assemble_messages_from_payload(payload, "What is the total?")
        assert len(msgs) == 1
        assert msgs[0].kind == MessageKind.FINAL
        assert msgs[0].content == "The total is 42."

    def test_empty_answer_fallback(self):
        payload = {"answer": ""}
        msgs = assemble_messages_from_payload(payload, "prompt")
        assert len(msgs) == 1
        assert "best effort" in msgs[0].content

    def test_whitespace_answer_fallback(self):
        payload = {"answer": "   "}
        msgs = assemble_messages_from_payload(payload, "prompt")
        assert "best effort" in msgs[0].content

    def test_missing_answer_fallback(self):
        payload = {"cell_updates": []}
        msgs = assemble_messages_from_payload(payload, "prompt")
        assert len(msgs) == 1
        assert msgs[0].kind == MessageKind.FINAL
        assert "best effort" in msgs[0].content

    def test_missing_answer_with_mutations_says_done(self):
        payload = {"pivot_table_inserts": [{"name": "PT1"}]}
        msgs = assemble_messages_from_payload(payload, "create a pivot table")
        assert len(msgs) == 1
        assert msgs[0].content == "Done."

    def test_always_final_kind(self):
        payload = {"answer": "hi"}
        msgs = assemble_messages_from_payload(payload, "prompt")
        assert all(m.kind == MessageKind.FINAL for m in msgs)


# ---- build_system_prompt ----

class TestBuildSystemPrompt:
    """Dynamic system prompt generation."""

    def test_no_mcp_tools_no_mcp_section(self):
        prompt = build_system_prompt([])
        assert "EXTERNAL MCP TOOLS" not in prompt
        assert "DECISION RULES" in prompt

    def test_with_mcp_tools_includes_namespaced(self):
        tool = MCPToolEntry(
            namespaced_name="mcp__srv1__find",
            server_id="srv1",
            server_name="TestServer",
            description="Find things",
            input_schema={"type": "object"},
        )
        prompt = build_system_prompt([tool])
        assert "EXTERNAL MCP TOOLS" in prompt
        assert "mcp__srv1__find" in prompt
        assert "TestServer" in prompt

    def test_decision_rules_present(self):
        prompt = build_system_prompt([])
        assert "DECISION RULES" in prompt

    def test_response_format_present(self):
        prompt = build_system_prompt([])
        assert "RESPONSE FORMAT" in prompt

    def test_pivot_table_rules_present(self):
        prompt = build_system_prompt([])
        assert "RULES FOR PIVOT TABLE INSERTS" in prompt
        assert "pivot_table_inserts" in prompt

    def test_format_update_schema_present(self):
        prompt = build_system_prompt([])
        assert "fillColor" in prompt
        assert "borderColor" in prompt
        assert "borderWeight" in prompt

    def test_chart_source_rule_present(self):
        prompt = build_system_prompt([])
        assert "CHART SOURCE RULE" in prompt

    def test_placement_rule_present(self):
        prompt = build_system_prompt([])
        assert "PLACEMENT RULE" in prompt

    def test_formula_rule_strengthened(self):
        prompt = build_system_prompt([])
        assert "FORMULA RULE" in prompt
        assert "=SUM" in prompt
        assert "cross-check column letters" in prompt

    def test_column_headers_in_context_description(self):
        prompt = build_system_prompt([])
        assert "columnHeaders" in prompt
        assert "column-letter markers" in prompt


# ---- build_prompt_payload ----

class TestBuildPromptPayload:
    """Building the user-turn JSON payload for the LLM."""

    def test_minimal_request(self):
        request = ChatRequest(
            prompt="hello",
            provider="mock",
            messages=[],
            selection=[],
        )
        payload = json.loads(build_prompt_payload(request))
        assert payload["user_prompt"] == "hello"

    def test_includes_metadata(self):
        from app.schemas import WorkbookMetadata, SheetMetadata
        request = ChatRequest(
            prompt="hi",
            provider="mock",
            messages=[],
            selection=[],
            workbook_metadata=WorkbookMetadata(
                success=True,
                file_name="test.xlsx",
                sheets_metadata=[],
                total_sheets=0,
            ),
        )
        payload = json.loads(build_prompt_payload(request))
        assert "workbook_metadata" in payload
        assert payload["workbook_metadata"]["fileName"] == "test.xlsx"

    def test_history_trimmed_to_last_6(self):
        msgs = [
            ChatMessage(
                id=f"m{i}",
                role=MessageRole.USER,
                kind=MessageKind.MESSAGE,
                content=f"msg-{i}",
                created_at="2024-01-01T00:00:00Z",
            )
            for i in range(10)
        ]
        request = ChatRequest(
            prompt="hi",
            provider="mock",
            messages=msgs,
            selection=[],
        )
        payload = json.loads(build_prompt_payload(request))
        history = payload["conversation_history"]
        # Should contain the last 6 messages
        assert "msg-4" in history
        assert "msg-9" in history
        assert "msg-0" not in history

    def test_context_messages_separated(self):
        msgs = [
            ChatMessage(
                id="m1",
                role=MessageRole.SYSTEM,
                kind=MessageKind.CONTEXT,
                content="context data",
                created_at="2024-01-01T00:00:00Z",
            ),
            ChatMessage(
                id="m2",
                role=MessageRole.USER,
                kind=MessageKind.MESSAGE,
                content="user message",
                created_at="2024-01-01T00:00:00Z",
            ),
        ]
        request = ChatRequest(
            prompt="hi",
            provider="mock",
            messages=msgs,
            selection=[],
        )
        payload = json.loads(build_prompt_payload(request))
        history = payload["conversation_history"]
        assert "tool_context" in history


# ---- MockProvider ----

class TestMockProvider:
    """MockProvider generates deterministic responses for testing."""

    @pytest.mark.asyncio
    async def test_basic_response(self, minimal_chat_request):
        provider = MockProvider()
        result = await provider.generate(minimal_chat_request)
        assert len(result.messages) >= 1
        assert any(m.kind == MessageKind.FINAL for m in result.messages)

    @pytest.mark.asyncio
    async def test_selection_triggers_cell_update(self):
        request = ChatRequest(
            prompt="sum it",
            provider="mock",
            messages=[],
            selection=[CellSelection(address="A1:A5", values=[[1], [2], [3]])],
        )
        result = await MockProvider().generate(request)
        assert len(result.cell_updates) >= 1
        assert result.cell_updates[0].address == "A1:A5"

    @pytest.mark.asyncio
    async def test_chart_keyword_triggers_chart_insert(self):
        request = ChatRequest(
            prompt="create a chart",
            provider="mock",
            messages=[],
            selection=[CellSelection(address="A1:B5", values=[[1, 2]])],
        )
        result = await MockProvider().generate(request)
        assert len(result.chart_inserts) >= 1

    @pytest.mark.asyncio
    async def test_color_keyword_triggers_format_update(self):
        request = ChatRequest(
            prompt="change the color",
            provider="mock",
            messages=[],
            selection=[CellSelection(address="A1", values=[["x"]])],
        )
        result = await MockProvider().generate(request)
        assert len(result.format_updates) >= 1

    @pytest.mark.asyncio
    async def test_no_selection_no_cell_update(self):
        request = ChatRequest(
            prompt="hello",
            provider="mock",
            messages=[],
            selection=[],
        )
        result = await MockProvider().generate(request)
        assert result.cell_updates == []

    @pytest.mark.asyncio
    async def test_pivot_keyword_triggers_pivot_insert(self):
        request = ChatRequest(
            prompt="create a pivot table",
            provider="mock",
            messages=[],
            selection=[CellSelection(address="A1:D20", values=[[1, 2, 3, 4]])],
        )
        result = await MockProvider().generate(request)
        assert len(result.pivot_table_inserts) >= 1
        assert result.pivot_table_inserts[0].name == "MockPivotTable"


# ---- _normalize_identifier ----

class TestNormalizeIdentifier:
    """Identifier normalization used by chart type resolution."""

    def test_strips_xl_prefix(self):
        assert _normalize_identifier("xlScatter") == "scatter"

    def test_removes_non_alphanum(self):
        assert _normalize_identifier("scatter-lines") == "scatterlines"

    def test_lowercases(self):
        assert _normalize_identifier("COLUMN") == "column"


# ---- build_pivot_table_inserts ----

class TestBuildPivotTableInserts:
    """Converting raw LLM pivot_table_inserts into PivotTableInsert models."""

    def test_valid_minimal(self):
        raw = [{"name": "PT1", "sourceAddress": "A1:D100"}]
        inserts = build_pivot_table_inserts(raw)
        assert len(inserts) == 1
        assert inserts[0].name == "PT1"
        assert inserts[0].source_address == "A1:D100"

    def test_camelcase_keys(self):
        raw = [{
            "name": "PT1",
            "sourceAddress": "A1:D50",
            "sourceWorksheet": "Data",
            "destinationAddress": "F1",
            "destinationWorksheet": "Summary",
        }]
        inserts = build_pivot_table_inserts(raw)
        assert inserts[0].source_worksheet == "Data"
        assert inserts[0].destination_address == "F1"
        assert inserts[0].destination_worksheet == "Summary"

    def test_missing_name_skipped(self):
        raw = [{"sourceAddress": "A1:D100"}]
        assert build_pivot_table_inserts(raw) == []

    def test_missing_source_skipped(self):
        raw = [{"name": "PT1"}]
        assert build_pivot_table_inserts(raw) == []

    def test_values_with_aggregation(self):
        raw = [{
            "name": "PT1",
            "sourceAddress": "A1:D100",
            "values": [{"name": "Revenue", "summarizeBy": "average"}],
        }]
        inserts = build_pivot_table_inserts(raw)
        assert len(inserts[0].values) == 1
        assert inserts[0].values[0].name == "Revenue"
        assert inserts[0].values[0].summarize_by.value == "average"

    def test_string_values_default_sum(self):
        raw = [{
            "name": "PT1",
            "sourceAddress": "A1:D100",
            "values": ["Revenue", "Quantity"],
        }]
        inserts = build_pivot_table_inserts(raw)
        assert len(inserts[0].values) == 2
        assert inserts[0].values[0].summarize_by.value == "sum"
        assert inserts[0].values[1].name == "Quantity"

    def test_avg_alias(self):
        raw = [{
            "name": "PT1",
            "sourceAddress": "A1:D100",
            "values": [{"name": "Price", "summarizeBy": "avg"}],
        }]
        inserts = build_pivot_table_inserts(raw)
        assert inserts[0].values[0].summarize_by.value == "average"

    def test_rows_columns_filters(self):
        raw = [{
            "name": "PT1",
            "sourceAddress": "A1:D100",
            "rows": ["Region"],
            "columns": ["Quarter"],
            "filters": ["Category"],
        }]
        inserts = build_pivot_table_inserts(raw)
        assert inserts[0].rows == ["Region"]
        assert inserts[0].columns == ["Quarter"]
        assert inserts[0].filters == ["Category"]

    def test_none_input(self):
        assert build_pivot_table_inserts(None) == []

    def test_non_dict_skipped(self):
        raw = ["not a dict", 42]
        assert build_pivot_table_inserts(raw) == []

    def test_destination_defaults(self):
        raw = [{"name": "PT1", "sourceAddress": "A1:D100"}]
        inserts = build_pivot_table_inserts(raw)
        assert inserts[0].destination_address is None
        assert inserts[0].destination_worksheet is None
