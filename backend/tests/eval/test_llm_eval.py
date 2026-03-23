"""LLM eval harness — golden-file fixture tests.

Each test class loads a fixture JSON file that represents a realistic LLM
response, feeds it through the same parsing/building pipeline used in
production, and asserts that the output is correct.  This catches
regressions in the data-transformation layer when the system prompt or
response format changes.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from app.providers import (
    _build_tool_call,
    assemble_messages_from_payload,
    build_cell_updates,
    build_chart_inserts,
    build_format_updates,
    build_pivot_table_inserts,
    parse_structured_response,
)
from app.schemas import MessageKind

FIXTURES_DIR = Path(__file__).parent / "fixtures"


def _load_fixture(name: str):
    """Load and return a fixture by filename."""
    path = FIXTURES_DIR / name
    with open(path, encoding="utf-8") as f:
        return json.load(f)


class TestDirectAnswer:
    """Fixture: direct_answer.json — simple answer with no mutations."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("direct_answer.json")

    def test_parsed_as_dict(self, payload):
        parsed = parse_structured_response(payload)
        assert isinstance(parsed, dict)
        assert "answer" in parsed

    def test_message_assembly(self, payload):
        parsed = parse_structured_response(payload)
        msgs = assemble_messages_from_payload(parsed, "What is Q1 revenue?")
        assert len(msgs) == 1
        assert msgs[0].kind == MessageKind.FINAL
        assert "42,500" in msgs[0].content

    def test_no_mutations(self, payload):
        parsed = parse_structured_response(payload)
        assert build_cell_updates(parsed.get("cell_updates", [])) == []
        assert build_format_updates(parsed.get("format_updates", [])) == []
        assert build_chart_inserts(parsed.get("chart_inserts", [])) == []


class TestNeedsDataExcel:
    """Fixture: needs_data_excel.json — LLM requests Excel tool data."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("needs_data_excel.json")

    def test_needs_data_flag(self, payload):
        parsed = parse_structured_response(payload)
        assert parsed["needs_data"] is True

    def test_tool_call_extraction(self, payload):
        parsed = parse_structured_response(payload)
        tc = _build_tool_call(parsed["tool_call"])
        assert tc is not None
        assert tc.tool == "get_xl_range_as_csv"
        assert tc.args["sheetName"] == "Sales"

    def test_not_mcp_tool(self, payload):
        parsed = parse_structured_response(payload)
        tc = _build_tool_call(parsed["tool_call"])
        assert not tc.tool.startswith("mcp__")


class TestNeedsDataMcp:
    """Fixture: needs_data_mcp.json — LLM requests MCP tool data."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("needs_data_mcp.json")

    def test_mcp_tool_name(self, payload):
        parsed = parse_structured_response(payload)
        tc = _build_tool_call(parsed["tool_call"])
        assert tc is not None
        assert tc.tool.startswith("mcp__")
        assert "find_customers" in tc.tool

    def test_args_preserved(self, payload):
        parsed = parse_structured_response(payload)
        tc = _build_tool_call(parsed["tool_call"])
        assert tc.args["query"] == "region = EMEA"
        assert tc.args["limit"] == 50


class TestCellUpdates:
    """Fixture: cell_updates.json — answer with formula cell writes."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("cell_updates.json")

    def test_cell_updates_count(self, payload):
        parsed = parse_structured_response(payload)
        updates = build_cell_updates(parsed["cell_updates"])
        assert len(updates) == 2

    def test_formula_in_values(self, payload):
        parsed = parse_structured_response(payload)
        updates = build_cell_updates(parsed["cell_updates"])
        assert updates[0].values[0][0].startswith("=SUM")

    def test_address_includes_sheet(self, payload):
        parsed = parse_structured_response(payload)
        updates = build_cell_updates(parsed["cell_updates"])
        assert "Sheet1" in updates[0].address


class TestChartInserts:
    """Fixture: chart_inserts.json — scatter chart with axis titles."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("chart_inserts.json")

    def test_chart_type_normalized(self, payload):
        parsed = parse_structured_response(payload)
        inserts = build_chart_inserts(parsed["chart_inserts"])
        assert len(inserts) == 1
        assert inserts[0].chart_type == "XYScatter"

    def test_axis_titles(self, payload):
        parsed = parse_structured_response(payload)
        inserts = build_chart_inserts(parsed["chart_inserts"])
        assert inserts[0].x_axis_title == "Units Sold"
        assert inserts[0].y_axis_title == "Revenue ($)"

    def test_series_by(self, payload):
        parsed = parse_structured_response(payload)
        inserts = build_chart_inserts(parsed["chart_inserts"])
        assert inserts[0].series_by.value == "columns"


class TestMalformedResponse:
    """Fixture: malformed_response.json — plain text (non-JSON)."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("malformed_response.json")

    def test_graceful_fallback(self, payload):
        parsed = parse_structured_response(payload)
        assert isinstance(parsed, dict)
        # The plain text should end up in the answer field
        assert "answer" in parsed
        assert "100" in str(parsed["answer"])

    def test_empty_mutations(self, payload):
        parsed = parse_structured_response(payload)
        assert build_cell_updates(parsed.get("cell_updates", [])) == []


class TestClarificationResponse:
    """Fixture: ambiguous_needs_clarification.json — clarification question."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("ambiguous_needs_clarification.json")

    def test_answer_is_question(self, payload):
        parsed = parse_structured_response(payload)
        msgs = assemble_messages_from_payload(parsed, "Analyze my data")
        assert len(msgs) == 1
        assert "which sheet" in msgs[0].content.lower()

    def test_no_tool_call(self, payload):
        parsed = parse_structured_response(payload)
        assert not parsed.get("needs_data")


class TestCamelCaseVariants:
    """Fixture: camelcase_variants.json — camelCase keys throughout."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("camelcase_variants.json")

    def test_format_updates_from_camelcase(self, payload):
        parsed = parse_structured_response(payload)
        updates = build_format_updates(parsed["format_updates"])
        assert len(updates) == 1
        assert updates[0].fill_color == "#FFFF00"

    def test_chart_type_alias_resolved(self, payload):
        parsed = parse_structured_response(payload)
        inserts = build_chart_inserts(parsed["chart_inserts"])
        assert inserts[0].chart_type == "XYScatter"

    def test_series_by_from_camelcase(self, payload):
        parsed = parse_structured_response(payload)
        inserts = build_chart_inserts(parsed["chart_inserts"])
        assert inserts[0].series_by.value == "rows"


class TestPivotTableInserts:
    """Fixture: pivot_table_inserts.json — pivot table with full hierarchy config."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("pivot_table_inserts.json")

    def test_parsed(self, payload):
        parsed = parse_structured_response(payload)
        inserts = build_pivot_table_inserts(parsed["pivot_table_inserts"])
        assert len(inserts) == 1
        assert inserts[0].name == "SalesPivot"

    def test_hierarchies(self, payload):
        parsed = parse_structured_response(payload)
        inserts = build_pivot_table_inserts(parsed["pivot_table_inserts"])
        assert inserts[0].rows == ["Region"]
        assert inserts[0].columns == ["Quarter"]
        assert inserts[0].filters == ["Category"]

    def test_values_and_aggregation(self, payload):
        parsed = parse_structured_response(payload)
        inserts = build_pivot_table_inserts(parsed["pivot_table_inserts"])
        values = inserts[0].values
        assert len(values) == 2
        assert values[0].name == "Revenue"
        assert values[0].summarize_by.value == "sum"
        assert values[1].name == "Quantity"
        assert values[1].summarize_by.value == "average"

    def test_source_and_destination(self, payload):
        parsed = parse_structured_response(payload)
        inserts = build_pivot_table_inserts(parsed["pivot_table_inserts"])
        assert inserts[0].source_address == "A1:D100"
        assert inserts[0].source_worksheet == "Sales"
        assert inserts[0].destination_address == "F1"
        assert inserts[0].destination_worksheet == "Summary"


class TestCustomFunctionResponse:
    """Fixture: custom_function_response.json — minimal =ASKAI single-shot response."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("custom_function_response.json")

    def test_parsed_as_dict(self, payload):
        parsed = parse_structured_response(payload)
        assert isinstance(parsed, dict)

    def test_has_answer(self, payload):
        parsed = parse_structured_response(payload)
        assert "answer" in parsed
        assert len(parsed["answer"]) > 0

    def test_no_tool_call(self, payload):
        parsed = parse_structured_response(payload)
        assert not parsed.get("needs_data")

    def test_minimal_context_accepted(self, payload):
        """Ensures the response parses without workbook_metadata, user_context, etc."""
        parsed = parse_structured_response(payload)
        msgs = assemble_messages_from_payload(parsed, "Summarize this data")
        assert len(msgs) == 1
        assert msgs[0].kind == MessageKind.FINAL


class TestTranscribeRoundtrip:
    """Fixture: transcribe_response.json — Whisper API response shape."""

    @pytest.fixture
    def payload(self):
        return _load_fixture("transcribe_response.json")

    def test_has_text_field(self, payload):
        assert "text" in payload
        assert isinstance(payload["text"], str)

    def test_text_non_empty(self, payload):
        assert len(payload["text"]) > 0

    def test_response_shape(self, payload):
        """The transcribe response should contain only the 'text' key."""
        assert set(payload.keys()) == {"text"}
