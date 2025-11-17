# Bug Fix: MongoDB Data Getting Truncated and Dropped from Context

## Issue Summary

When querying MongoDB via MCP tools (e.g., "Give me 1st 10 records from equities_daily collection"), the data was being returned from MongoDB but was not appearing in the chat or Excel cell updates. The LLM responded with "the actual data values are not present in this chat session" even though the tool successfully retrieved the data.

## Root Cause

The issue was in `backend/app/orchestrator.py`, specifically in the `_summarize_tool_result` method:

```python
# BEFORE (Line 581-582):
if len(serialized) > 600:
    serialized = serialized[:600] + "…"
```

The 600-character truncation limit was far too small for database query results. When MongoDB returned 10 records, the response was being truncated before being passed to the LLM as context. This caused the LLM to not receive the actual data, leading it to report that the data was not available.

### Data Flow

1. MongoDB MCP tool returns data wrapped in `<untrusted-user-data>` tags
2. `_summarize_tool_result()` processes the response:
   - Tries to extract rows and format as a table (preferred path)
   - Falls back to JSON serialization if extraction fails
   - **[BUG]** Truncated to 600 characters
3. Summary is added as a CONTEXT message
4. Context messages are sent to the LLM via `build_prompt_payload()`
5. LLM uses this context to generate answers and cell updates

## Fix Applied

### 1. Increased Truncation Limit (Primary Fix)

**File**: `backend/app/orchestrator.py`, line 582

```python
# AFTER:
# Increase limit to 50000 characters to avoid truncating database results
if len(serialized) > 50000:
    serialized = serialized[:50000] + "…"
```

This ensures that database query results with multiple records can pass through to the LLM without being cut off.

### 2. Enhanced Logging for Debugging

Added comprehensive debug logging throughout the data extraction pipeline:

- `_extract_rows()`: Logs structuredContent and content array processing
- `_extract_json_from_text()`: Logs untrusted-user-data tag extraction
- `_render_table()`: Logs table dimensions

This helps diagnose future issues with data extraction and formatting.

### 3. Improved JSON Extraction

Enhanced `_extract_json_from_text()` to better handle:
- Untrusted-user-data tags with UUIDs (e.g., `<untrusted-user-data-b67c12f1-...>`)
- Better error handling and fallback logic
- Clearer logging of extraction steps

### 4. Table Cell Value Limiting

Added per-cell truncation in `_render_table()`:

```python
# Limit each cell value to 200 chars to prevent extremely wide tables
values = [str(row.get(header, ""))[:200] for header in headers]
```

This prevents individual cell values from making tables unwieldy while still preserving the overall dataset.

### 5. Better Single-Object Handling

Enhanced `_extract_rows()` to handle both list and single-object JSON responses:

```python
elif isinstance(data, dict):
    logger.debug("Parsed single JSON object")
    rows.append(self._flatten_document(data))
```

## Testing Recommendations

To verify the fix works:

1. Restart the backend server to load the updated code
2. Re-run the query: "Give me 1st 10 records from equities_daily collection in mktdata database"
3. Check the logs for debug messages showing:
   - "Extracted X rows from tool response"
   - "Rendered table with X rows and Y columns"
4. Verify the LLM response includes:
   - The actual data in the chat
   - Cell updates with the records
   - No message about missing data

## Expected Behavior After Fix

1. MongoDB tool returns 10 records
2. Orchestrator extracts and formats them as a markdown table
3. Full table (up to 50,000 chars) is passed to LLM as context
4. LLM sees the data and:
   - Shows it in chat as a formatted table
   - Generates cell_updates to populate Excel
   - Creates appropriate formatting and charts if requested

## Related Files Modified

- `backend/app/orchestrator.py`: Main fixes applied

## Notes

- The original 600-character limit was likely set conservatively to avoid token limits, but modern LLMs can handle much larger contexts
- The 50,000 character limit is reasonable for most database queries while still preventing extreme cases
- Individual cell values in tables are limited to 200 chars to maintain readability
- The extraction logic properly handles MongoDB's JSON format with BSON type wrappers like `$date`, `$numberDouble`, `$oid`

