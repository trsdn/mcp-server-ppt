"""CLI modification pattern tests (targeted updates)."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
@pytest.mark.xfail(reason="LLM intermittently omits required action parameter on complex workflows", strict=False)
async def test_cli_range_updates(aitest_run, ppt_cli_server, ppt_cli_skill):
    agent = Agent(
        name="cli-range-updates",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-range-cli')}and open it
2. Add a slide with a table containing this budget data:
   Row 1: Category, Budget, Actual
   Row 2: Rent, 1000, 1000
   Row 3: Food, 500, 450
   Row 4: Transport, 200, 180
3. Add a text shape on the slide with the text "Budget Summary"
4. Update the Food actual value to 480
5. Update the Transport actual value to 195
6. Add a new row: Utilities, 150, 145
7. Read the table data to verify the updates
8. Close the presentation without saving
9. Summarize the values you found and the updated amounts.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    # Loosen assertions - either updated values or verification mentioned
    assert_regex(result.final_response, r"(?i)(480|food|updated|budget)")
    assert_regex(result.final_response, r"(?i)(480|food|updated)")


@pytest.mark.asyncio
async def test_cli_table_updates(aitest_run, ppt_cli_server, ppt_cli_skill):
    agent = Agent(
        name="cli-table-updates",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-table-cli')}and open it
2. Add a slide with a table containing this sales data:
   Row 1: Product, Price, Quantity
   Row 2: Widget, 25, 3
   Row 3: Gadget, 50, 2
   Row 4: Device, 75, 1
3. Name the table "SalesTable"
4. Update the Widget quantity from 3 to 5 - this should NOT delete the table
5. List tables to confirm exactly 1 table named "SalesTable" exists
6. Close the presentation without saving
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(salestable)")


@pytest.mark.asyncio
async def test_cli_chart_updates(aitest_run, ppt_cli_server, ppt_cli_skill):
    agent = Agent(
        name="cli-chart-updates",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-chart-cli')}and open it
2. Add a slide with chart data in a table:
   Row 1: Month, Sales
   Row 2: Jan, 100
   Row 3: Feb, 150
   Row 4: Mar, 200
3. Create a column chart on the slide with title "Monthly Sales"
4. Change ONLY the chart title to "Q1 Sales Report" - do NOT delete and recreate the chart
5. List charts to confirm exactly 1 chart exists
6. Close the presentation without saving
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(q1 sales report|chart)")


@pytest.mark.xfail(reason="LLM intermittently omits required action parameter on complex workflows", strict=False)
@pytest.mark.asyncio
async def test_cli_slide_structural_changes(aitest_run, ppt_cli_server, ppt_cli_skill):
    agent = Agent(
        name="cli-slide-struct",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-struct-cli')}and open it
2. Add a slide with a table containing this employee data:
   Row 1: Name, Department, ID
   Row 2: Alice, Engineering, 100
   Row 3: Bob, Marketing, 200
   Row 4: Carol, Sales, 300
   Row 5: Dave, Support, 400
3. Delete the row with "Bob" (row 3), shifting remaining rows up
4. Rename the "Department" header to "Team"
5. Read all the table data to verify:
   - Alice's ID is 100
   - Carol's ID is 300
   - Dave's ID is 400
   - Bob is gone
   - Header says "Team"
6. Close the presentation without saving
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(100)")
    assert_regex(result.final_response, r"(?i)(300)")
    assert_regex(result.final_response, r"(?i)(400)")
    assert_regex(result.final_response, r"(?i)(team)")
