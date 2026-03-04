"""CLI chart positioning workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.xfail(reason="LLM intermittently omits required action parameter on complex workflows", strict=False)
@pytest.mark.asyncio
async def test_cli_chart_position_below_data(aitest_run, ppt_cli_server, ppt_cli_skill):
    agent = Agent(
        name="cli-chart-below",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-chart-pos-cli')}and open it
2. Add a slide with this sales data as a table:
   Row 1: Month, Revenue, Expenses
   Row 2: January, 50000, 35000
   Row 3: February, 55000, 38000
   Row 4: March, 48000, 32000
   Row 5: April, 62000, 41000
   Row 6: May, 58000, 39000
3. Create a column chart from the Revenue and Expenses data with Month labels
4. Position the chart so it does NOT overlap with the table on the slide
5. List the charts and report the exact chart position
6. Save and close the presentation
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(chart|created)")


@pytest.mark.asyncio
async def test_cli_chart_position_right_of_table(aitest_run, ppt_cli_server, ppt_cli_skill):
    agent = Agent(
        name="cli-chart-right",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-chart-table-cli')}and open it
2. Add a slide with this product data as a table:
   Row 1: Product, Q1, Q2, Q3
   Row 2: Widget, 100, 150, 120
   Row 3: Gadget, 80, 90, 110
   Row 4: Device, 200, 180, 220
   Row 5: Tool, 50, 60, 75
3. Create a line chart from the table's numeric data (Q1, Q2, Q3 columns)
4. Position the chart so it doesn't overlap the table on the slide
5. Save and close the presentation
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(chart|created|productsales)")
