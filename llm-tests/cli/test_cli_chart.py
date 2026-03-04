"""CLI chart workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_chart_workflows(aitest_run, ppt_cli_server, ppt_cli_skill):
    agent = Agent(
        name="cli-chart-workflows",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=25,
        retries=DEFAULT_RETRIES,
    )

    messages = None

    prompt = f"""
Create a sales analysis chart:

1. Create a new PowerPoint presentation at {unique_path('chart-from-table-cli')}
2. Add a slide with this sales data as a table:
   Product, Q1 Sales, Q2 Sales
   Laptop, 45000, 52000
   Phone, 38000, 41000
   Tablet, 22000, 28000
   Monitor, 15000, 18000
3. Create a column chart from the table data
4. Position it on the slide
5. Save and close
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(salesdata|chart|created)")
    messages = result.messages

    prompt = f"""
I need a chart on a slide:

1. Create a new PowerPoint presentation at {unique_path('chart-position-cli')}
2. Add a slide with this budget data as a table:
   Month, Revenue, Expenses
   January, 50000, 35000
   February, 55000, 38000
   March, 48000, 32000
   April, 62000, 41000
   May, 58000, 39000
3. Create a line chart from the data on the same slide
4. Make sure the chart does not overlap the table
5. Read chart info and confirm the position
6. Save and close
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(chart|row [7-9]|position)")
    messages = result.messages

    prompt = f"""
Create a presentation with multiple chart types:

1. Create a new PowerPoint presentation at {unique_path('multi-chart-cli')}
2. Add a slide with this market data as a table:
   Company, Revenue, Market Share
   Alpha, 500000, 35
   Beta, 400000, 28
   Gamma, 300000, 22
   Delta, 200000, 15
3. Create a PIE chart showing Market Share
4. Create a BAR chart showing Revenue
5. List charts and confirm both exist without overlapping
6. Save and close
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(pie|bar|2 chart|two chart)")
    messages = result.messages

    prompt = f"""
Create a chart with positioning on a slide:

1. Create a new PowerPoint presentation at {unique_path('chart-targetrange-cli')}
2. Add a slide with this quarterly data as a table:
   Region, Q1, Q2, Q3, Q4
   North, 1000, 1200, 1100, 1400
   South, 800, 900, 950, 1000
   East, 1500, 1600, 1450, 1700
   West, 600, 700, 650, 800
3. Create a bar chart on the slide from the data
4. Verify the chart was created and its position
5. Save and close
"""
    result = await aitest_run(agent, prompt, messages=messages, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(chart|created)")
