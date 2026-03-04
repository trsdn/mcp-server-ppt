"""CLI styling workflows — validates correct style system usage per object type."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_cli_exit_codes, assert_regex, unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_styling_table_style(aitest_run, ppt_cli_server, ppt_cli_skill):
    """LLM should use table(set-style) for table visual styling, not range_format on header."""
    agent = Agent(
        name="cli-styling-table",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
Create a new PowerPoint presentation at {unique_path('llm-test-styling-table')}

Add a slide with this quarterly sales data as a table:
Region, Q1, Q2, Q3, Q4
North, 120000, 135000, 118000, 142000
South, 98000, 102000, 115000, 128000
West, 85000, 91000, 99000, 108000

Name the table "QuarterlySales" and apply a visually appealing style.

Close the presentation without saving.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(QuarterlySales|table|style)")


@pytest.mark.asyncio
async def test_cli_styling_semantic_status(aitest_run, ppt_cli_server, ppt_cli_skill):
    """LLM should use range_format(set-style) with Good/Bad/Neutral for status cells."""
    agent = Agent(
        name="cli-styling-status",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
Create a new PowerPoint presentation at {unique_path('llm-test-styling-status')}

Add a slide with this project status data as a table:
Task, Owner, Status
Design, Alice, Complete
Development, Bob, In Progress
Testing, Carol, Overdue
Deployment, Dave, Complete

Format the Status column shapes with distinct colours to make the status
visually clear at a glance — green for Complete, red for Overdue,
yellow or neutral for In Progress.

Close the presentation without saving.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(format|style|colour|color|green|red)")


@pytest.mark.asyncio
async def test_cli_styling_header_fill(aitest_run, ppt_cli_server, ppt_cli_skill):
    """LLM should use format-range (not set-style) for a header row with a fill colour."""
    agent = Agent(
        name="cli-styling-header",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
Create a new PowerPoint presentation at {unique_path('llm-test-styling-header')}

Add a slide with this data as a table:
Product, Units, Revenue
Widget, 450, 13500
Gadget, 280, 19600
Doohickey, 175, 8750

Give the header row a dark blue background with white bold text,
centred horizontally.

Close the presentation without saving.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_regex(result.final_response, r"(?i)(header|format|blue|white|bold)")
