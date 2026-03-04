"""MCP file and slide workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_file_and_slide_workflow(aitest_run, ppt_mcp_server, ppt_mcp_skill):
    agent = Agent(
        name="mcp-file-slide",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[ppt_mcp_server],
        skill=ppt_mcp_skill,
        max_turns=25,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
Create a new PowerPoint presentation at {unique_path('budget')}

Set it up with two slides: one titled "Income" and one titled "Expenses".

On the Income slide, add a table with this data:
- Headers: Source, Amount
- Salary: 5000
- Freelance: 1200

On the Expenses slide, add a table with:
- Headers: Category, Amount
- Rent: 1500
- Utilities: 200
- Food: 600

Save the presentation when done.
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
