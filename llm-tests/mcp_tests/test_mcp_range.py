"""MCP range workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import assert_regex, unique_path, DEFAULT_RETRIES, DEFAULT_TIMEOUT_MS

pytestmark = [pytest.mark.aitest, pytest.mark.mcp]


@pytest.mark.asyncio
async def test_mcp_range_set_get(aitest_run, ppt_mcp_server, ppt_mcp_skill, fixtures_dir):
    agent = Agent(
        name="mcp-range",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[ppt_mcp_server],
        skill=ppt_mcp_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )
    values_file= (fixtures_dir / "range-test-data.json").as_posix()

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-range')}
2. Add a slide and place a shape with this data:
   Row 1: Product, Quantity, Price
   Row 2: Widget, 10, 5.99
3. Read back the shape content to verify it was written correctly
4. Close the presentation without saving
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert result.tool_was_called("range")
    assert_regex(result.final_response, r"(?i)(Product)")


@pytest.mark.asyncio
async def test_mcp_range_error_handling(aitest_run, ppt_mcp_server, ppt_mcp_skill):
    agent = Agent(
        name="mcp-range-error",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        mcp_servers=[ppt_mcp_server],
        skill=ppt_mcp_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-range-error')}
2. Try to read content from a shape that may not exist to see what happens
3. Then close the presentation without saving
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
