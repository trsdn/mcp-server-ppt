"""CLI range workflows."""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, Provider

from conftest import (
    assert_cli_args_contain,
    assert_cli_exit_codes,
    assert_regex,
    unique_path,
    DEFAULT_RETRIES,
    DEFAULT_TIMEOUT_MS,
)

pytestmark = [pytest.mark.aitest, pytest.mark.cli]


@pytest.mark.asyncio
async def test_cli_range_set_get(aitest_run, ppt_cli_server, ppt_cli_skill, fixtures_dir):
    agent = Agent(
        name="cli-range",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )
    values_file= (fixtures_dir / "range-test-data.json").as_posix()

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-range-cli')}
2. Add a slide and place a text shape with this data using the values from this JSON file:
   {values_file}

   IMPORTANT: JSON arrays with commas break CLI argument parsing.
   You MUST use --values-file with the path above instead of --values with inline JSON.
3. Read back the shape content to verify it was written correctly
4. Close the presentation without saving
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
    assert_cli_args_contain(result, "--values-file")
    assert_regex(result.final_response, r"(?i)(Product)")


@pytest.mark.asyncio
async def test_cli_range_error_handling(aitest_run, ppt_cli_server, ppt_cli_skill):
    agent = Agent(
        name="cli-range-error",
        provider=Provider(model="azure/gpt-4.1", rpm=10, tpm=10000),
        cli_servers=[ppt_cli_server],
        skill=ppt_cli_skill,
        max_turns=20,
        retries=DEFAULT_RETRIES,
    )

    prompt = f"""
1. Create a new empty PowerPoint presentation at {unique_path('llm-test-range-error-cli')}
2. Try to read content from a shape that may not exist to see what happens
3. Then close the presentation without saving
"""
    result = await aitest_run(agent, prompt, timeout_ms=DEFAULT_TIMEOUT_MS)
    assert result.success
    assert_cli_exit_codes(result)
