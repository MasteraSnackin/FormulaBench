"""One-shot LLM baseline via Pydantic AI + OpenRouter.

    uv sync --locked --extra baseline
    uv run baseline/llm_predict.py --out-dir submissions/my-llm --model deepseek/deepseek-v3.2 --ids 13-1,51-12

Needs OPENROUTER_API_KEY in .env. Writes predictions.jsonl, outputs/, traces/ and run.log into --out-dir.
"""

import argparse
import asyncio
import os
from pathlib import Path

from common import SYSTEM_PROMPT, SpreadsheetAnswer, load_env, parse_ids, run, selected_tasks
from pydantic_ai import Agent
from pydantic_ai.models.openrouter import OpenRouterModelSettings

from sb import DEFAULT_DATASET


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("--out-dir", required=True)
    p.add_argument("--dataset-dir", default=str(DEFAULT_DATASET))
    p.add_argument("--ids", help="comma-separated task ids (default: all)")
    p.add_argument("--model", default=os.environ.get("MODEL", "deepseek/deepseek-v3.2"))
    p.add_argument("--concurrency", type=int, default=4, help="parallel requests")
    return p.parse_args()


def build_agent(model: str) -> Agent[None, SpreadsheetAnswer]:
    return Agent(
        model=f"openrouter:{model}",
        output_type=SpreadsheetAnswer,
        system_prompt=SYSTEM_PROMPT,
        model_settings=OpenRouterModelSettings(temperature=0),
    )


async def main():
    load_env()
    args = parse_args()
    agent = build_agent(args.model)

    async def complete(prompt: str):
        result = await agent.run(prompt)
        usage = result.usage() if callable(result.usage) else result.usage
        return result.output.model_dump_json(), usage.input_tokens, usage.output_tokens

    tasks = selected_tasks(Path(args.dataset_dir), parse_ids(args.ids))
    await run(complete, f"openrouter:{args.model}", tasks, Path(args.out_dir), args.concurrency)


if __name__ == "__main__":
    asyncio.run(main())
