"""Same baseline through Tinker: a base model, or your fine-tuned sampler checkpoint.

    uv sync --locked --extra native-tinker
    uv run baseline/tinker_predict.py --out-dir submissions/qwen3-8b \
        --base-model Qwen/Qwen3.8-27B --ids 13-1,51-12
    uv run baseline/tinker_predict.py --out-dir submissions/mine \
        --base-model Qwen/Qwen3.8-27B \
        --model-path tinker://<run-id>/sampler_weights/final

Needs TINKER_API_KEY in .env. The base model picks the tokenizer and chat template. Writes the same
files as llm_predict.py.
"""

import argparse
import asyncio
import os
from pathlib import Path

# This comparison script owns its process environment. Disable SDK telemetry
# before importing Tinker as well as before constructing the client below.
os.environ["TINKER_TELEMETRY"] = "0"

import tinker
from common import FORMAT_HINT, SYSTEM_PROMPT, load_env, parse_ids, run, selected_tasks
from tinker import types
from tinker_cookbook import renderers
from tinker_cookbook.model_info import get_recommended_renderer_name
from tinker_cookbook.tokenizer_utils import get_tokenizer

from sb import DEFAULT_DATASET


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("--out-dir", required=True)
    p.add_argument("--dataset-dir", default=str(DEFAULT_DATASET))
    p.add_argument("--ids", help="comma-separated task ids (default: all)")
    p.add_argument("--base-model", required=True, help="e.g. Qwen/Qwen3-8B")
    p.add_argument(
        "--model-path", help="tinker://... sampler checkpoint. Omit to sample the base model."
    )
    p.add_argument(
        "--renderer",
        help=(
            "Tinker Cookbook renderer name. Omit to use the renderer recommended for "
            "--base-model; Qwen3.8 options include qwen3_8_xhigh_reasoning, "
            "qwen3_8_medium_reasoning, qwen3_8_low_reasoning and "
            "qwen3_8_disable_thinking."
        ),
    )
    p.add_argument("--concurrency", type=int, default=4)
    p.add_argument(
        "--max-tokens", type=int, default=8192, help="sheet-level tasks need long replies"
    )
    return p.parse_args()


async def main():
    load_env()
    # Tinker 0.27.1 can attach complete request headers to telemetry for some
    # API errors. Never allow the comparison runner to transmit credential
    # headers as event data, even if a parent shell or .env enables telemetry.
    os.environ["TINKER_TELEMETRY"] = "0"
    args = parse_args()
    sampler = tinker.ServiceClient().create_sampling_client(
        base_model=args.base_model, model_path=args.model_path
    )
    renderer_name = args.renderer or get_recommended_renderer_name(args.base_model)
    renderer = renderers.get_renderer(renderer_name, get_tokenizer(args.base_model))
    params = types.SamplingParams(
        max_tokens=args.max_tokens, temperature=0, stop=renderer.get_stop_sequences()
    )

    async def complete(prompt: str):
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt + FORMAT_HINT},
        ]
        model_input = renderer.build_generation_prompt(messages)
        response = await sampler.sample_async(
            prompt=model_input, num_samples=1, sampling_params=params
        )
        tokens = response.sequences[0].tokens
        content = renderer.parse_response(tokens)[0]["content"]
        if not isinstance(
            content, str
        ):  # thinking renderers return parts; keep the text, drop the thinking
            content = "".join(
                part.get("text", "") for part in content if part.get("type") == "text"
            )
        return content, model_input.length, len(tokens)

    tasks = selected_tasks(Path(args.dataset_dir), parse_ids(args.ids))
    await run(
        complete, args.model_path or args.base_model, tasks, Path(args.out_dir), args.concurrency
    )


if __name__ == "__main__":
    asyncio.run(main())
