"""Submission-wide constants that must remain identical in code and traces."""

MODEL_ID = "Qwen/Qwen3.8-27B"
KEY_ENV_VAR = "TINKER_API_KEY"
TEMPERATURE = 0.0
# The pinned Qwen chat template inserts an empty thinking block and starts the
# answer immediately after it when this flag is false.
THINKING_MODE = "disabled"
REASONING_EFFORT = None
TOKENIZER_REVISION = "1d4bf0f2ff6012fd82039f2fa52739d0dd7c60c0"
TOKENIZER_CHAT_TEMPLATE_SHA256 = "c3cf9e34abf4f9e36c2d72165aa9c132d3e2a725b6c2586aaa3a8af9d7a81041"
STOP_TOKEN = "<|im_end|>"
STOP_TOKEN_ID = 248_046
MODEL_PROVENANCE = "configured_base_model"
RENDERER_ID = "transformers-chat-template/qwen3.8-disable-thinking-pinned"
TRANSPORT_ID = "tinker-native-sampling/0.27.1"
PROVIDER_MAX_RETRIES = 0
NUM_SAMPLES = 1
TINKER_TELEMETRY_ENV_VAR = "TINKER_TELEMETRY"
TINKER_TELEMETRY_VALUE = "0"
TINKER_TELEMETRY_MODE = "disabled"
MAX_AUTHORED_CHARS = 20_000
DEFAULT_MAX_OUTPUT_TOKENS = 8_192
