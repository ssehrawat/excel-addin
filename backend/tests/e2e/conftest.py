"""Layer 4 E2E test configuration.

Registers --live and --provider CLI options for pytest.
"""


def pytest_addoption(parser):
    parser.addoption("--live", action="store_true", default=False,
                     help="Run live LLM calls instead of replaying cassettes")
    parser.addoption("--provider", default="openai",
                     help="LLM provider for live mode (openai or anthropic)")
