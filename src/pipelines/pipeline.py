import time

from src.agents.agents import build_search_agent, build_reader_agent, writer_chain, critic_chain


class PipelineStepError(Exception):
    """Raised when a pipeline step fails after all retries.
    Carries the step key so callers (e.g. the Streamlit UI) know exactly
    which stage broke, instead of getting an opaque crash."""
    def __init__(self, step: str, original: Exception):
        self.step = step
        self.original = original
        super().__init__(f"Step '{step}' failed: {original}")


def _with_retries(fn, retries: int = 2, delay: float = 1.5):
    """Groq's openai/gpt-oss-120b model occasionally hallucinates malformed
    tool-call arguments (e.g. inventing {"cursor": 1, "id": 0} instead of
    {"query": ...}), which the API rejects with a 400 BadRequestError. This
    is a stochastic model glitch, not a logic bug, so a fresh retry usually
    succeeds. We rebuild the agent each attempt so no bad conversation state
    carries over."""
    last_err = None
    for attempt in range(1, retries + 2):  # e.g. retries=2 -> 3 total attempts
        try:
            return fn()
        except Exception as e:
            last_err = e
            print(f"  [retry] attempt {attempt} failed: {e}")
            if attempt < retries + 1:
                time.sleep(delay)
    raise last_err


def run_research_pipeline(topic: str, on_step_start=None, on_step_done=None) -> dict:
    """
    Runs the full search -> reader -> writer -> critic pipeline.

    on_step_start(step_key) and on_step_done(step_key, result) are optional
    callbacks so a UI (e.g. app.py) can render live progress without
    reimplementing this orchestration itself. Both are no-ops if omitted,
    so main.py's plain `run_research_pipeline(topic)` call still works
    unchanged.

    step_key is one of: "search", "reader", "writer", "critic".
    Raises PipelineStepError(step, original_exception) if a step fails
    after retries, so the caller knows exactly where it broke.
    """
    state = {}

    def run_step(step_key: str, fn):
        if on_step_start:
            on_step_start(step_key)
        try:
            result = _with_retries(fn)
        except Exception as e:
            raise PipelineStepError(step_key, e) from e
        state[step_key] = result
        if on_step_done:
            on_step_done(step_key, result)
        return result

    # ── Step 1: Search ──
    print("step 1 - search agent is working ...")

    def _search():
        search_agent = build_search_agent()
        sr = search_agent.invoke({
            "messages": [("user", f"Find recent, reliable and detailed information about: {topic}")]
        })
        return sr["messages"][-1].content

    run_step("search", _search)
    print("\n search result ", state["search"])

    # ── Step 2: Reader ──
    print("step 2 - Reader agent is scraping top resources ...")

    def _reader():
        reader_agent = build_reader_agent()
        rr = reader_agent.invoke({
            "messages": [(
                "user",
                f"Based on the following search results about '{topic}', "
                f"pick the most relevant URL and scrape it for deeper content.\n\n"
                f"Search Results:\n{state['search'][:800]}"
            )]
        })
        return rr["messages"][-1].content

    run_step("reader", _reader)
    print("\nscraped content: \n", state["reader"])

    # ── Step 3: Writer ──
    print("step 3 - Writer is drafting the report ...")

    def _writer():
        research_combined = (
            f"SEARCH RESULTS : \n {state['search']} \n\n"
            f"DETAILED SCRAPED CONTENT : \n {state['reader']}"
        )
        return writer_chain.invoke({"topic": topic, "research": research_combined})

    run_step("writer", _writer)
    print("\n Final Report\n", state["writer"])

    # ── Step 4: Critic ──
    print("step 4 - critic is reviewing the report ")

    def _critic():
        return critic_chain.invoke({"report": state["writer"]})

    run_step("critic", _critic)
    print("\n critic report \n", state["critic"])

    return state