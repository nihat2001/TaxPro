from operator import add
from typing import Annotated, Any, NotRequired, TypedDict

from langchain_core.messages import AnyMessage
from langgraph.graph.message import add_messages


class BaseAgentState(TypedDict):
    messages: Annotated[list[AnyMessage], add_messages]
    remaining_steps: NotRequired[int]


class ParserState(TypedDict, total=False):
    file_bytes: str
    source_name: str
    rows: list[dict]
    inserted: int
    skipped: list[str]


class BankVatState(TypedDict, total=False):
    file_bytes: str
    source: str
    analysis: dict


class RetrieverState(TypedDict, total=False):
    query: str
    profile_id: str | None
    answer: str
    sources: list[str]


def _merge_dict(left: Any, right: Any) -> dict:
    merged = dict(left or {})
    merged.update(right or {})
    return merged


class AgentState(BaseAgentState):
    parser: Annotated[ParserState, _merge_dict]
    bank: Annotated[BankVatState, _merge_dict]
    retriever: Annotated[RetrieverState, _merge_dict]
    try_count: Annotated[int, add]
    max_try: int
