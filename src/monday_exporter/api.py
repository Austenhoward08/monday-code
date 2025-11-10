"""HTTP client for interacting with the Monday.com GraphQL API."""

from __future__ import annotations

import logging
from contextlib import AbstractContextManager
from typing import Any, Dict, List, Optional

import requests
from requests import Response, Session

from .config import Settings
from .models import Board, Item
from .version import __version__

logger = logging.getLogger(__name__)


BOARD_METADATA_QUERY = """
query ($board_id: [ID!]) {
  boards(ids: $board_id) {
    id
    name
    description
    state
    columns {
      id
      title
      type
    }
    groups {
      id
      title
      color
      position
    }
  }
}
"""


ITEMS_PAGE_QUERY = """
query (
  $board_id: ID!,
  $limit: Int!,
  $cursor: String,
  $include_subitems: Boolean!
) {
  boards(ids: [$board_id]) {
    items_page(limit: $limit, cursor: $cursor) {
      cursor
      items {
        id
        name
        created_at
        updated_at
        group {
          id
          title
          color
          position
        }
        creator {
          id
          name
        }
        column_values {
          id
          text
          title
          type
          value
        }
        subitems @include(if: $include_subitems) {
          id
          name
          created_at
          updated_at
          group {
            id
            title
            color
            position
          }
          creator {
            id
            name
          }
          column_values {
            id
            text
            title
            type
            value
          }
        }
      }
    }
  }
}
"""


class MondayClient(AbstractContextManager["MondayClient"]):
    """Wraps HTTP operations required to communicate with Monday.com."""

    def __init__(self, settings: Settings):
        self._settings = settings
        self._session: Session = requests.Session()
        token = settings.api_token.get_secret_value()
        self._session.headers.update(
            {
                "Authorization": token,
                "Content-Type": "application/json",
                "User-Agent": f"monday-exporter/{__version__}",
            }
        )

    def __enter__(self) -> "MondayClient":
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> Optional[bool]:
        self.close()
        return None

    def close(self) -> None:
        """Close the underlying HTTP session."""
        self._session.close()

    def execute(self, query: str, variables: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Execute a GraphQL query against the Monday.com API."""
        payload = {"query": query, "variables": variables or {}}
        response = self._session.post(
            str(self._settings.api_url),
            json=payload,
            timeout=self._settings.timeout_seconds,
        )
        _raise_for_status(response)

        data = response.json()
        if "errors" in data:
            raise RuntimeError(f"Monday.com API returned errors: {data['errors']!r}")
        return data.get("data", {})

    def fetch_board(
        self,
        board_id: int,
        include_subitems: bool = False,
    ) -> Board:
        """
        Fetch complete board information, including all items.

        Args:
            board_id: Identifier of the Monday.com board to export.
            include_subitems: Whether to fetch subitems for each item.
        """
        metadata = self.execute(BOARD_METADATA_QUERY, {"board_id": [board_id]})
        boards = metadata.get("boards", [])
        if not boards:
            raise RuntimeError(f"No board found for id {board_id}")

        board_model = Board.model_validate({**boards[0], "items": []})

        items: List[Item] = []
        cursor: Optional[str] = None

        while True:
            page = self.execute(
                ITEMS_PAGE_QUERY,
                {
                    "board_id": board_id,
                    "limit": self._settings.page_size,
                    "cursor": cursor,
                    "include_subitems": include_subitems,
                },
            )

            boards_payload = page.get("boards", [])
            if not boards_payload:
                break

            page_info = boards_payload[0].get("items_page", {})
            page_items = page_info.get("items", []) or []

            for item_payload in page_items:
                items.append(Item.model_validate(item_payload))

            cursor = page_info.get("cursor")
            if not cursor:
                break

        if not items:
            logger.warning("No items retrieved for board %s", board_id)

        return board_model.model_copy(update={"items": items})


def _raise_for_status(response: Response) -> None:
    """Raise a detailed HTTP error if the response indicates failure."""
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        details: Dict[str, Any] = {}
        try:
            details = response.json()
        except ValueError:
            details = {"text": response.text}
        raise RuntimeError(
            f"Monday.com API request failed with status {response.status_code}: {details}"
        ) from exc


__all__ = ["MondayClient"]
