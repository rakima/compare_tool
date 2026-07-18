from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Callable
from typing import Generic, TypeVar

from .models import CompareOptions, CompareResult

T = TypeVar("T")
CancelCheck = Callable[[], bool]


class Comparer(ABC, Generic[T]):
    """Format-independent comparison strategy."""

    @abstractmethod
    def compare(
        self,
        old: T,
        new: T,
        options: CompareOptions,
        cancel_requested: CancelCheck | None = None,
    ) -> CompareResult:
        raise NotImplementedError
