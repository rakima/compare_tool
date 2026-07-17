from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Generic, TypeVar

from .models import CompareOptions, CompareResult

T = TypeVar("T")


class Comparer(ABC, Generic[T]):
    """Format-independent comparison strategy."""

    @abstractmethod
    def compare(self, old: T, new: T, options: CompareOptions) -> CompareResult:
        raise NotImplementedError
