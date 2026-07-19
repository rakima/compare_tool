"""Excel comparison application."""

from ._version import __version__
from .models import CompareAlgorithm, CompareOptions, CompareResult, Difference, DifferenceType
from .usecase import CompareUseCase

__all__ = [
    "CompareOptions",
    "CompareAlgorithm",
    "CompareResult",
    "CompareUseCase",
    "Difference",
    "DifferenceType",
    "__version__",
]
