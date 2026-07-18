"""Excel comparison application."""

from importlib.metadata import PackageNotFoundError, version

from .models import CompareAlgorithm, CompareOptions, CompareResult, Difference, DifferenceType
from .usecase import CompareUseCase

try:
    __version__ = version("compare-tool")
except PackageNotFoundError:
    __version__ = "0.0.0"

__all__ = [
    "CompareOptions",
    "CompareAlgorithm",
    "CompareResult",
    "CompareUseCase",
    "Difference",
    "DifferenceType",
    "__version__",
]
