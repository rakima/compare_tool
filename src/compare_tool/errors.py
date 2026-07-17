class CompareToolError(Exception):
    """A user-facing application error."""


class InvalidInputError(CompareToolError):
    pass


class WorkbookReadError(CompareToolError):
    pass


class PasswordProtectedWorkbookError(WorkbookReadError):
    pass


class OutputWriteError(CompareToolError):
    pass
