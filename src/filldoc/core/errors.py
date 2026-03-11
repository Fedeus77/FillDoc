class FillDocError(Exception):
    """Base error for user-facing failures."""


class SettingsError(FillDocError):
    pass


class ExcelError(FillDocError):
    pass


class TemplateError(FillDocError):
    pass


class GenerationError(FillDocError):
    pass

