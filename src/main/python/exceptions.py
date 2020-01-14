class Error(Exception):
    """Base class for exceptions in this module"""
    pass


class MultipleFilesFoundError(Error):
    """Exception raised when more than one spreadsheet is found.

    Attributes:
        expression -- input expression in which the error occurred
        message -- explanation of the error
    """
    pass
