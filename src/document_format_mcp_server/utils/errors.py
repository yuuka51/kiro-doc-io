"""Custom exception classes for document format MCP server."""


class DocumentMCPError(Exception):
    """Base exception for all document MCP server errors."""
    
    def __init__(self, message: str, details: dict | None = None):
        self.message = message
        self.details = details or {}
        super().__init__(self.message)


class FileNotFoundError(DocumentMCPError):
    """Raised when a file cannot be found."""
    pass


class CorruptedFileError(DocumentMCPError):
    """Raised when a file is corrupted or cannot be read."""
    pass


class AuthenticationError(DocumentMCPError):
    """Raised when Google API authentication fails."""
    pass


class PermissionError(DocumentMCPError):
    """Raised when file access permission is denied."""
    pass


class APIError(DocumentMCPError):
    """Raised when a Google API call fails."""
    pass


class ValidationError(DocumentMCPError):
    """Raised when input data validation fails."""
    pass


class ConfigurationError(DocumentMCPError):
    """Raised when configuration is invalid or missing."""
    pass
