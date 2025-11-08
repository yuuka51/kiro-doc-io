"""Configuration management for document format MCP server."""

import json
import os
from pathlib import Path
from typing import Any

from .errors import ConfigurationError


class Config:
    """Configuration manager for MCP server."""
    
    DEFAULT_CONFIG = {
        "google_credentials_path": "~/.config/kiro-mcp/google-credentials.json",
        "output_directory": "~/Documents/kiro-output",
        "max_file_size_mb": 100,
        "max_sheets": 100,
        "max_slides": 500,
        "api_timeout_seconds": 60,
        "enable_google_workspace": True,
    }
    
    def __init__(self, config_path: str | None = None):
        """
        Initialize configuration.
        
        Args:
            config_path: Optional path to configuration file
        """
        self._config = self.DEFAULT_CONFIG.copy()
        
        # Load from config file if provided
        if config_path:
            self._load_from_file(config_path)
        
        # Override with environment variables
        self._load_from_env()
        
        # Expand paths
        self._expand_paths()
        
        # Validate configuration
        self._validate()
    
    def _load_from_file(self, config_path: str) -> None:
        """Load configuration from JSON file."""
        path = Path(config_path).expanduser()
        
        if not path.exists():
            raise ConfigurationError(
                f"Configuration file not found: {config_path}",
                {"path": str(path)}
            )
        
        try:
            with open(path, "r", encoding="utf-8") as f:
                file_config = json.load(f)
                self._config.update(file_config)
        except json.JSONDecodeError as e:
            raise ConfigurationError(
                f"Invalid JSON in configuration file: {config_path}",
                {"error": str(e)}
            )
        except Exception as e:
            raise ConfigurationError(
                f"Failed to load configuration file: {config_path}",
                {"error": str(e)}
            )
    
    def _load_from_env(self) -> None:
        """Load configuration from environment variables."""
        env_mappings = {
            "GOOGLE_APPLICATION_CREDENTIALS": "google_credentials_path",
            "MCP_OUTPUT_DIR": "output_directory",
            "MCP_LOG_LEVEL": "log_level",
            "MCP_MAX_FILE_SIZE_MB": ("max_file_size_mb", int),
            "MCP_API_TIMEOUT": ("api_timeout_seconds", int),
        }
        
        for env_var, config_key in env_mappings.items():
            value = os.environ.get(env_var)
            if value:
                if isinstance(config_key, tuple):
                    key, converter = config_key
                    try:
                        self._config[key] = converter(value)
                    except ValueError:
                        raise ConfigurationError(
                            f"Invalid value for {env_var}: {value}",
                            {"env_var": env_var, "value": value}
                        )
                else:
                    self._config[config_key] = value
    
    def _expand_paths(self) -> None:
        """Expand user paths in configuration."""
        path_keys = ["google_credentials_path", "output_directory"]
        
        for key in path_keys:
            if key in self._config and isinstance(self._config[key], str):
                self._config[key] = str(Path(self._config[key]).expanduser())
    
    def _validate(self) -> None:
        """Validate configuration values."""
        # Validate numeric values
        if self._config["max_file_size_mb"] <= 0:
            raise ConfigurationError("max_file_size_mb must be positive")
        
        if self._config["max_sheets"] <= 0:
            raise ConfigurationError("max_sheets must be positive")
        
        if self._config["max_slides"] <= 0:
            raise ConfigurationError("max_slides must be positive")
        
        if self._config["api_timeout_seconds"] <= 0:
            raise ConfigurationError("api_timeout_seconds must be positive")
        
        # Create output directory if it doesn't exist
        output_dir = Path(self._config["output_directory"])
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            raise ConfigurationError(
                f"Failed to create output directory: {output_dir}",
                {"error": str(e)}
            )
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value."""
        return self._config.get(key, default)
    
    def __getitem__(self, key: str) -> Any:
        """Get configuration value using dictionary syntax."""
        return self._config[key]
    
    def __contains__(self, key: str) -> bool:
        """Check if configuration key exists."""
        return key in self._config
    
    @property
    def google_credentials_path(self) -> str:
        """Get Google credentials path."""
        return self._config["google_credentials_path"]
    
    @property
    def output_directory(self) -> str:
        """Get output directory path."""
        return self._config["output_directory"]
    
    @property
    def max_file_size_mb(self) -> int:
        """Get maximum file size in MB."""
        return self._config["max_file_size_mb"]
    
    @property
    def max_sheets(self) -> int:
        """Get maximum number of sheets."""
        return self._config["max_sheets"]
    
    @property
    def max_slides(self) -> int:
        """Get maximum number of slides."""
        return self._config["max_slides"]
    
    @property
    def api_timeout_seconds(self) -> int:
        """Get API timeout in seconds."""
        return self._config["api_timeout_seconds"]
    
    @property
    def enable_google_workspace(self) -> bool:
        """Check if Google Workspace is enabled."""
        return self._config["enable_google_workspace"]
