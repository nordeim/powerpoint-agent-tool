# requirements.txt
```txt
# PowerPoint Agent Tool - Dependencies
# Install with: pip install -r requirements.txt

# Core dependencies (required)
python-pptx==0.6.23       # PowerPoint manipulation
Pillow>=12.0.0            # Image processing
pandas>=2.3.2             # Data handling for charts (optional)
jsonschema>=4.25.1         # JSON Schema validation

# Development dependencies (for testing)
# pytest>=8.4.2           # Test runner (optional)
# pytest-cov>=6.3.0       # Coverage reporting (optional)

# Note: Python 3.8+ required
# Note: For PDF export, install LibreOffice separately:
#   - Ubuntu/Debian: sudo apt install libreoffice-impress
#   - macOS: brew install --cask libreoffice
#   - Windows: https://www.libreoffice.org/download

```

# core/__init__.py
```py
"""
PowerPoint Agent Core Package
"""

from .powerpoint_agent_core import (
    # Core class
    PowerPointAgent,
    
    # Exceptions
    PowerPointAgentError,
    SlideNotFoundError,
    LayoutNotFoundError,
    ImageNotFoundError,
    InvalidPositionError,
    TemplateError,
    ThemeError,
    AccessibilityError,
    AssetValidationError,
    FileLockError,
    
    # Helpers
    Position,
    Size,
    ColorHelper,
    TemplateProfile,
    AccessibilityChecker,
    AssetValidator,
    
    # Enums
    ShapeType,
    ChartType,
    TextAlignment,
    VerticalAlignment,
    BulletStyle,
    ImageFormat,
    ExportFormat,
    
    # Constants
    SLIDE_WIDTH_INCHES,
    SLIDE_HEIGHT_INCHES,
    ANCHOR_POINTS,
    CORPORATE_COLORS,
    STANDARD_FONTS,
)

__version__ = "1.0.0"
__all__ = [
    "PowerPointAgent",
    "PowerPointAgentError",
    "SlideNotFoundError",
    "LayoutNotFoundError",
    "ImageNotFoundError",
    "InvalidPositionError",
    "TemplateError",
    "ThemeError",
    "AccessibilityError",
    "AssetValidationError",
    "FileLockError",
    "Position",
    "Size",
    "ColorHelper",
    "TemplateProfile",
    "AccessibilityChecker",
    "AssetValidator",
    "ShapeType",
    "ChartType",
    "TextAlignment",
    "VerticalAlignment",
    "BulletStyle",
    "ImageFormat",
    "ExportFormat",
    "SLIDE_WIDTH_INCHES",
    "SLIDE_HEIGHT_INCHES",
    "ANCHOR_POINTS",
    "CORPORATE_COLORS",
    "STANDARD_FONTS",
]

```

# core/powerpoint_agent_core.py
```py
#!/usr/bin/env python3
"""
PowerPoint Agent Core Library v3.1
Production-grade PowerPoint manipulation with validation, accessibility, and full
alignment with Presentation Architect System Prompt v3.0.

This is the foundational library used by all CLI tools.
Designed for stateless, security-hardened PowerPoint operations.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Changelog v3.1.0 (Security & Governance Release):
- SECURITY: Added approval_token requirement for destructive operations (delete_slide, remove_shape)
- SECURITY: Added Path Traversal protection to PathValidator
- SECURITY: Hardened FileLock with cross-platform atomic operations
- SECURITY: Standardized on SHA-256 for all hashing operations
- OBSERVABILITY: All mutation methods now return presentation_version_before/after
- OBSERVABILITY: Version hashing now includes shape geometry (position/size) to detect layout changes
- SAFETY: Removed silent index clamping (now raises SlideNotFoundError)
- SAFETY: Strict validation for shape types
- FIXED: _log_warning now correctly uses the logger instead of stderr
- FIXED: Redundant imports and duplicate logic consolidated

Changelog v3.0.0 (Major Release):
- NEW: add_notes() - Add/append/prepend/overwrite speaker notes
- NEW: set_z_order() - Control shape layering with 4 actions
- NEW: remove_shape() - Remove shapes from slides
- NEW: set_footer() - Configure footer text, numbers, date
- NEW: set_background() - Set slide/presentation background color or image
- NEW: crop_image() - True image cropping (not just resize)
- NEW: clone_presentation() - Clone presentation to new file
- NEW: get_presentation_version() - Compute deterministic version hash
- NEW: PathValidator class - Security-hardened path validation
- NEW: ShapeNotFoundError, ChartNotFoundError, PathValidationError exceptions
- NEW: ZOrderAction, NotesMode enums
- FIXED: FileLock now uses atomic os.open() with O_CREAT|O_EXCL
- FIXED: Lock released in finally block on open() failure
- FIXED: Slide insertion XML manipulation corrected
- FIXED: Placeholder type handling normalized to integers
- FIXED: Alt text detection checks 'descr' attribute
- FIXED: All bounds checks include negative index validation
- FIXED: Chart update error handling improved
- IMPROVED: All add_* methods return shape index for chaining
- IMPROVED: TemplateProfile uses lazy loading
- IMPROVED: Layout lookup cached for performance
- IMPROVED: Comprehensive docstrings with examples
- IMPROVED: Full alignment with System Prompt v3.0

Changelog v1.1.0:
- Added missing subprocess import for PDF export
- Added missing PP_PLACEHOLDER import and constants
- Replaced all magic numbers with named constants
- Removed text truncation in get_slide_info()
- Added position/size information to shape inspection
- Added placeholder subtype decoding
- Replaced print() with proper logging

Dependencies:
- python-pptx >= 0.6.21 (required)
- Pillow >= 9.0.0 (optional, for image operations)
"""

import os
import re
import sys
import json
import hashlib
import subprocess
import tempfile
import shutil
import time
import logging
import platform
import errno
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, Tuple
from enum import Enum
from datetime import datetime
from io import BytesIO
from lxml import etree
from pptx.oxml.ns import qn

# ============================================================================
# THIRD-PARTY IMPORTS WITH GRACEFUL DEGRADATION
# ============================================================================

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.dml import MSO_THEME_COLOR
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False
    raise ImportError(
        "python-pptx is required. Install with:\n"
        "  pip install python-pptx\n"
        "  or: uv pip install python-pptx"
    )

try:
    from PIL import Image as PILImage
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False
    PILImage = None


# ============================================================================
# LOGGING SETUP
# ============================================================================

logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter('%(levelname)s:%(name)s:%(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.WARNING)


# ============================================================================
# EXCEPTIONS
# ============================================================================

class PowerPointAgentError(Exception):
    """Base exception for all PowerPoint agent errors."""
    
    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert exception to JSON-serializable dict."""
        return {
            "error": self.__class__.__name__,
            "message": self.message,
            "details": self.details
        }
    
    def to_json(self) -> str:
        """Convert exception to JSON string."""
        return json.dumps(self.to_dict())


class SlideNotFoundError(PowerPointAgentError):
    """Raised when slide index is out of range."""
    pass


class ShapeNotFoundError(PowerPointAgentError):
    """Raised when shape index is out of range."""
    pass


class ChartNotFoundError(PowerPointAgentError):
    """Raised when chart is not found at specified index."""
    pass


class LayoutNotFoundError(PowerPointAgentError):
    """Raised when requested layout doesn't exist."""
    pass


class ImageNotFoundError(PowerPointAgentError):
    """Raised when image file is not found."""
    pass


class InvalidPositionError(PowerPointAgentError):
    """Raised when position specification is invalid."""
    pass


class TemplateError(PowerPointAgentError):
    """Raised when template operations fail."""
    pass


class ThemeError(PowerPointAgentError):
    """Raised when theme operations fail."""
    pass


class AccessibilityError(PowerPointAgentError):
    """Raised when accessibility validation fails."""
    pass


class AssetValidationError(PowerPointAgentError):
    """Raised when asset validation fails."""
    pass


class FileLockError(PowerPointAgentError):
    """Raised when file cannot be locked for exclusive access."""
    pass


class PathValidationError(PowerPointAgentError):
    """Raised when path validation fails (security)."""
    pass


class ApprovalTokenError(PowerPointAgentError):
    """Raised when a destructive operation lacks a valid approval token."""
    pass


# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"
__author__ = "PowerPoint Agent Team"
__license__ = "MIT"

# Standard slide dimensions (16:9 widescreen) in inches
SLIDE_WIDTH_INCHES = 13.333
SLIDE_HEIGHT_INCHES = 7.5

# Alternative dimensions (4:3 standard) in inches
SLIDE_WIDTH_4_3_INCHES = 10.0
SLIDE_HEIGHT_4_3_INCHES = 7.5

# EMU conversion constant
EMU_PER_INCH = 914400

# Governance Scopes
APPROVAL_SCOPE_DELETE_SLIDE = "delete:slide"
APPROVAL_SCOPE_REMOVE_SHAPE = "remove:shape"

# Standard anchor points for positioning
ANCHOR_POINTS = {
    "top_left": (0.0, 0.0),
    "top_center": (0.5, 0.0),
    "top_right": (1.0, 0.0),
    "center_left": (0.0, 0.5),
    "center": (0.5, 0.5),
    "center_right": (1.0, 0.5),
    "bottom_left": (0.0, 1.0),
    "bottom_center": (0.5, 1.0),
    "bottom_right": (1.0, 1.0)
}

# Standard corporate colors (RGB tuples)
CORPORATE_COLORS = {
    "primary_blue": RGBColor(0, 112, 192),
    "secondary_gray": RGBColor(89, 89, 89),
    "accent_orange": RGBColor(237, 125, 49),
    "success_green": RGBColor(112, 173, 71),
    "warning_yellow": RGBColor(255, 192, 0),
    "danger_red": RGBColor(192, 0, 0),
    "white": RGBColor(255, 255, 255),
    "black": RGBColor(0, 0, 0)
}

# Standard fonts
STANDARD_FONTS = {
    "title": "Calibri Light",
    "body": "Calibri",
    "code": "Consolas"
}

# WCAG 2.1 color contrast ratios
WCAG_CONTRAST_NORMAL = 4.5
WCAG_CONTRAST_LARGE = 3.0

# Maximum recommended file size (MB)
MAX_RECOMMENDED_FILE_SIZE_MB = 50

# Valid PowerPoint extensions
VALID_PPTX_EXTENSIONS = {'.pptx', '.pptm', '.potx', '.potm'}

# Placeholder type mapping (integer keys for compatibility)
PLACEHOLDER_TYPE_NAMES = {
    0: "OBJECT",
    1: "TITLE",
    2: "BODY",
    3: "CENTER_TITLE",
    4: "SUBTITLE",
    5: "DATE",
    6: "SLIDE_NUMBER",
    7: "FOOTER",
    8: "HEADER",
    9: "OBJECT",
    10: "CHART",
    11: "TABLE",
    12: "CLIP_ART",
    13: "ORG_CHART",
    14: "MEDIA_CLIP",
    15: "BITMAP",
    16: "VERTICAL_TITLE",
    17: "VERTICAL_BODY",
    18: "PICTURE",
}

# Placeholder types that represent titles
TITLE_PLACEHOLDER_TYPES = {1, 3}  # TITLE and CENTER_TITLE

# Placeholder type for subtitle
SUBTITLE_PLACEHOLDER_TYPE = 4


def get_placeholder_type_name(ph_type_value: Any) -> str:
    """
    Safely get human-readable name for placeholder type.
    
    Args:
        ph_type_value: Placeholder type (int or enum)
        
    Returns:
        Human-readable string name
    """
    if ph_type_value is None:
        return "NONE"
    
    # Handle enum types
    if hasattr(ph_type_value, 'value'):
        ph_type_value = ph_type_value.value
    
    try:
        int_value = int(ph_type_value)
        return PLACEHOLDER_TYPE_NAMES.get(int_value, f"UNKNOWN_{int_value}")
    except (TypeError, ValueError):
        return f"UNKNOWN_{ph_type_value}"


def _get_placeholder_type_int_helper(ph_type: Any) -> int:
    """
    Centralized helper to convert placeholder type to integer.
    
    Args:
        ph_type: Placeholder type object or value
        
    Returns:
        Integer representation of type
    """
    if ph_type is None:
        return 0
    if hasattr(ph_type, 'value'):
        return ph_type.value
    try:
        return int(ph_type)
    except (TypeError, ValueError):
        return 0


# ============================================================================
# ENUMS
# ============================================================================

class ShapeType(Enum):
    """Common shape types supported by python-pptx."""
    RECTANGLE = "rectangle"
    ROUNDED_RECTANGLE = "rounded_rectangle"
    ELLIPSE = "ellipse"
    OVAL = "ellipse"
    TRIANGLE = "triangle"
    ARROW_RIGHT = "arrow_right"
    ARROW_LEFT = "arrow_left"
    ARROW_UP = "arrow_up"
    ARROW_DOWN = "arrow_down"
    STAR = "star"
    PENTAGON = "pentagon"
    HEXAGON = "hexagon"


class ChartType(Enum):
    """Supported chart types."""
    COLUMN = "column"
    COLUMN_CLUSTERED = "column"
    COLUMN_STACKED = "column_stacked"
    BAR = "bar"
    BAR_CLUSTERED = "bar"
    BAR_STACKED = "bar_stacked"
    LINE = "line"
    LINE_MARKERS = "line_markers"
    PIE = "pie"
    PIE_EXPLODED = "pie_exploded"
    AREA = "area"
    SCATTER = "scatter"


class TextAlignment(Enum):
    """Text alignment options."""
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "justify"


class VerticalAlignment(Enum):
    """Vertical text alignment."""
    TOP = "top"
    MIDDLE = "middle"
    BOTTOM = "bottom"


class BulletStyle(Enum):
    """Bullet list styles."""
    BULLET = "bullet"
    NUMBERED = "numbered"
    NONE = "none"


class ImageFormat(Enum):
    """Supported image formats."""
    PNG = "png"
    JPG = "jpg"
    JPEG = "jpeg"
    GIF = "gif"
    BMP = "bmp"


class ExportFormat(Enum):
    """Export format options."""
    PDF = "pdf"
    PNG = "png"
    JPG = "jpg"
    PPTX = "pptx"


class ZOrderAction(Enum):
    """Z-order manipulation actions."""
    BRING_TO_FRONT = "bring_to_front"
    SEND_TO_BACK = "send_to_back"
    BRING_FORWARD = "bring_forward"
    SEND_BACKWARD = "send_backward"


class NotesMode(Enum):
    """Speaker notes insertion modes."""
    APPEND = "append"
    PREPEND = "prepend"
    OVERWRITE = "overwrite"


# ============================================================================
# UTILITY CLASSES
# ============================================================================

class FileLock:
    """
    Atomic file locking mechanism for concurrent access prevention.
    
    Uses OS-level atomic file creation to ensure only one process
    can hold the lock at a time.
    """
    
    def __init__(self, filepath: Path, timeout: float = 10.0):
        """
        Initialize file lock.
        
        Args:
            filepath: Path to file to lock
            timeout: Maximum seconds to wait for lock acquisition
        """
        self.filepath = Path(filepath)
        self.lockfile = self.filepath.parent / f".{self.filepath.name}.lock"
        self.timeout = timeout
        self.acquired = False
        self._fd: Optional[int] = None
    
    def acquire(self) -> bool:
        """
        Acquire lock with timeout using atomic file creation.
        
        Returns:
            True if lock acquired, False if timeout
        """
        start_time = time.time()
        
        while time.time() - start_time < self.timeout:
            try:
                # Use O_CREAT | O_EXCL for atomic creation
                # This is atomic on POSIX systems
                self._fd = os.open(
                    str(self.lockfile),
                    os.O_CREAT | os.O_EXCL | os.O_WRONLY,
                    0o644
                )
                self.acquired = True
                return True
            except FileExistsError:
                time.sleep(0.1)
            except OSError as e:
                # EEXIST (cross-platform way via errno)
                if e.errno == errno.EEXIST:
                    time.sleep(0.1)
                else:
                    raise
        
        return False
    
    def release(self) -> None:
        """Release lock and clean up lock file."""
        if self._fd is not None:
            try:
                os.close(self._fd)
            except OSError:
                pass
            self._fd = None
        
        if self.acquired:
            try:
                self.lockfile.unlink(missing_ok=True)
            except OSError:
                pass
            self.acquired = False
    
    def __enter__(self) -> 'FileLock':
        if not self.acquire():
            raise FileLockError(
                f"Could not acquire lock on {self.filepath} within {self.timeout}s",
                details={"filepath": str(self.filepath), "timeout": self.timeout}
            )
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        self.release()
        return False


class PathValidator:
    """
    Security-hardened path validation utility.
    
    Validates file paths to prevent path traversal attacks
    and ensure files are of expected types.
    """
    
    @staticmethod
    def validate_pptx_path(
        filepath: Union[str, Path],
        must_exist: bool = True,
        must_be_writable: bool = False,
        allowed_base_dirs: Optional[List[Path]] = None
    ) -> Path:
        """
        Validate a PowerPoint file path.
        
        Args:
            filepath: Path to validate
            must_exist: If True, file must exist
            must_be_writable: If True, parent directory must be writable
            allowed_base_dirs: Optional list of base directories to restrict access (traversal protection)
            
        Returns:
            Resolved absolute Path
            
        Raises:
            PathValidationError: If validation fails
        """
        try:
            path = Path(filepath).resolve()
        except Exception as e:
            raise PathValidationError(
                f"Invalid path: {filepath}",
                details={"error": str(e)}
            )
        
        # Security: Path Traversal Protection
        if allowed_base_dirs:
            is_allowed = False
            for base in allowed_base_dirs:
                try:
                    # Check if path is relative to base
                    if path.is_relative_to(base.resolve()):
                        is_allowed = True
                        break
                except Exception:
                    continue
            
            if not is_allowed:
                raise PathValidationError(
                    f"Path is not within allowed directories: {path}",
                    details={
                        "path": str(path),
                        "allowed_base_dirs": [str(b) for b in allowed_base_dirs]
                    }
                )

        # Check extension
        if path.suffix.lower() not in VALID_PPTX_EXTENSIONS:
            raise PathValidationError(
                f"Invalid file extension: {path.suffix}",
                details={
                    "path": str(path),
                    "valid_extensions": list(VALID_PPTX_EXTENSIONS)
                }
            )
        
        # Check existence
        if must_exist and not path.exists():
            raise PathValidationError(
                f"File does not exist: {path}",
                details={"path": str(path)}
            )
        
        # Check if it's a file (not directory)
        if must_exist and not path.is_file():
            raise PathValidationError(
                f"Path is not a file: {path}",
                details={"path": str(path)}
            )
        
        # Check writability
        if must_be_writable:
            parent = path.parent
            if not parent.exists():
                raise PathValidationError(
                    f"Parent directory does not exist: {parent}",
                    details={"path": str(path), "parent": str(parent)}
                )
            if not os.access(str(parent), os.W_OK):
                raise PathValidationError(
                    f"Parent directory is not writable: {parent}",
                    details={"path": str(path), "parent": str(parent)}
                )
        
        return path
    
    @staticmethod
    def validate_image_path(filepath: Union[str, Path]) -> Path:
        """
        Validate an image file path.
        
        Args:
            filepath: Path to validate
            
        Returns:
            Resolved absolute Path
            
        Raises:
            ImageNotFoundError: If validation fails
        """
        try:
            path = Path(filepath).resolve()
        except Exception as e:
            raise ImageNotFoundError(
                f"Invalid image path: {filepath}",
                details={"error": str(e)}
            )
        
        if not path.exists():
            raise ImageNotFoundError(
                f"Image file does not exist: {path}",
                details={"path": str(path)}
            )
        
        if not path.is_file():
            raise ImageNotFoundError(
                f"Image path is not a file: {path}",
                details={"path": str(path)}
            )
        
        valid_image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'}
        if path.suffix.lower() not in valid_image_extensions:
            raise ImageNotFoundError(
                f"Invalid image extension: {path.suffix}",
                details={"path": str(path), "valid_extensions": list(valid_image_extensions)}
            )
        
        return path


class Position:
    """Flexible position system supporting multiple input formats."""
    
    @staticmethod
    def from_dict(
        pos_dict: Dict[str, Any],
        slide_width: float = SLIDE_WIDTH_INCHES,
        slide_height: float = SLIDE_HEIGHT_INCHES
    ) -> Tuple[float, float]:
        """
        Convert position dict to (left, top) in inches.
        
        Supports multiple formats:
        1. Absolute inches: {"left": 1.5, "top": 2.0}
        2. Percentage: {"left": "20%", "top": "30%"}
        3. Anchor-based: {"anchor": "center", "offset_x": 0.5, "offset_y": -1.0}
        4. Grid system: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
        
        Args:
            pos_dict: Position specification dictionary
            slide_width: Slide width in inches (for percentage calculations)
            slide_height: Slide height in inches (for percentage calculations)
            
        Returns:
            Tuple of (left, top) in inches
            
        Raises:
            InvalidPositionError: If format is invalid
        """
        if not isinstance(pos_dict, dict):
            raise InvalidPositionError(
                f"Position must be a dictionary, got {type(pos_dict).__name__}",
                details={"value": str(pos_dict)}
            )
        
        # Format 1 & 2: Absolute or percentage with left/top
        if "left" in pos_dict and "top" in pos_dict:
            left = Position._parse_dimension(pos_dict["left"], slide_width)
            top = Position._parse_dimension(pos_dict["top"], slide_height)
            return (left, top)
        
        # Format 3: Anchor-based
        if "anchor" in pos_dict:
            anchor_name = pos_dict["anchor"].lower().replace("-", "_").replace(" ", "_")
            anchor = ANCHOR_POINTS.get(anchor_name)
            
            if anchor is None:
                raise InvalidPositionError(
                    f"Unknown anchor: {pos_dict['anchor']}",
                    details={"available_anchors": list(ANCHOR_POINTS.keys())}
                )
            
            # Anchor is in relative coordinates (0-1), convert to inches
            base_left = anchor[0] * slide_width
            base_top = anchor[1] * slide_height
            
            offset_x = float(pos_dict.get("offset_x", 0))
            offset_y = float(pos_dict.get("offset_y", 0))
            
            return (base_left + offset_x, base_top + offset_y)
        
        # Format 4: Grid system
        if "grid_row" in pos_dict and "grid_col" in pos_dict:
            grid_size = int(pos_dict.get("grid_size", 12))
            cell_width = slide_width / grid_size
            cell_height = slide_height / grid_size
            
            col = int(pos_dict["grid_col"])
            row = int(pos_dict["grid_row"])
            
            left = col * cell_width
            top = row * cell_height
            
            return (left, top)
        
        raise InvalidPositionError(
            "Invalid position format",
            details={
                "provided": pos_dict,
                "expected_formats": [
                    {"left": "value", "top": "value"},
                    {"anchor": "center", "offset_x": 0, "offset_y": 0},
                    {"grid_row": 0, "grid_col": 0, "grid_size": 12}
                ]
            }
        )
    
    @staticmethod
    def _parse_dimension(value: Union[str, float, int], max_dimension: float) -> float:
        """
        Parse dimension value (supports percentages or absolute values).
        
        Args:
            value: Dimension value (e.g., "50%", 2.5, "2.5")
            max_dimension: Maximum dimension for percentage calculation
            
        Returns:
            Dimension in inches
        """
        if isinstance(value, str):
            value = value.strip()
            if value.endswith('%'):
                percent = float(value[:-1]) / 100.0
                return percent * max_dimension
            else:
                return float(value)
        return float(value)


class Size:
    """Flexible size system supporting multiple input formats."""
    
    @staticmethod
    def from_dict(
        size_dict: Dict[str, Any],
        slide_width: float = SLIDE_WIDTH_INCHES,
        slide_height: float = SLIDE_HEIGHT_INCHES,
        aspect_ratio: Optional[float] = None
    ) -> Tuple[Optional[float], Optional[float]]:
        """
        Convert size dict to (width, height) in inches.
        
        Supports:
        - {"width": 5.0, "height": 3.0}  # Absolute inches
        - {"width": "50%", "height": "30%"}  # Percentage of slide
        - {"width": "auto", "height": 3.0}  # Maintain aspect ratio
        - {"width": 5.0, "height": "auto"}  # Maintain aspect ratio
        
        Args:
            size_dict: Size specification dictionary
            slide_width: Slide width in inches
            slide_height: Slide height in inches
            aspect_ratio: Optional aspect ratio (width/height) for "auto" calculations
            
        Returns:
            Tuple of (width, height) in inches, either can be None for "auto"
        """
        if not isinstance(size_dict, dict):
            raise ValueError(f"Size must be a dictionary, got {type(size_dict).__name__}")
        
        if "width" not in size_dict and "height" not in size_dict:
            raise ValueError("Size must have at least 'width' or 'height'")
        
        width_spec = size_dict.get("width")
        height_spec = size_dict.get("height")
        
        # Parse width
        if width_spec == "auto" or width_spec is None:
            width = None
        else:
            width = Position._parse_dimension(width_spec, slide_width)
        
        # Parse height
        if height_spec == "auto" or height_spec is None:
            height = None
        else:
            height = Position._parse_dimension(height_spec, slide_height)
        
        # Apply aspect ratio if one dimension is auto
        if aspect_ratio is not None:
            if width is None and height is not None:
                width = height * aspect_ratio
            elif height is None and width is not None:
                height = width / aspect_ratio
        
        return (width, height)


class ColorHelper:
    """Utilities for color conversion and validation."""
    
    @staticmethod
    def from_hex(hex_color: str) -> RGBColor:
        """
        Convert hex color string to RGBColor.
        
        Args:
            hex_color: Hex color string (e.g., "#FF0000" or "FF0000")
            
        Returns:
            RGBColor object
            
        Raises:
            ValueError: If hex color format is invalid
        """
        hex_color = hex_color.strip().lstrip('#')
        
        if len(hex_color) != 6:
            raise ValueError(f"Invalid hex color: {hex_color}. Must be 6 hex digits.")
        
        if not all(c in '0123456789ABCDEFabcdef' for c in hex_color):
            raise ValueError(f"Invalid hex color: {hex_color}. Contains non-hex characters.")
        
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        
        return RGBColor(r, g, b)
    
    @staticmethod
    def to_hex(rgb_color: RGBColor) -> str:
        """
        Convert RGBColor to hex string.
        
        Args:
            rgb_color: RGBColor object
            
        Returns:
            Hex color string with # prefix
        """
        if hasattr(rgb_color, '__iter__') and len(rgb_color) == 3:
            r, g, b = rgb_color
        elif hasattr(rgb_color, 'r'):
            r, g, b = rgb_color.r, rgb_color.g, rgb_color.b
        else:
            # Handle string representation
            hex_str = str(rgb_color).lstrip('#')
            return f"#{hex_str}"
        
        return f"#{r:02x}{g:02x}{b:02x}"
    
    @staticmethod
    def luminance(rgb_color: Union[RGBColor, Tuple[int, int, int]]) -> float:
        """
        Calculate relative luminance for WCAG contrast calculations.
        
        Args:
            rgb_color: RGBColor or (r, g, b) tuple
            
        Returns:
            Relative luminance value (0.0 to 1.0)
        """
        # Extract RGB values
        if hasattr(rgb_color, 'r'):
            r, g, b = rgb_color.r, rgb_color.g, rgb_color.b
        elif hasattr(rgb_color, '__iter__'):
            r, g, b = rgb_color
        else:
            # Handle string representation
            hex_str = str(rgb_color).lstrip('#')
            if len(hex_str) == 6:
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
            else:
                raise ValueError(f"Cannot parse color: {rgb_color}")
        
        def _linearize(channel: int) -> float:
            c = channel / 255.0
            if c <= 0.03928:
                return c / 12.92
            return ((c + 0.055) / 1.055) ** 2.4
        
        r_lin = _linearize(r)
        g_lin = _linearize(g)
        b_lin = _linearize(b)
        
        return 0.2126 * r_lin + 0.7152 * g_lin + 0.0722 * b_lin
    
    @staticmethod
    def contrast_ratio(color1: RGBColor, color2: RGBColor) -> float:
        """
        Calculate WCAG contrast ratio between two colors.
        
        Args:
            color1: First color
            color2: Second color
            
        Returns:
            Contrast ratio (1.0 to 21.0)
        """
        lum1 = ColorHelper.luminance(color1)
        lum2 = ColorHelper.luminance(color2)
        
        lighter = max(lum1, lum2)
        darker = min(lum1, lum2)
        
        return (lighter + 0.05) / (darker + 0.05)
    
    @staticmethod
    def meets_wcag(
        foreground: RGBColor,
        background: RGBColor,
        is_large_text: bool = False
    ) -> bool:
        """
        Check if color combination meets WCAG 2.1 AA standards.
        
        Args:
            foreground: Text/foreground color
            background: Background color
            is_large_text: True if text is 18pt+ or 14pt+ bold
            
        Returns:
            True if contrast is sufficient
        """
        ratio = ColorHelper.contrast_ratio(foreground, background)
        threshold = WCAG_CONTRAST_LARGE if is_large_text else WCAG_CONTRAST_NORMAL
        return ratio >= threshold


# ============================================================================
# ANALYSIS CLASSES
# ============================================================================

class TemplateProfile:
    """
    Captures and provides access to PowerPoint template formatting.
    
    Uses lazy loading to avoid performance penalty when profile is not needed.
    """
    
    def __init__(self, prs: Optional['Presentation'] = None):
        """
        Initialize template profile.
        
        Args:
            prs: Optional Presentation to analyze immediately
        """
        self._prs = prs
        self._captured = False
        self._slide_layouts: List[Dict[str, Any]] = []
        self._theme_colors: Dict[str, str] = {}
        self._theme_fonts: Dict[str, str] = {}
    
    def _ensure_captured(self) -> None:
        """Ensure template data has been captured (lazy loading)."""
        if self._captured or self._prs is None:
            return
        
        self._capture_layouts()
        self._capture_theme()
        self._captured = True
    
    def _capture_layouts(self) -> None:
        """Capture layout information from presentation."""
        for layout in self._prs.slide_layouts:
            layout_info = {
                "name": layout.name,
                "placeholders": []
            }
            
            for ph in layout.placeholders:
                try:
                    ph_info = {
                        "type": _get_placeholder_type_int_helper(ph.placeholder_format.type),
                        "idx": ph.placeholder_format.idx
                    }
                    if hasattr(ph, 'left') and ph.left is not None:
                        ph_info["position"] = {
                            "left": ph.left / EMU_PER_INCH,
                            "top": ph.top / EMU_PER_INCH
                        }
                    if hasattr(ph, 'width') and ph.width is not None:
                        ph_info["size"] = {
                            "width": ph.width / EMU_PER_INCH,
                            "height": ph.height / EMU_PER_INCH
                        }
                    layout_info["placeholders"].append(ph_info)
                except Exception:
                    continue
            
            self._slide_layouts.append(layout_info)
    
    def _capture_theme(self) -> None:
        """Capture theme colors and fonts from presentation."""
        try:
            # Attempt to extract theme colors
            if hasattr(self._prs, 'slide_master') and self._prs.slide_master:
                master = self._prs.slide_master
                
                # Extract fonts from shapes
                for shape in master.shapes:
                    if hasattr(shape, 'text_frame'):
                        try:
                            for para in shape.text_frame.paragraphs:
                                if para.font.name:
                                    font_key = f"font_{len(self._theme_fonts)}"
                                    if para.font.name not in self._theme_fonts.values():
                                        self._theme_fonts[font_key] = para.font.name
                        except Exception:
                            continue
        except Exception:
            pass
    
    @property
    def slide_layouts(self) -> List[Dict[str, Any]]:
        """Get slide layout information."""
        self._ensure_captured()
        return self._slide_layouts
    
    @property
    def theme_colors(self) -> Dict[str, str]:
        """Get theme colors."""
        self._ensure_captured()
        return self._theme_colors
    
    @property
    def theme_fonts(self) -> Dict[str, str]:
        """Get theme fonts."""
        self._ensure_captured()
        return self._theme_fonts
    
    def get_layout_names(self) -> List[str]:
        """Get list of available layout names."""
        self._ensure_captured()
        return [layout["name"] for layout in self._slide_layouts]
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert profile to JSON-serializable dict."""
        self._ensure_captured()
        return {
            "slide_layouts": self._slide_layouts,
            "theme_colors": self._theme_colors,
            "theme_fonts": self._theme_fonts
        }


class AccessibilityChecker:
    """WCAG 2.1 compliance checker for presentations."""
    
    @staticmethod
    def check_presentation(prs: 'Presentation') -> Dict[str, Any]:
        """
        Comprehensive accessibility check.
        
        Args:
            prs: Presentation to check
            
        Returns:
            Dict containing:
            - status: "accessible" or "issues_found"
            - total_issues: Count of all issues
            - issues: Detailed issue breakdown
            - wcag_level: "AA" if passing, "fail" otherwise
        """
        issues = {
            "missing_alt_text": [],
            "low_contrast": [],
            "missing_titles": [],
            "small_text": [],
            "reading_order_warnings": []
        }
        
        for slide_idx, slide in enumerate(prs.slides):
            # Check for title
            has_title = AccessibilityChecker._check_slide_has_title(slide)
            if not has_title:
                issues["missing_titles"].append({
                    "slide": slide_idx,
                    "message": "Slide lacks a title for screen reader navigation"
                })
            
            # Check each shape
            for shape_idx, shape in enumerate(slide.shapes):
                # Check images for alt text
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    if not AccessibilityChecker._has_alt_text(shape):
                        issues["missing_alt_text"].append({
                            "slide": slide_idx,
                            "shape": shape_idx,
                            "shape_name": shape.name,
                            "message": "Image lacks alternative text"
                        })
                
                # Check text for contrast and size
                if hasattr(shape, 'text_frame') and shape.has_text_frame:
                    AccessibilityChecker._check_text_accessibility(
                        shape, slide_idx, shape_idx, issues
                    )
        
        total_issues = sum(len(v) for v in issues.values())
        
        return {
            "status": "issues_found" if total_issues > 0 else "accessible",
            "total_issues": total_issues,
            "issues": issues,
            "wcag_level": "AA" if total_issues == 0 else "fail",
            "checked_slides": len(prs.slides)
        }
    
    @staticmethod
    def _check_slide_has_title(slide) -> bool:
        """Check if slide has a non-empty title."""
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = _get_placeholder_type_int_helper(
                    shape.placeholder_format.type
                )
                if ph_type in TITLE_PLACEHOLDER_TYPES:
                    if shape.has_text_frame and shape.text_frame.text.strip():
                        return True
        return False
    
    @staticmethod
    def _has_alt_text(shape) -> bool:
        """
        Check if image shape has meaningful alt text.
        
        Checks both the description attribute (proper alt text)
        and the shape name as fallback.
        """
        # Check description attribute (the actual alt text storage)
        try:
            element = shape._element
            # Check for description in various possible locations
            descr = element.get('descr')
            if descr and descr.strip() and len(descr.strip()) > 3:
                return True
            
            # Check nvPicPr/cNvPr for descr
            for child in element.iter():
                if child.get('descr'):
                    descr = child.get('descr')
                    if descr and descr.strip() and len(descr.strip()) > 3:
                        return True
        except Exception:
            pass
        
        # Fallback: check name (not ideal, but some tools use this)
        if shape.name:
            name = shape.name.strip()
            # Reject generic names
            if name.lower().startswith('picture'):
                return False
            if name.lower().startswith('image'):
                return False
            if len(name) > 5:  # Meaningful name
                return True
        
        return False
    
    @staticmethod
    def _check_text_accessibility(
        shape,
        slide_idx: int,
        shape_idx: int,
        issues: Dict[str, Any]
    ) -> None:
        """Check text shape for accessibility issues."""
        try:
            text_frame = shape.text_frame
            for para in text_frame.paragraphs:
                # Check font size
                if para.font.size is not None:
                    size_pt = para.font.size.pt
                    if size_pt < 10:
                        issues["small_text"].append({
                            "slide": slide_idx,
                            "shape": shape_idx,
                            "size_pt": size_pt,
                            "text_preview": para.text[:50] if para.text else "",
                            "message": f"Text size {size_pt}pt is below minimum 10pt"
                        })
        except Exception:
            pass


class AssetValidator:
    """Validates and provides information about presentation assets."""
    
    @staticmethod
    def validate_presentation_assets(
        prs: 'Presentation',
        filepath: Optional[Path] = None
    ) -> Dict[str, Any]:
        """
        Validate all assets in presentation.
        
        Args:
            prs: Presentation to validate
            filepath: Optional file path for size check
            
        Returns:
            Validation report dict
        """
        issues = {
            "large_images": [],
            "total_embedded_size_bytes": 0,
            "image_count": 0
        }
        
        for slide_idx, slide in enumerate(prs.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    issues["image_count"] += 1
                    try:
                        image_blob = shape.image.blob
                        image_size = len(image_blob)
                        issues["total_embedded_size_bytes"] += image_size
                        
                        # Flag images over 2MB
                        if image_size > 2 * 1024 * 1024:
                            issues["large_images"].append({
                                "slide": slide_idx,
                                "shape": shape_idx,
                                "size_bytes": image_size,
                                "size_mb": round(image_size / (1024 * 1024), 2)
                            })
                    except Exception:
                        pass
        
        # Check total file size
        if filepath and Path(filepath).exists():
            file_size = Path(filepath).stat().st_size
            issues["file_size_bytes"] = file_size
            issues["file_size_mb"] = round(file_size / (1024 * 1024), 2)
            
            if file_size > MAX_RECOMMENDED_FILE_SIZE_MB * 1024 * 1024:
                issues["large_file_warning"] = {
                    "size_mb": issues["file_size_mb"],
                    "recommended_max_mb": MAX_RECOMMENDED_FILE_SIZE_MB
                }
        
        total_issues = len(issues["large_images"])
        if "large_file_warning" in issues:
            total_issues += 1
        
        return {
            "status": "issues_found" if total_issues > 0 else "valid",
            "total_issues": total_issues,
            "issues": issues
        }
    
    @staticmethod
    def compress_image(
        image_path: Path,
        max_width: int = 1920,
        quality: int = 85
    ) -> BytesIO:
        """
        Compress image for PowerPoint embedding.
        
        Args:
            image_path: Path to source image
            max_width: Maximum width in pixels
            quality: JPEG quality (1-100)
            
        Returns:
            BytesIO containing compressed image
            
        Raises:
            ImportError: If Pillow is not available
        """
        if not HAS_PILLOW:
            raise ImportError("Pillow is required for image compression")
        
        with PILImage.open(image_path) as img:
            # Resize if needed
            if img.width > max_width:
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), PILImage.LANCZOS)
            
            # Convert to RGB if necessary
            if img.mode in ('RGBA', 'LA', 'P'):
                background = PILImage.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                if img.mode in ('RGBA', 'LA'):
                    background.paste(img, mask=img.split()[-1])
                else:
                    background.paste(img)
                img = background
            
            # Save to BytesIO
            output = BytesIO()
            img.save(output, format='JPEG', quality=quality, optimize=True)
            output.seek(0)
            
            return output


# ============================================================================
# MAIN POWERPOINT AGENT CLASS
# ============================================================================

class PowerPointAgent:
    """
    Core PowerPoint manipulation class for stateless tool operations.
    
    Provides comprehensive PowerPoint editing capabilities optimized for
    AI agent consumption through simple, composable operations.
    
    Features:
    - Stateless design for tool-based workflows
    - Comprehensive validation and accessibility checking
    - Atomic file locking for concurrent access safety
    - Full alignment with Presentation Architect System Prompt v3.0
    - Approval token governance for destructive operations
    - Geometry-aware version tracking for state detection
    
    Example:
        with PowerPointAgent() as agent:
            agent.open(Path("presentation.pptx"))
            agent.add_slide("Title and Content")
            agent.save()
    """
    
    def __init__(self, filepath: Optional[Union[str, Path]] = None):
        """
        Initialize PowerPoint agent.
        
        Args:
            filepath: Optional path to open immediately
        """
        self.filepath: Optional[Path] = None
        self.prs: Optional[Presentation] = None
        self._lock: Optional[FileLock] = None
        self._template_profile: Optional[TemplateProfile] = None
        self._layout_cache: Optional[Dict[str, Any]] = None
        
        if filepath:
            self.filepath = Path(filepath)
    
    # ========================================================================
    # CONTEXT MANAGEMENT
    # ========================================================================
    
    def __enter__(self) -> 'PowerPointAgent':
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        self.close()
        return False
    
    # ========================================================================
    # HELPER METHODS (Governance & Observability)
    # ========================================================================

    def _validate_token(self, token: Optional[str], scope: str) -> None:
        """
        Validate approval token for destructive operations.
        
        Args:
            token: The approval token string
            scope: The required permission scope (e.g., "delete:slide")
            
        Raises:
            ApprovalTokenError: If token is missing or invalid
        """
        # NOTE: In a production environment, this would verify a JWT or HMAC.
        # For this implementation, we check presence and basic format.
        if not token:
            raise ApprovalTokenError(
                f"Destructive operation requires approval token (scope: {scope})",
                details={"scope_required": scope}
            )
        
        # Placeholder validation - real implementation would check signature
        if len(token) < 8:
            raise ApprovalTokenError(
                "Invalid approval token format",
                details={"token_length": len(token)}
            )

    def _capture_version(self) -> str:
        """Capture current presentation version hash."""
        return self.get_presentation_version()

    def _log_warning(self, message: str) -> None:
        """Log a warning message through the configured logger."""
        logger.warning(message)

    # ========================================================================
    # FILE OPERATIONS
    # ========================================================================
    
    def create_new(self, template: Optional[Union[str, Path]] = None) -> None:
        """
        Create new presentation, optionally from template.
        
        Args:
            template: Optional path to template .pptx file
            
        Raises:
            FileNotFoundError: If template doesn't exist
            TemplateError: If template cannot be loaded
        """
        if template:
            template_path = PathValidator.validate_pptx_path(template, must_exist=True)
            try:
                self.prs = Presentation(str(template_path))
            except Exception as e:
                raise TemplateError(
                    f"Failed to load template: {template_path}",
                    details={"error": str(e)}
                )
        else:
            self.prs = Presentation()
        
        self._template_profile = TemplateProfile(self.prs)
        self._layout_cache = None
    
    def open(
        self,
        filepath: Union[str, Path],
        acquire_lock: bool = True
    ) -> None:
        """
        Open existing presentation.
        
        Args:
            filepath: Path to .pptx file
            acquire_lock: Whether to acquire exclusive file lock
            
        Raises:
            PathValidationError: If path is invalid
            FileLockError: If lock cannot be acquired
            PowerPointAgentError: If file cannot be opened
        """
        validated_path = PathValidator.validate_pptx_path(filepath, must_exist=True)
        self.filepath = validated_path
        
        # Acquire lock if requested
        if acquire_lock:
            self._lock = FileLock(validated_path)
            if not self._lock.acquire():
                raise FileLockError(
                    f"Could not acquire lock on {validated_path}",
                    details={"filepath": str(validated_path)}
                )
        
        # Load presentation (with lock release on failure)
        try:
            self.prs = Presentation(str(validated_path))
            self._template_profile = TemplateProfile(self.prs)
            self._layout_cache = None
        except Exception as e:
            # Release lock on failure
            if self._lock:
                self._lock.release()
                self._lock = None
            raise PowerPointAgentError(
                f"Failed to open presentation: {validated_path}",
                details={"error": str(e)}
            )
    
    def save(self, filepath: Optional[Union[str, Path]] = None) -> None:
        """
        Save presentation.
        
        Args:
            filepath: Output path (uses original path if None)
            
        Raises:
            PowerPointAgentError: If no presentation loaded
            PathValidationError: If output path is invalid
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        target = filepath or self.filepath
        if not target:
            raise PowerPointAgentError("No output path specified")
        
        target_path = PathValidator.validate_pptx_path(
            target,
            must_exist=False,
            must_be_writable=True
        )
        
        # Ensure parent directory exists
        target_path.parent.mkdir(parents=True, exist_ok=True)
        
        self.prs.save(str(target_path))
        self.filepath = target_path
    
    def close(self) -> None:
        """Close presentation and release resources."""
        self.prs = None
        self._template_profile = None
        self._layout_cache = None
        
        if self._lock:
            self._lock.release()
            self._lock = None
    
    def clone_presentation(self, output_path: Union[str, Path]) -> 'PowerPointAgent':
        """
        Clone current presentation to a new file.
        
        Args:
            output_path: Path for the cloned presentation
            
        Returns:
            New PowerPointAgent instance with cloned presentation
            
        Raises:
            PowerPointAgentError: If no presentation loaded
            PathValidationError: If output path is invalid
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        output = PathValidator.validate_pptx_path(
            output_path,
            must_exist=False,
            must_be_writable=True
        )
        
        # Save to new location
        output.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(output))
        
        # Create new agent with cloned file
        new_agent = PowerPointAgent()
        new_agent.open(output)
        
        return new_agent
    
    # ========================================================================
    # SLIDE OPERATIONS
    # ========================================================================
    
    def add_slide(
        self,
        layout_name: str = "Title and Content",
        index: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        Add new slide with specified layout.
        
        Args:
            layout_name: Name of layout to use
            index: Position to insert (None = append at end)
            
        Returns:
            Dict with slide_index and layout_name
            
        Raises:
            PowerPointAgentError: If no presentation loaded
            LayoutNotFoundError: If layout doesn't exist
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        version_before = self._capture_version()
        
        layout = self._get_layout(layout_name)
        slide = self.prs.slides.add_slide(layout)
        
        result_index = len(self.prs.slides) - 1
        
        if index is not None:
            max_valid = len(self.prs.slides)
            if not 0 <= index <= max_valid:
                 raise SlideNotFoundError(
                    f"Insert index {index} out of range (0-{max_valid})",
                    details={"index": index, "valid_range": f"0-{max_valid}"}
                )
            
            # Move slide from end to target position
            xml_slides = self.prs.slides._sldIdLst
            slide_elem = xml_slides[-1]
            xml_slides.remove(slide_elem)
            xml_slides.insert(index, slide_elem)
            result_index = index
        
        version_after = self._capture_version()
        
        return {
            "slide_index": result_index,
            "layout_name": layout_name,
            "total_slides": len(self.prs.slides),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def delete_slide(
        self,
        index: int,
        approval_token: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Delete slide at index.
        
        ⚠️ DESTRUCTIVE OPERATION - Requires approval token.
        
        Args:
            index: Slide index (0-based)
            approval_token: Token authorizing destructive operation
            
        Returns:
            Dict with deleted index and new slide count
            
        Raises:
            SlideNotFoundError: If index is out of range
            ApprovalTokenError: If token is missing/invalid
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        self._validate_token(approval_token, APPROVAL_SCOPE_DELETE_SLIDE)
        
        slide_count = len(self.prs.slides)
        if not 0 <= index < slide_count:
            raise SlideNotFoundError(
                f"Slide index {index} out of range",
                details={"index": index, "slide_count": slide_count}
            )
        
        version_before = self._capture_version()
        
        # Get slide relationship ID and remove
        rId = self.prs.slides._sldIdLst[index].rId
        self.prs.part.drop_rel(rId)
        del self.prs.slides._sldIdLst[index]
        
        version_after = self._capture_version()
        
        return {
            "deleted_index": index,
            "previous_count": slide_count,
            "new_count": len(self.prs.slides),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def duplicate_slide(self, index: int) -> Dict[str, Any]:
        """
        Duplicate slide at index.
        
        Args:
            index: Slide index to duplicate
            
        Returns:
            Dict with new slide index
            
        Raises:
            SlideNotFoundError: If index is out of range
        """
        source_slide = self._get_slide(index)
        version_before = self._capture_version()
        
        # Add new slide with same layout
        layout = source_slide.slide_layout
        new_slide = self.prs.slides.add_slide(layout)
        new_index = len(self.prs.slides) - 1
        
        # Copy shapes
        for shape in source_slide.shapes:
            try:
                self._copy_shape(shape, new_slide)
            except Exception as e:
                logger.warning(f"Could not copy shape: {e}")
        
        version_after = self._capture_version()
        
        return {
            "source_index": index,
            "new_index": new_index,
            "total_slides": len(self.prs.slides),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def reorder_slides(self, from_index: int, to_index: int) -> Dict[str, Any]:
        """
        Move slide from one position to another.
        
        Args:
            from_index: Current position
            to_index: Desired position
            
        Returns:
            Dict with movement details
            
        Raises:
            SlideNotFoundError: If either index is out of range
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        slide_count = len(self.prs.slides)
        
        if not 0 <= from_index < slide_count:
            raise SlideNotFoundError(
                f"Source index {from_index} out of range",
                details={"from_index": from_index, "slide_count": slide_count}
            )
        
        if not 0 <= to_index < slide_count:
            raise SlideNotFoundError(
                f"Target index {to_index} out of range",
                details={"to_index": to_index, "slide_count": slide_count}
            )
        
        version_before = self._capture_version()
        
        xml_slides = self.prs.slides._sldIdLst
        slide_elem = xml_slides[from_index]
        xml_slides.remove(slide_elem)
        xml_slides.insert(to_index, slide_elem)
        
        version_after = self._capture_version()
        
        return {
            "from_index": from_index,
            "to_index": to_index,
            "total_slides": slide_count,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def get_slide_count(self) -> int:
        """
        Get total number of slides.
        
        Returns:
            Number of slides
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        return len(self.prs.slides)
    
    # ========================================================================
    # TEXT OPERATIONS
    # ========================================================================
    
    def add_text_box(
        self,
        slide_index: int,
        text: str,
        position: Dict[str, Any],
        size: Dict[str, Any],
        font_name: Optional[str] = None,
        font_size: int = 18,
        bold: bool = False,
        italic: bool = False,
        color: Optional[str] = None,
        alignment: str = "left"
    ) -> Dict[str, Any]:
        """
        Add text box to slide.
        
        Args:
            slide_index: Target slide index
            text: Text content
            position: Position dict (see Position.from_dict)
            size: Size dict (see Size.from_dict)
            font_name: Font name (None uses theme font)
            font_size: Font size in points
            bold: Bold text
            italic: Italic text
            color: Text color hex (e.g., "#FF0000")
            alignment: Text alignment ("left", "center", "right", "justify")
            
        Returns:
            Dict with shape_index and details
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            InvalidPositionError: If position is invalid
        """
        slide = self._get_slide(slide_index)
        version_before = self._capture_version()
        
        # Parse position and size
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Text box must have explicit width and height")
        
        # Create text box
        text_box = slide.shapes.add_textbox(
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        # Configure text frame
        text_frame = text_box.text_frame
        text_frame.text = text
        text_frame.word_wrap = True
        
        # Apply formatting
        paragraph = text_frame.paragraphs[0]
        if font_name:
            paragraph.font.name = font_name
        paragraph.font.size = Pt(font_size)
        paragraph.font.bold = bold
        paragraph.font.italic = italic
        
        if color:
            paragraph.font.color.rgb = ColorHelper.from_hex(color)
        
        # Set alignment
        alignment_map = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            "justify": PP_ALIGN.JUSTIFY
        }
        paragraph.alignment = alignment_map.get(alignment.lower(), PP_ALIGN.LEFT)
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": len(slide.shapes) - 1,
            "text_length": len(text),
            "position": {"left": left, "top": top},
            "size": {"width": width, "height": height},
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def set_title(
        self,
        slide_index: int,
        title: str,
        subtitle: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Set slide title and optional subtitle.
        
        Args:
            slide_index: Target slide index
            title: Title text
            subtitle: Optional subtitle text
            
        Returns:
            Dict with title/subtitle set status
            
        Raises:
            SlideNotFoundError: If slide index is invalid
        """
        slide = self._get_slide(slide_index)
        version_before = self._capture_version()
        
        title_set = False
        subtitle_set = False
        title_shape_index = None
        subtitle_shape_index = None
        
        for idx, shape in enumerate(slide.shapes):
            if shape.is_placeholder:
                ph_type = _get_placeholder_type_int_helper(shape.placeholder_format.type)
                
                # Check for title placeholder
                if ph_type in TITLE_PLACEHOLDER_TYPES:
                    if shape.has_text_frame:
                        shape.text_frame.text = title
                        title_set = True
                        title_shape_index = idx
                
                # Check for subtitle placeholder
                elif ph_type == SUBTITLE_PLACEHOLDER_TYPE:
                    if subtitle and shape.has_text_frame:
                        shape.text_frame.text = subtitle
                        subtitle_set = True
                        subtitle_shape_index = idx
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "title_set": title_set,
            "subtitle_set": subtitle_set,
            "title_shape_index": title_shape_index,
            "subtitle_shape_index": subtitle_shape_index,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def add_bullet_list(
        self,
        slide_index: int,
        items: List[str],
        position: Dict[str, Any],
        size: Dict[str, Any],
        bullet_style: str = "bullet",
        font_size: int = 18,
        font_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Add bullet list to slide.
        
        Args:
            slide_index: Target slide index
            items: List of bullet items
            position: Position dict
            size: Size dict
            bullet_style: "bullet", "numbered", or "none"
            font_size: Font size in points
            font_name: Optional font name
            
        Returns:
            Dict with shape_index and item count
        """
        slide = self._get_slide(slide_index)
        version_before = self._capture_version()
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Bullet list must have explicit width and height")
        
        # Create text box for bullets
        text_box = slide.shapes.add_textbox(
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        for idx, item in enumerate(items):
            if idx == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            if bullet_style == "numbered":
                p.text = f"{idx + 1}. {item}"
            else:
                p.text = item
            
            p.level = 0
            p.font.size = Pt(font_size)
            if font_name:
                p.font.name = font_name
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": len(slide.shapes) - 1,
            "item_count": len(items),
            "bullet_style": bullet_style,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def format_text(
        self,
        slide_index: int,
        shape_index: int,
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Format existing text shape.
        
        Args:
            slide_index: Target slide index
            shape_index: Shape index on slide
            font_name: Optional font name
            font_size: Optional font size in points
            bold: Optional bold setting
            italic: Optional italic setting
            color: Optional color hex
            
        Returns:
            Dict with formatting applied
        """
        shape = self._get_shape(slide_index, shape_index)
        
        if not hasattr(shape, 'text_frame') or not shape.has_text_frame:
            raise ValueError(f"Shape at index {shape_index} does not have text")
            
        version_before = self._capture_version()
        changes = []
        
        for paragraph in shape.text_frame.paragraphs:
            if font_name is not None:
                paragraph.font.name = font_name
                changes.append("font_name")
            if font_size is not None:
                paragraph.font.size = Pt(font_size)
                changes.append("font_size")
            if bold is not None:
                paragraph.font.bold = bold
                changes.append("bold")
            if italic is not None:
                paragraph.font.italic = italic
                changes.append("italic")
            if color is not None:
                paragraph.font.color.rgb = ColorHelper.from_hex(color)
                changes.append("color")
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "changes_applied": list(set(changes)),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def replace_text(
        self,
        find: str,
        replace: str,
        slide_index: Optional[int] = None,
        shape_index: Optional[int] = None,
        match_case: bool = False
    ) -> Dict[str, Any]:
        """
        Find and replace text in presentation.
        
        Args:
            find: Text to find
            replace: Replacement text
            slide_index: Optional specific slide (None = all slides)
            shape_index: Optional specific shape (requires slide_index)
            match_case: Case-sensitive matching
            
        Returns:
            Dict with replacement count and locations
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        if shape_index is not None and slide_index is None:
            raise ValueError("shape_index requires slide_index to be specified")
        
        version_before = self._capture_version()
        
        replacements = []
        total_count = 0
        
        # Determine slides to process
        if slide_index is not None:
            slides_to_process = [(slide_index, self._get_slide(slide_index))]
        else:
            slides_to_process = list(enumerate(self.prs.slides))
        
        for s_idx, slide in slides_to_process:
            # Determine shapes to process
            if shape_index is not None:
                shapes_to_process = [(shape_index, self._get_shape(s_idx, shape_index))]
            else:
                shapes_to_process = list(enumerate(slide.shapes))
            
            for sh_idx, shape in shapes_to_process:
                if not hasattr(shape, 'text_frame') or not shape.has_text_frame:
                    continue
                
                count = self._replace_text_in_shape(shape, find, replace, match_case)
                if count > 0:
                    total_count += count
                    replacements.append({
                        "slide": s_idx,
                        "shape": sh_idx,
                        "count": count
                    })
        
        version_after = self._capture_version()
        
        return {
            "find": find,
            "replace": replace,
            "match_case": match_case,
            "total_replacements": total_count,
            "locations": replacements,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def _replace_text_in_shape(
        self,
        shape,
        find: str,
        replace: str,
        match_case: bool
    ) -> int:
        """Replace text within a single shape, preserving formatting where possible."""
        count = 0
        
        try:
            text_frame = shape.text_frame
        except (AttributeError, TypeError):
            return 0
        
        # Strategy 1: Replace in runs (preserves formatting)
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                if match_case:
                    if find in run.text:
                        occurrences = run.text.count(find)
                        run.text = run.text.replace(find, replace)
                        count += occurrences
                else:
                    if find.lower() in run.text.lower():
                        pattern = re.compile(re.escape(find), re.IGNORECASE)
                        matches = pattern.findall(run.text)
                        run.text = pattern.sub(replace, run.text)
                        count += len(matches)
        
        if count > 0:
            return count
        
        # Strategy 2: Full text replacement (if text spans runs)
        try:
            full_text = shape.text
            if not full_text:
                return 0
            
            if match_case:
                if find in full_text:
                    occurrences = full_text.count(find)
                    shape.text = full_text.replace(find, replace)
                    return occurrences
            else:
                if find.lower() in full_text.lower():
                    pattern = re.compile(re.escape(find), re.IGNORECASE)
                    matches = pattern.findall(full_text)
                    shape.text = pattern.sub(replace, full_text)
                    return len(matches)
        except (AttributeError, TypeError):
            pass
        
        return 0
    
    def add_notes(
        self,
        slide_index: int,
        text: str,
        mode: Union[str, NotesMode] = NotesMode.APPEND
    ) -> Dict[str, Any]:
        """
        Add speaker notes to a slide.
        
        Args:
            slide_index: Target slide index
            text: Notes text to add
            mode: "append", "prepend", or "overwrite"
            
        Returns:
            Dict with notes details
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ValueError: If mode is invalid
        """
        if isinstance(mode, str):
            try:
                mode = NotesMode(mode.lower())
            except ValueError:
                raise ValueError(f"Invalid mode: {mode}")

        slide = self._get_slide(slide_index)
        version_before = self._capture_version()
        
        # Access or create notes slide
        notes_slide = slide.notes_slide
        text_frame = notes_slide.notes_text_frame
        
        original_text = text_frame.text or ""
        original_length = len(original_text)
        
        if mode == NotesMode.OVERWRITE:
            final_text = text
        elif mode == NotesMode.APPEND:
            if original_text.strip():
                final_text = original_text + "\n" + text
            else:
                final_text = text
        elif mode == NotesMode.PREPEND:
            if original_text.strip():
                final_text = text + "\n" + original_text
            else:
                final_text = text
        
        text_frame.text = final_text
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "mode": mode.value,
            "original_length": original_length,
            "new_length": len(final_text),
            "text_preview": final_text[:100] + "..." if len(final_text) > 100 else final_text,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def set_footer(
        self,
        text: Optional[str] = None,
        show_slide_number: bool = False,
        show_date: bool = False,
        slide_index: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        Set footer properties for slide(s).
        
        Note: Footer configuration in python-pptx is limited.
        This method sets footer placeholders where available.
        
        Args:
            text: Footer text
            show_slide_number: Show slide numbers
            show_date: Show date
            slide_index: Specific slide (None = all slides)
            
        Returns:
            Dict with footer configuration results
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        version_before = self._capture_version()
        results = []
        
        # Determine slides to process
        if slide_index is not None:
            slides = [(slide_index, self._get_slide(slide_index))]
        else:
            slides = list(enumerate(self.prs.slides))
        
        for s_idx, slide in slides:
            slide_result = {
                "slide_index": s_idx,
                "footer_set": False,
                "slide_number_set": False,
                "date_set": False
            }
            
            for shape in slide.shapes:
                if not shape.is_placeholder:
                    continue
                
                ph_type = _get_placeholder_type_int_helper(shape.placeholder_format.type)
                
                # Footer placeholder (type 7)
                if ph_type == 7 and text is not None:
                    if shape.has_text_frame:
                        shape.text_frame.text = text
                        slide_result["footer_set"] = True
                
                # Slide number placeholder (type 6)
                if ph_type == 6 and show_slide_number:
                    slide_result["slide_number_set"] = True
                
                # Date placeholder (type 5)
                if ph_type == 5 and show_date:
                    slide_result["date_set"] = True
            
            results.append(slide_result)
        
        version_after = self._capture_version()
        
        return {
            "text": text,
            "show_slide_number": show_slide_number,
            "show_date": show_date,
            "slides_processed": len(results),
            "results": results,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    # ========================================================================
    # SHAPE OPERATIONS
    # ========================================================================
    
    def _set_fill_opacity(self, shape, opacity: float) -> bool:
        """
        Set the fill opacity of a shape by manipulating the underlying XML.
        
        Args:
            shape: The shape object with a fill
            opacity: Opacity value (0.0 = fully transparent, 1.0 = fully opaque)
            
        Returns:
            True if opacity was set, False if not applicable
            
        Note:
            python-pptx doesn't directly expose fill transparency, so we
            manipulate the OOXML directly. The alpha value uses a scale
            where 100000 = 100% opaque.
        """
        if opacity >= 1.0:
            # No need to set alpha for fully opaque - it's the default
            return True
        
        if opacity < 0.0:
            opacity = 0.0
        
        try:
            # Access the shape's spPr (shape properties) element
            spPr = shape._sp.spPr
            if spPr is None:
                return False
            
            # Find the solidFill element
            solidFill = spPr.find(qn('a:solidFill'))
            if solidFill is None:
                return False
            
            # Find the color element (could be srgbClr or schemeClr)
            color_elem = solidFill.find(qn('a:srgbClr'))
            if color_elem is None:
                color_elem = solidFill.find(qn('a:schemeClr'))
            if color_elem is None:
                return False
            
            # Calculate alpha value (Office uses 0-100000 scale, where 100000 = 100%)
            alpha_value = int(opacity * 100000)
            
            # Remove existing alpha element if present
            existing_alpha = color_elem.find(qn('a:alpha'))
            if existing_alpha is not None:
                color_elem.remove(existing_alpha)
            
            # Create and add new alpha element
            # Using SubElement to create properly namespaced element
            nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            alpha_elem = etree.SubElement(color_elem, qn('a:alpha'))
            alpha_elem.set('val', str(alpha_value))
            
            return True
            
        except Exception as e:
            # Log but don't fail - opacity is enhancement, not critical
            self._log_warning(f"Could not set fill opacity: {e}")
            return False
    
    def _set_line_opacity(self, shape, opacity: float) -> bool:
        """
        Set the line/border opacity of a shape by manipulating the underlying XML.
        
        Args:
            shape: The shape object with a line
            opacity: Opacity value (0.0 = fully transparent, 1.0 = fully opaque)
            
        Returns:
            True if opacity was set, False if not applicable
            
        Note:
            Line opacity requires the line to have a solid fill. We manipulate
            the OOXML <a:ln><a:solidFill><a:srgbClr><a:alpha> structure.
        """
        if opacity >= 1.0:
            return True
        
        if opacity < 0.0:
            opacity = 0.0
        
        try:
            # Access the shape's spPr element
            spPr = shape._sp.spPr
            if spPr is None:
                return False
            
            # Find the line element
            ln = spPr.find(qn('a:ln'))
            if ln is None:
                return False
            
            # Find solidFill within line
            solidFill = ln.find(qn('a:solidFill'))
            if solidFill is None:
                # Line might not have a fill yet - try to find/create one
                return False
            
            # Find color element
            color_elem = solidFill.find(qn('a:srgbClr'))
            if color_elem is None:
                color_elem = solidFill.find(qn('a:schemeClr'))
            if color_elem is None:
                return False
            
            # Calculate and set alpha
            alpha_value = int(opacity * 100000)
            
            existing_alpha = color_elem.find(qn('a:alpha'))
            if existing_alpha is not None:
                color_elem.remove(existing_alpha)
            
            alpha_elem = etree.SubElement(color_elem, qn('a:alpha'))
            alpha_elem.set('val', str(alpha_value))
            
            return True
            
        except Exception as e:
            self._log_warning(f"Could not set line opacity: {e}")
            return False
    
    def _ensure_line_solid_fill(self, shape, color_hex: str) -> bool:
        """
        Ensure the shape's line has a solid fill with the specified color.
        This is necessary before setting line opacity.
        
        Args:
            shape: The shape object
            color_hex: Hex color string for the line
            
        Returns:
            True if successful
        """
        try:
            # Set line color through python-pptx first
            shape.line.color.rgb = ColorHelper.from_hex(color_hex)
            
            # Now ensure the XML structure is correct for opacity
            spPr = shape._sp.spPr
            ln = spPr.find(qn('a:ln'))
            
            if ln is None:
                return False
            
            # Check if solidFill exists
            solidFill = ln.find(qn('a:solidFill'))
            if solidFill is None:
                # Create solidFill structure
                solidFill = etree.SubElement(ln, qn('a:solidFill'))
                color_elem = etree.SubElement(solidFill, qn('a:srgbClr'))
                # Remove # from hex color
                color_val = color_hex.lstrip('#').upper()
                color_elem.set('val', color_val)
            
            return True
            
        except Exception as e:
            self._log_warning(f"Could not ensure line solid fill: {e}")
            return False
    
    def add_shape(
        self,
        slide_index: int,
        shape_type: str,
        position: Dict[str, Any],
        size: Dict[str, Any],
        fill_color: Optional[str] = None,
        fill_opacity: float = 1.0,
        line_color: Optional[str] = None,
        line_opacity: float = 1.0,
        line_width: float = 1.0,
        text: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Add shape to slide with optional transparency/opacity support.
        
        Args:
            slide_index: Target slide index
            shape_type: Shape type name (rectangle, ellipse, arrow_right, etc.)
            position: Position dict (percentage, inches, anchor, or grid)
            size: Size dict (percentage or inches)
            fill_color: Fill color hex (e.g., "#0070C0") or None for no fill
            fill_opacity: Fill opacity from 0.0 (transparent) to 1.0 (opaque).
                         Default is 1.0 (fully opaque). Use 0.15 for subtle overlays.
            line_color: Line/border color hex or None for no line
            line_opacity: Line opacity from 0.0 (transparent) to 1.0 (opaque).
                         Default is 1.0 (fully opaque).
            line_width: Line width in points (default: 1.0)
            text: Optional text to add inside shape
            
        Returns:
            Dict with shape_index, position, size, and applied styling details
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ValueError: If size is not specified or opacity is out of range
            
        Example:
            # Subtle white overlay for improved text readability
            agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "0%", "top": "0%"},
                size={"width": "100%", "height": "100%"},
                fill_color="#FFFFFF",
                fill_opacity=0.15  # 15% opaque = 85% transparent
            )
        """
        # Validate opacity ranges
        if not 0.0 <= fill_opacity <= 1.0:
            raise ValueError(
                f"fill_opacity must be between 0.0 and 1.0, got {fill_opacity}"
            )
        if not 0.0 <= line_opacity <= 1.0:
            raise ValueError(
                f"line_opacity must be between 0.0 and 1.0, got {line_opacity}"
            )
        
        slide = self._get_slide(slide_index)
        version_before = self._capture_version()
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Shape must have explicit width and height")
        
        # Map shape type string to MSO constant
        shape_type_map = {
            "rectangle": MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            "rounded_rectangle": MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            "ellipse": MSO_AUTO_SHAPE_TYPE.OVAL,
            "oval": MSO_AUTO_SHAPE_TYPE.OVAL,
            "triangle": MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE,
            "arrow_right": MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW,
            "arrow_left": MSO_AUTO_SHAPE_TYPE.LEFT_ARROW,
            "arrow_up": MSO_AUTO_SHAPE_TYPE.UP_ARROW,
            "arrow_down": MSO_AUTO_SHAPE_TYPE.DOWN_ARROW,
            "diamond": MSO_AUTO_SHAPE_TYPE.DIAMOND,
            "pentagon": MSO_AUTO_SHAPE_TYPE.PENTAGON,
            "hexagon": MSO_AUTO_SHAPE_TYPE.HEXAGON,
            "star": MSO_AUTO_SHAPE_TYPE.STAR_5_POINT,
            "heart": MSO_AUTO_SHAPE_TYPE.HEART,
            "lightning": MSO_AUTO_SHAPE_TYPE.LIGHTNING_BOLT,
            "sun": MSO_AUTO_SHAPE_TYPE.SUN,
            "moon": MSO_AUTO_SHAPE_TYPE.MOON,
            "cloud": MSO_AUTO_SHAPE_TYPE.CLOUD,
        }
        
        mso_shape = shape_type_map.get(shape_type.lower())
        if mso_shape is None:
             raise ValueError(
                f"Unknown shape type: {shape_type}",
                details={"valid_types": list(shape_type_map.keys())}
            )
        
        # Add shape
        shape = slide.shapes.add_shape(
            mso_shape,
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        # Track what was actually applied
        styling_applied = {
            "fill_color": None,
            "fill_opacity": 1.0,
            "fill_opacity_applied": False,
            "line_color": None,
            "line_opacity": 1.0,
            "line_opacity_applied": False,
            "line_width": line_width
        }
        
        # Apply fill color and opacity
        if fill_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = ColorHelper.from_hex(fill_color)
            styling_applied["fill_color"] = fill_color
            styling_applied["fill_opacity"] = fill_opacity
            
            # Apply fill opacity if not fully opaque
            if fill_opacity < 1.0:
                opacity_set = self._set_fill_opacity(shape, fill_opacity)
                styling_applied["fill_opacity_applied"] = opacity_set
        else:
            # No fill - make background transparent
            shape.fill.background()
        
        # Apply line color and opacity
        if line_color:
            # Ensure line has solid fill for opacity support
            self._ensure_line_solid_fill(shape, line_color)
            shape.line.width = Pt(line_width)
            styling_applied["line_color"] = line_color
            styling_applied["line_opacity"] = line_opacity
            
            # Apply line opacity if not fully opaque
            if line_opacity < 1.0:
                opacity_set = self._set_line_opacity(shape, line_opacity)
                styling_applied["line_opacity_applied"] = opacity_set
        else:
            # No line
            shape.line.fill.background()
        
        # Add text if provided
        if text and shape.has_text_frame:
            shape.text_frame.text = text
        
        shape_index = len(slide.shapes) - 1
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "shape_type": shape_type,
            "position": {"left": left, "top": top},
            "size": {"width": width, "height": height},
            "styling": styling_applied,
            "has_text": text is not None,
            "text_preview": text[:50] + "..." if text and len(text) > 50 else text,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def format_shape(
        self,
        slide_index: int,
        shape_index: int,
        fill_color: Optional[str] = None,
        fill_opacity: Optional[float] = None,
        line_color: Optional[str] = None,
        line_opacity: Optional[float] = None,
        line_width: Optional[float] = None,
        transparency: Optional[float] = None
    ) -> Dict[str, Any]:
        """
        Format existing shape with optional transparency/opacity support.
        
        Args:
            slide_index: Target slide index
            shape_index: Shape index on slide
            fill_color: Fill color hex (e.g., "#0070C0")
            fill_opacity: Fill opacity from 0.0 (transparent) to 1.0 (opaque)
            line_color: Line/border color hex
            line_opacity: Line opacity from 0.0 (transparent) to 1.0 (opaque)
            line_width: Line width in points
            transparency: DEPRECATED - Use fill_opacity instead.
                         If provided, converted to fill_opacity (transparency = 1 - opacity).
                         Will be removed in v4.0.
            
        Returns:
            Dict with formatting changes applied and their status
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ShapeNotFoundError: If shape index is invalid
            ValueError: If opacity values are out of range
            
        Example:
            # Make an existing shape semi-transparent
            agent.format_shape(
                slide_index=0,
                shape_index=3,
                fill_opacity=0.5  # 50% opaque
            )
        """
        shape = self._get_shape(slide_index, shape_index)
        version_before = self._capture_version()
        
        changes: List[str] = []
        changes_detail: Dict[str, Any] = {}
        
        # Handle deprecated transparency parameter
        if transparency is not None:
            if fill_opacity is None:
                # Convert transparency to opacity (they're inverses)
                # transparency: 0.0 = opaque, 1.0 = invisible
                # opacity: 1.0 = opaque, 0.0 = invisible
                fill_opacity = 1.0 - transparency
                changes.append("transparency_converted_to_opacity")
                changes_detail["transparency_deprecated"] = True
                changes_detail["transparency_value"] = transparency
                changes_detail["converted_opacity"] = fill_opacity
                self._log_warning(
                    "The 'transparency' parameter is deprecated. "
                    "Use 'fill_opacity' instead (opacity = 1 - transparency)."
                )
            else:
                # Both provided - fill_opacity takes precedence
                changes.append("transparency_ignored")
                changes_detail["transparency_ignored"] = True
                self._log_warning(
                    "Both 'transparency' and 'fill_opacity' provided. "
                    "Using 'fill_opacity', ignoring 'transparency'."
                )
        
        # Validate opacity ranges
        if fill_opacity is not None and not 0.0 <= fill_opacity <= 1.0:
            raise ValueError(
                f"fill_opacity must be between 0.0 and 1.0, got {fill_opacity}"
            )
        if line_opacity is not None and not 0.0 <= line_opacity <= 1.0:
            raise ValueError(
                f"line_opacity must be between 0.0 and 1.0, got {line_opacity}"
            )
        
        # Apply fill color
        if fill_color is not None:
            shape.fill.solid()
            shape.fill.fore_color.rgb = ColorHelper.from_hex(fill_color)
            changes.append("fill_color")
            changes_detail["fill_color"] = fill_color
        
        # Apply fill opacity
        if fill_opacity is not None:
            # Ensure shape has solid fill before applying opacity
            if fill_color is None:
                try:
                    shape.fill.solid()
                except Exception:
                    pass
            
            if fill_opacity < 1.0:
                success = self._set_fill_opacity(shape, fill_opacity)
                if success:
                    changes.append("fill_opacity")
                    changes_detail["fill_opacity"] = fill_opacity
                    changes_detail["fill_opacity_applied"] = True
                else:
                    changes.append("fill_opacity_failed")
                    changes_detail["fill_opacity"] = fill_opacity
                    changes_detail["fill_opacity_applied"] = False
            else:
                # Opacity 1.0 = fully opaque (default, no XML change needed)
                changes.append("fill_opacity_reset")
                changes_detail["fill_opacity"] = 1.0
        
        # Apply line color
        if line_color is not None:
            self._ensure_line_solid_fill(shape, line_color)
            changes.append("line_color")
            changes_detail["line_color"] = line_color
        
        # Apply line opacity
        if line_opacity is not None:
            if line_opacity < 1.0:
                success = self._set_line_opacity(shape, line_opacity)
                if success:
                    changes.append("line_opacity")
                    changes_detail["line_opacity"] = line_opacity
                    changes_detail["line_opacity_applied"] = True
                else:
                    changes.append("line_opacity_failed")
                    changes_detail["line_opacity"] = line_opacity
                    changes_detail["line_opacity_applied"] = False
            else:
                changes.append("line_opacity_reset")
                changes_detail["line_opacity"] = 1.0
        
        # Apply line width
        if line_width is not None:
            shape.line.width = Pt(line_width)
            changes.append("line_width")
            changes_detail["line_width"] = line_width
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "changes_applied": changes,
            "changes_detail": changes_detail,
            "success": "failed" not in " ".join(changes),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def remove_shape(
        self,
        slide_index: int,
        shape_index: int,
        approval_token: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Remove shape from slide.
        
        ⚠️ DESTRUCTIVE OPERATION - Requires approval token.
        
        Args:
            slide_index: Target slide index
            shape_index: Shape index to remove
            
        Returns:
            Dict with removal details
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ShapeNotFoundError: If shape index is invalid
        """
        self._validate_token(approval_token, APPROVAL_SCOPE_REMOVE_SHAPE)
        
        slide = self._get_slide(slide_index)
        shape = self._get_shape(slide_index, shape_index)
        version_before = self._capture_version()
        
        # Get shape info before removal
        shape_name = shape.name
        shape_type = str(shape.shape_type)
        
        # Remove shape from slide
        sp = shape.element
        sp.getparent().remove(sp)
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "removed_shape_index": shape_index,
            "removed_shape_name": shape_name,
            "removed_shape_type": shape_type,
            "new_shape_count": len(slide.shapes),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def set_z_order(
        self,
        slide_index: int,
        shape_index: int,
        action: str
    ) -> Dict[str, Any]:
        """
        Change the z-order (stacking order) of a shape.
        
        Args:
            slide_index: Target slide index
            shape_index: Shape index to modify
            action: One of "bring_to_front", "send_to_back", 
                   "bring_forward", "send_backward"
            
        Returns:
            Dict with z-order change details including old and new positions
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ShapeNotFoundError: If shape index is invalid
            ValueError: If action is invalid
        """
        valid_actions = {"bring_to_front", "send_to_back", "bring_forward", "send_backward"}
        if action not in valid_actions:
            raise ValueError(f"Invalid action: {action}. Must be one of {valid_actions}")
        
        slide = self._get_slide(slide_index)
        shape = self._get_shape(slide_index, shape_index)
        version_before = self._capture_version()
        
        # Access the shape tree XML element
        sp_tree = slide.shapes._spTree
        element = shape.element
        
        # Find current position in XML tree
        current_index = -1
        shape_elements = [child for child in sp_tree if child.tag.endswith('}sp') or 
                         child.tag.endswith('}pic') or child.tag.endswith('}graphicFrame')]
        
        for i, child in enumerate(sp_tree):
            if child == element:
                current_index = i
                break
        
        if current_index == -1:
            raise PowerPointAgentError(
                "Could not locate shape in XML tree",
                details={"slide_index": slide_index, "shape_index": shape_index}
            )
        
        new_index = current_index
        max_index = len(sp_tree) - 1
        
        # Execute the z-order action
        if action == "bring_to_front":
            sp_tree.remove(element)
            sp_tree.append(element)
            new_index = len(sp_tree) - 1
            
        elif action == "send_to_back":
            sp_tree.remove(element)
            # Insert after nvGrpSpPr and grpSpPr (indices 0 and 1 typically)
            sp_tree.insert(2, element)
            new_index = 2
            
        elif action == "bring_forward":
            if current_index < max_index:
                sp_tree.remove(element)
                sp_tree.insert(current_index + 1, element)
                new_index = current_index + 1
                
        elif action == "send_backward":
            if current_index > 2:  # Don't go before required elements
                sp_tree.remove(element)
                sp_tree.insert(current_index - 1, element)
                new_index = current_index - 1
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "action": action,
            "z_order_change": {
                "from": current_index,
                "to": new_index
            },
            "warning": "Shape indices may have changed after z-order operation. Re-query slide info.",
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def add_table(
        self,
        slide_index: int,
        rows: int,
        cols: int,
        position: Dict[str, Any],
        size: Dict[str, Any],
        data: Optional[List[List[Any]]] = None,
        header_row: bool = True
    ) -> Dict[str, Any]:
        """
        Add table to slide.
        
        Args:
            slide_index: Target slide index
            rows: Number of rows
            cols: Number of columns
            position: Position dict
            size: Size dict
            data: Optional 2D list of cell values
            header_row: Whether first row is header (styling hint)
            
        Returns:
            Dict with shape_index and table details
        """
        slide = self._get_slide(slide_index)
        version_before = self._capture_version()
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Table must have explicit width and height")
        
        # Create table
        table_shape = slide.shapes.add_table(
            rows, cols,
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        table = table_shape.table
        
        # Populate with data if provided
        cells_filled = 0
        if data:
            for row_idx, row_data in enumerate(data):
                if row_idx >= rows:
                    break
                for col_idx, cell_value in enumerate(row_data):
                    if col_idx >= cols:
                        break
                    table.cell(row_idx, col_idx).text = str(cell_value)
                    cells_filled += 1
        
        shape_index = len(slide.shapes) - 1
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "rows": rows,
            "cols": cols,
            "cells_filled": cells_filled,
            "position": {"left": left, "top": top},
            "size": {"width": width, "height": height},
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def add_connector(
        self,
        slide_index: int,
        from_shape_index: int,
        to_shape_index: int,
        connector_type: str = "straight"
    ) -> Dict[str, Any]:
        """
        Add connector line between two shapes.
        
        Args:
            slide_index: Target slide index
            from_shape_index: Starting shape index
            to_shape_index: Ending shape index
            connector_type: "straight", "elbow", or "curved"
            
        Returns:
            Dict with connector details
        """
        slide = self._get_slide(slide_index)
        version_before = self._capture_version()
        
        shape1 = self._get_shape(slide_index, from_shape_index)
        shape2 = self._get_shape(slide_index, to_shape_index)
        
        # Calculate center points
        x1 = shape1.left + shape1.width // 2
        y1 = shape1.top + shape1.height // 2
        x2 = shape2.left + shape2.width // 2
        y2 = shape2.top + shape2.height // 2
        
        # Map connector type
        connector_map = {
            "straight": MSO_CONNECTOR.STRAIGHT,
            "elbow": MSO_CONNECTOR.ELBOW,
            "curved": MSO_CONNECTOR.CURVE
        }
        mso_connector = connector_map.get(connector_type.lower(), MSO_CONNECTOR.STRAIGHT)
        
        # Add connector
        connector = slide.shapes.add_connector(
            mso_connector,
            x1, y1, x2, y2
        )
        
        shape_index = len(slide.shapes) - 1
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "from_shape": from_shape_index,
            "to_shape": to_shape_index,
            "connector_type": connector_type,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    # ========================================================================
    # IMAGE OPERATIONS
    # ========================================================================
    
    def insert_image(
        self,
        slide_index: int,
        image_path: Union[str, Path],
        position: Dict[str, Any],
        size: Optional[Dict[str, Any]] = None,
        alt_text: Optional[str] = None,
        compress: bool = False
    ) -> Dict[str, Any]:
        """
        Insert image on slide.
        
        Args:
            slide_index: Target slide index
            image_path: Path to image file
            position: Position dict
            size: Optional size dict (can use "auto" for aspect ratio)
            alt_text: Alternative text for accessibility
            compress: Compress image before inserting
            
        Returns:
            Dict with shape_index and image details
        """
        slide = self._get_slide(slide_index)
        image_path = PathValidator.validate_image_path(image_path)
        version_before = self._capture_version()
        
        left, top = Position.from_dict(position)
        
        # Get aspect ratio if Pillow available
        aspect_ratio = None
        if HAS_PILLOW:
            try:
                with PILImage.open(image_path) as img:
                    aspect_ratio = img.width / img.height
            except Exception:
                pass
        
        # Parse size
        if size:
            width, height = Size.from_dict(size, aspect_ratio=aspect_ratio)
        else:
            # Default to half slide width, maintain aspect ratio
            width = SLIDE_WIDTH_INCHES * 0.5
            if aspect_ratio:
                height = width / aspect_ratio
            else:
                height = SLIDE_HEIGHT_INCHES * 0.3
        
        # Compress if requested
        if compress and HAS_PILLOW:
            image_stream = AssetValidator.compress_image(image_path)
            picture = slide.shapes.add_picture(
                image_stream,
                Inches(left), Inches(top),
                width=Inches(width) if width else None,
                height=Inches(height) if height else None
            )
        else:
            picture = slide.shapes.add_picture(
                str(image_path),
                Inches(left), Inches(top),
                width=Inches(width) if width else None,
                height=Inches(height) if height else None
            )
        
        # Set alt text
        if alt_text:
            picture.name = alt_text
            try:
                # Set description attribute for proper alt text
                picture._element.set('descr', alt_text)
            except Exception:
                pass
        
        shape_index = len(slide.shapes) - 1
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "image_path": str(image_path),
            "position": {"left": left, "top": top},
            "size": {"width": width, "height": height},
            "alt_text_set": alt_text is not None,
            "compressed": compress,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def replace_image(
        self,
        slide_index: int,
        old_image_name: str,
        new_image_path: Union[str, Path],
        compress: bool = False
    ) -> Dict[str, Any]:
        """
        Replace existing image by name.
        
        Args:
            slide_index: Target slide index
            old_image_name: Name or partial name of image to replace
            new_image_path: Path to new image file
            compress: Compress new image
            
        Returns:
            Dict with replacement details
        """
        slide = self._get_slide(slide_index)
        new_image_path = PathValidator.validate_image_path(new_image_path)
        version_before = self._capture_version()
        
        replaced = False
        old_shape_index = None
        new_shape_index = None
        
        for idx, shape in enumerate(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if shape.name == old_image_name or old_image_name in (shape.name or ""):
                    # Store position and size
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    old_shape_index = idx
                    
                    # Remove old image
                    sp = shape.element
                    sp.getparent().remove(sp)
                    
                    # Add new image
                    if compress and HAS_PILLOW:
                        image_stream = AssetValidator.compress_image(new_image_path)
                        new_picture = slide.shapes.add_picture(
                            image_stream, left, top,
                            width=width, height=height
                        )
                    else:
                        new_picture = slide.shapes.add_picture(
                            str(new_image_path), left, top,
                            width=width, height=height
                        )
                    
                    new_shape_index = len(slide.shapes) - 1
                    replaced = True
                    break
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "replaced": replaced,
            "old_image_name": old_image_name,
            "old_shape_index": old_shape_index,
            "new_image_path": str(new_image_path),
            "new_shape_index": new_shape_index,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def set_image_properties(
        self,
        slide_index: int,
        shape_index: int,
        alt_text: Optional[str] = None,
        name: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Set image properties.
        
        Args:
            slide_index: Target slide index
            shape_index: Image shape index
            alt_text: Alternative text for accessibility
            name: Shape name
            
        Returns:
            Dict with properties set
        """
        shape = self._get_shape(slide_index, shape_index)
        version_before = self._capture_version()
        
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            raise ValueError(f"Shape at index {shape_index} is not an image")
        
        changes = []
        
        if alt_text is not None:
            try:
                shape._element.set('descr', alt_text)
                changes.append("alt_text")
            except Exception:
                # Fallback to name
                shape.name = alt_text
                changes.append("alt_text_via_name")
        
        if name is not None:
            shape.name = name
            changes.append("name")
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "changes_applied": changes,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def crop_image(
        self,
        slide_index: int,
        shape_index: int,
        left: float = 0.0,
        top: float = 0.0,
        right: float = 0.0,
        bottom: float = 0.0
    ) -> Dict[str, Any]:
        """
        Crop image by specifying crop amounts from each edge.
        
        Args:
            slide_index: Target slide index
            shape_index: Image shape index
            left: Crop from left (0.0 to 1.0, proportion of width)
            top: Crop from top (0.0 to 1.0, proportion of height)
            right: Crop from right (0.0 to 1.0, proportion of width)
            bottom: Crop from bottom (0.0 to 1.0, proportion of height)
            
        Returns:
            Dict with crop details
        """
        shape = self._get_shape(slide_index, shape_index)
        version_before = self._capture_version()
        
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            raise ValueError(f"Shape at index {shape_index} is not an image")
        
        # Validate crop values
        for name, value in [("left", left), ("top", top), ("right", right), ("bottom", bottom)]:
            if not 0.0 <= value < 1.0:
                raise ValueError(f"Crop {name} must be between 0.0 and 1.0, got {value}")
        
        if left + right >= 1.0:
            raise ValueError("Left + right crop cannot equal or exceed 1.0")
        if top + bottom >= 1.0:
            raise ValueError("Top + bottom crop cannot equal or exceed 1.0")
        
        # Apply crop using picture's crop properties
        try:
            # Access the picture element
            pic = shape._element
            
            # Find or create blipFill element
            blip_fill = pic.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}blipFill')
            if blip_fill is None:
                blip_fill = pic.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blipFill')
            
            if blip_fill is not None:
                # Find or create srcRect element
                ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
                src_rect = blip_fill.find(f'{ns}srcRect')
                
                if src_rect is None:
                    src_rect = etree.SubElement(blip_fill, f'{ns}srcRect')
                
                # Set crop values (in percentage * 1000)
                src_rect.set('l', str(int(left * 100000)))
                src_rect.set('t', str(int(top * 100000)))
                src_rect.set('r', str(int(right * 100000)))
                src_rect.set('b', str(int(bottom * 100000)))
                
                version_after = self._capture_version()
                
                return {
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "crop_applied": True,
                    "crop_values": {
                        "left": left,
                        "top": top,
                        "right": right,
                        "bottom": bottom
                    },
                    "presentation_version_before": version_before,
                    "presentation_version_after": version_after
                }
        except Exception as e:
            logger.warning(f"Could not apply crop via XML: {e}")
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "crop_applied": False,
            "error": "Crop not supported for this image type",
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def resize_image(
        self,
        slide_index: int,
        shape_index: int,
        width: Optional[float] = None,
        height: Optional[float] = None,
        maintain_aspect: bool = True
    ) -> Dict[str, Any]:
        """
        Resize image shape.
        
        Args:
            slide_index: Target slide index
            shape_index: Image shape index
            width: New width in inches (None = keep current)
            height: New height in inches (None = keep current)
            maintain_aspect: Maintain aspect ratio
            
        Returns:
            Dict with new dimensions
        """
        shape = self._get_shape(slide_index, shape_index)
        version_before = self._capture_version()
        
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            raise ValueError(f"Shape at index {shape_index} is not an image")
        
        original_width = shape.width / EMU_PER_INCH
        original_height = shape.height / EMU_PER_INCH
        aspect = original_width / original_height if original_height > 0 else 1.0
        
        new_width = width
        new_height = height
        
        if maintain_aspect:
            if width is not None and height is None:
                new_height = width / aspect
            elif height is not None and width is None:
                new_width = height * aspect
        
        if new_width is not None:
            shape.width = Inches(new_width)
        if new_height is not None:
            shape.height = Inches(new_height)
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "original_size": {"width": original_width, "height": original_height},
            "new_size": {
                "width": new_width or original_width,
                "height": new_height or original_height
            },
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    # ========================================================================
    # CHART OPERATIONS
    # ========================================================================
    
    def add_chart(
        self,
        slide_index: int,
        chart_type: str,
        data: Dict[str, Any],
        position: Dict[str, Any],
        size: Dict[str, Any],
        title: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Add chart to slide.
        
        Args:
            slide_index: Target slide index
            chart_type: Chart type (column, bar, line, pie, etc.)
            data: Chart data dict with "categories" and "series"
            position: Position dict
            size: Size dict
            title: Optional chart title
            
        Returns:
            Dict with shape_index and chart details
            
        Example data:
            {
                "categories": ["Q1", "Q2", "Q3", "Q4"],
                "series": [
                    {"name": "Revenue", "values": [100, 120, 140, 160]},
                    {"name": "Costs", "values": [80, 90, 100, 110]}
                ]
            }
        """
        slide = self._get_slide(slide_index)
        version_before = self._capture_version()
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Chart must have explicit width and height")
        
        # Map chart type string to XL constant
        chart_type_map = {
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "column_clustered": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "column_stacked": XL_CHART_TYPE.COLUMN_STACKED,
            "bar": XL_CHART_TYPE.BAR_CLUSTERED,
            "bar_clustered": XL_CHART_TYPE.BAR_CLUSTERED,
            "bar_stacked": XL_CHART_TYPE.BAR_STACKED,
            "line": XL_CHART_TYPE.LINE,
            "line_markers": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE,
            "pie_exploded": XL_CHART_TYPE.PIE_EXPLODED,
            "area": XL_CHART_TYPE.AREA,
            "scatter": XL_CHART_TYPE.XY_SCATTER,
            "doughnut": XL_CHART_TYPE.DOUGHNUT,
        }
        
        xl_chart_type = chart_type_map.get(
            chart_type.lower(),
            XL_CHART_TYPE.COLUMN_CLUSTERED
        )
        
        # Build chart data
        chart_data = CategoryChartData()
        chart_data.categories = data.get("categories", [])
        
        for series in data.get("series", []):
            chart_data.add_series(series["name"], series["values"])
        
        # Add chart
        chart_shape = slide.shapes.add_chart(
            xl_chart_type,
            Inches(left), Inches(top),
            Inches(width), Inches(height),
            chart_data
        )
        
        # Set title if provided
        if title:
            chart_shape.chart.has_title = True
            chart_shape.chart.chart_title.text_frame.text = title
        
        shape_index = len(slide.shapes) - 1
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "chart_type": chart_type,
            "categories_count": len(data.get("categories", [])),
            "series_count": len(data.get("series", [])),
            "title": title,
            "position": {"left": left, "top": top},
            "size": {"width": width, "height": height},
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def update_chart_data(
        self,
        slide_index: int,
        chart_index: int,
        data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Update existing chart data.
        
        Args:
            slide_index: Target slide index
            chart_index: Chart index on slide (not shape index)
            data: New chart data dict
            
        Returns:
            Dict with update details
        """
        chart_shape = self._get_chart_shape(slide_index, chart_index)
        version_before = self._capture_version()
        
        # Build new chart data
        chart_data = CategoryChartData()
        chart_data.categories = data.get("categories", [])
        
        for series in data.get("series", []):
            chart_data.add_series(series["name"], series["values"])
        
        # Try to replace data (preserves formatting)
        try:
            chart_shape.chart.replace_data(chart_data)
            method = "replace_data"
        except AttributeError:
            # Fallback: recreate chart (loses some formatting)
            logger.warning(
                "chart.replace_data() not available. "
                "Recreating chart (some formatting may be lost)."
            )
            
            slide = self._get_slide(slide_index)
            
            # Store chart properties
            left = chart_shape.left
            top = chart_shape.top
            width = chart_shape.width
            height = chart_shape.height
            chart_type = chart_shape.chart.chart_type
            has_title = chart_shape.chart.has_title
            title_text = None
            if has_title:
                try:
                    title_text = chart_shape.chart.chart_title.text_frame.text
                except Exception:
                    pass
            
            # Remove old chart
            sp = chart_shape.element
            sp.getparent().remove(sp)
            
            # Create new chart
            new_chart_shape = slide.shapes.add_chart(
                chart_type, left, top, width, height, chart_data
            )
            
            # Restore title
            if title_text:
                new_chart_shape.chart.has_title = True
                new_chart_shape.chart.chart_title.text_frame.text = title_text
            
            method = "recreate"
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "chart_index": chart_index,
            "categories_count": len(data.get("categories", [])),
            "series_count": len(data.get("series", [])),
            "update_method": method,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def format_chart(
        self,
        slide_index: int,
        chart_index: int,
        title: Optional[str] = None,
        legend_position: Optional[str] = None,
        has_legend: Optional[bool] = None
    ) -> Dict[str, Any]:
        """
        Format existing chart.
        
        Args:
            slide_index: Target slide index
            chart_index: Chart index on slide
            title: Chart title
            legend_position: Legend position ("bottom", "left", "right", "top")
            has_legend: Show/hide legend
            
        Returns:
            Dict with formatting changes
        """
        chart_shape = self._get_chart_shape(slide_index, chart_index)
        version_before = self._capture_version()
        
        chart = chart_shape.chart
        
        changes = []
        
        if title is not None:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
            changes.append("title")
        
        if has_legend is not None:
            chart.has_legend = has_legend
            changes.append("has_legend")
        
        if legend_position is not None and chart.has_legend:
            from pptx.enum.chart import XL_LEGEND_POSITION
            position_map = {
                "bottom": XL_LEGEND_POSITION.BOTTOM,
                "left": XL_LEGEND_POSITION.LEFT,
                "right": XL_LEGEND_POSITION.RIGHT,
                "top": XL_LEGEND_POSITION.TOP,
                "corner": XL_LEGEND_POSITION.CORNER,
            }
            if legend_position.lower() in position_map:
                chart.legend.position = position_map[legend_position.lower()]
                changes.append("legend_position")
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "chart_index": chart_index,
            "changes_applied": changes,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    # ========================================================================
    # LAYOUT & THEME OPERATIONS
    # ========================================================================
    
    def set_slide_layout(self, slide_index: int, layout_name: str) -> Dict[str, Any]:
        """
        Change slide layout.
        
        Note: This changes the layout but may not reposition existing content.
        
        Args:
            slide_index: Target slide index
            layout_name: Name of new layout
            
        Returns:
            Dict with layout change details
        """
        slide = self._get_slide(slide_index)
        version_before = self._capture_version()
        
        layout = self._get_layout(layout_name)
        
        old_layout = slide.slide_layout.name
        slide.slide_layout = layout
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "old_layout": old_layout,
            "new_layout": layout_name,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def set_background(
        self,
        slide_index: Optional[int] = None,
        color: Optional[str] = None,
        image_path: Optional[Union[str, Path]] = None
    ) -> Dict[str, Any]:
        """
        Set slide background color or image.
        
        Args:
            slide_index: Target slide (None = all slides)
            color: Background color hex
            image_path: Background image path
            
        Returns:
            Dict with background change details
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        if color is None and image_path is None:
            raise ValueError("Must specify either color or image_path")
        
        version_before = self._capture_version()
        results = []
        
        # Determine slides to process
        if slide_index is not None:
            slides = [(slide_index, self._get_slide(slide_index))]
        else:
            slides = list(enumerate(self.prs.slides))
        
        for s_idx, slide in slides:
            result = {"slide_index": s_idx, "success": False}
            
            try:
                background = slide.background
                fill = background.fill
                
                if color:
                    fill.solid()
                    fill.fore_color.rgb = ColorHelper.from_hex(color)
                    result["success"] = True
                    result["type"] = "color"
                    result["color"] = color
                
                elif image_path:
                    # Note: python-pptx has limited background image support
                    # This is a best-effort implementation
                    image_path = PathValidator.validate_image_path(image_path)
                    result["type"] = "image"
                    result["image_path"] = str(image_path)
                    result["note"] = "Background image support is limited in python-pptx"
                    
            except Exception as e:
                result["error"] = str(e)
            
            results.append(result)
        
        version_after = self._capture_version()
        
        return {
            "slides_processed": len(results),
            "results": results,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def get_available_layouts(self) -> List[str]:
        """
        Get list of available layout names.
        
        Returns:
            List of layout name strings
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        self._ensure_layout_cache()
        return list(self._layout_cache.keys())
    
    # ========================================================================
    # VALIDATION OPERATIONS
    # ========================================================================
    
    def validate_presentation(self) -> Dict[str, Any]:
        """
        Comprehensive presentation validation.
        
        Returns:
            Validation report dict
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        issues = {
            "empty_slides": [],
            "slides_without_titles": [],
            "fonts_used": set(),
            "large_shapes": []
        }
        
        for idx, slide in enumerate(self.prs.slides):
            # Check for empty slides
            if len(slide.shapes) == 0:
                issues["empty_slides"].append(idx)
            
            # Check for title
            has_title = False
            for shape in slide.shapes:
                if shape.is_placeholder:
                    ph_type = _get_placeholder_type_int_helper(shape.placeholder_format.type)
                    if ph_type in TITLE_PLACEHOLDER_TYPES:
                        if shape.has_text_frame and shape.text_frame.text.strip():
                            has_title = True
                            break
                
                # Collect fonts
                if hasattr(shape, 'text_frame') and shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        if para.font.name:
                            issues["fonts_used"].add(para.font.name)
            
            if not has_title:
                issues["slides_without_titles"].append(idx)
        
        issues["fonts_used"] = list(issues["fonts_used"])
        
        total_issues = (
            len(issues["empty_slides"]) +
            len(issues["slides_without_titles"])
        )
        
        return {
            "status": "issues_found" if total_issues > 0 else "valid",
            "total_issues": total_issues,
            "slide_count": len(self.prs.slides),
            "issues": issues
        }
    
    def check_accessibility(self) -> Dict[str, Any]:
        """
        Run accessibility checker.
        
        Returns:
            Accessibility report dict
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        return AccessibilityChecker.check_presentation(self.prs)
    
    def validate_assets(self) -> Dict[str, Any]:
        """
        Run asset validator.
        
        Returns:
            Asset validation report dict
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        return AssetValidator.validate_presentation_assets(self.prs, self.filepath)
    
    # ========================================================================
    # EXPORT OPERATIONS
    # ========================================================================
    
    def export_to_pdf(self, output_path: Union[str, Path]) -> Dict[str, Any]:
        """
        Export presentation to PDF.
        
        Requires LibreOffice or Microsoft Office installed.
        
        Args:
            output_path: Output PDF path
            
        Returns:
            Dict with export details
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        output_path = Path(output_path)
        if output_path.suffix.lower() != '.pdf':
            output_path = output_path.with_suffix('.pdf')
        
        # Ensure parent directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Save to temp file first
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
            temp_pptx = Path(tmp.name)
        
        try:
            self.prs.save(str(temp_pptx))
            
            # Try LibreOffice conversion
            result = subprocess.run(
                [
                    'soffice', '--headless', '--convert-to', 'pdf',
                    '--outdir', str(output_path.parent), str(temp_pptx)
                ],
                capture_output=True,
                timeout=120
            )
            
            if result.returncode != 0:
                raise PowerPointAgentError(
                    "PDF export failed. LibreOffice is required for PDF export.",
                    details={
                        "stderr": result.stderr.decode() if result.stderr else None,
                        "install_instructions": {
                            "linux": "sudo apt install libreoffice-impress",
                            "macos": "brew install --cask libreoffice",
                            "windows": "Download from libreoffice.org"
                        }
                    }
                )
            
            # Rename output file to desired name
            generated_pdf = output_path.parent / f"{temp_pptx.stem}.pdf"
            if generated_pdf.exists() and generated_pdf != output_path:
                shutil.move(str(generated_pdf), str(output_path))
            
            return {
                "success": True,
                "output_path": str(output_path),
                "file_size_bytes": output_path.stat().st_size if output_path.exists() else 0
            }
            
        finally:
            temp_pptx.unlink(missing_ok=True)
    
    def extract_notes(self) -> Dict[int, str]:
        """
        Extract speaker notes from all slides.
        
        Returns:
            Dict mapping slide index to notes text
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        notes = {}
        
        for idx, slide in enumerate(self.prs.slides):
            if slide.has_notes_slide:
                try:
                    notes_slide = slide.notes_slide
                    text_frame = notes_slide.notes_text_frame
                    if text_frame.text and text_frame.text.strip():
                        notes[idx] = text_frame.text
                except Exception:
                    pass
        
        return notes
    
    # ========================================================================
    # INFORMATION & VERSIONING
    # ========================================================================
    
    def get_presentation_info(self) -> Dict[str, Any]:
        """
        Get presentation metadata and information.
        
        Returns:
            Dict with presentation information
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        info = {
            "slide_count": len(self.prs.slides),
            "layouts": self.get_available_layouts(),
            "slide_width_inches": self.prs.slide_width / EMU_PER_INCH,
            "slide_height_inches": self.prs.slide_height / EMU_PER_INCH,
            "presentation_version": self.get_presentation_version()
        }
        
        # Calculate aspect ratio
        width = info["slide_width_inches"]
        height = info["slide_height_inches"]
        if height > 0:
            ratio = width / height
            if abs(ratio - 16/9) < 0.1:
                info["aspect_ratio"] = "16:9"
            elif abs(ratio - 4/3) < 0.1:
                info["aspect_ratio"] = "4:3"
            else:
                info["aspect_ratio"] = f"{width:.2f}:{height:.2f}"
        
        # File info
        if self.filepath and self.filepath.exists():
            stat = self.filepath.stat()
            info["file"] = str(self.filepath)
            info["file_size_bytes"] = stat.st_size
            info["file_size_mb"] = round(stat.st_size / (1024 * 1024), 2)
            info["modified"] = datetime.fromtimestamp(stat.st_mtime).isoformat()
        
        return info
    
    def get_slide_info(self, slide_index: int) -> Dict[str, Any]:
        """
        Get detailed information about a specific slide.
        
        Args:
            slide_index: Slide index to inspect
            
        Returns:
            Dict with comprehensive slide information
        """
        slide = self._get_slide(slide_index)
        
        shapes_info = []
        for idx, shape in enumerate(slide.shapes):
            # Determine shape type string
            shape_type_str = str(shape.shape_type).replace("MSO_SHAPE_TYPE.", "")
            
            if shape.is_placeholder:
                ph_type = _get_placeholder_type_int_helper(shape.placeholder_format.type)
                ph_name = get_placeholder_type_name(ph_type)
                shape_type_str = f"PLACEHOLDER ({ph_name})"
            
            shape_info = {
                "index": idx,
                "type": shape_type_str,
                "name": shape.name,
                "has_text": hasattr(shape, 'text_frame') and shape.has_text_frame,
                "position": {
                    "left_inches": round(shape.left / EMU_PER_INCH, 3),
                    "top_inches": round(shape.top / EMU_PER_INCH, 3),
                    "left_percent": f"{(shape.left / self.prs.slide_width * 100):.1f}%",
                    "top_percent": f"{(shape.top / self.prs.slide_height * 100):.1f}%"
                },
                "size": {
                    "width_inches": round(shape.width / EMU_PER_INCH, 3),
                    "height_inches": round(shape.height / EMU_PER_INCH, 3),
                    "width_percent": f"{(shape.width / self.prs.slide_width * 100):.1f}%",
                    "height_percent": f"{(shape.height / self.prs.slide_height * 100):.1f}%"
                }
            }
            
            # Add text content if present
            if shape.has_text_frame:
                try:
                    full_text = shape.text_frame.text
                    shape_info["text"] = full_text
                    shape_info["text_length"] = len(full_text)
                except Exception:
                    pass
            
            # Add image info if picture
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    shape_info["image_size_bytes"] = len(shape.image.blob)
                    shape_info["image_content_type"] = shape.image.content_type
                except Exception:
                    pass
            
            # Add chart info if chart
            if hasattr(shape, 'has_chart') and shape.has_chart:
                try:
                    shape_info["chart_type"] = str(shape.chart.chart_type)
                except Exception:
                    pass
            
            shapes_info.append(shape_info)
        
        # Check for notes
        has_notes = False
        notes_preview = None
        if slide.has_notes_slide:
            try:
                notes_text = slide.notes_slide.notes_text_frame.text
                if notes_text and notes_text.strip():
                    has_notes = True
                    notes_preview = notes_text[:100] + "..." if len(notes_text) > 100 else notes_text
            except Exception:
                pass
        
        return {
            "slide_index": slide_index,
            "layout": slide.slide_layout.name,
            "shape_count": len(slide.shapes),
            "shapes": shapes_info,
            "has_notes": has_notes,
            "notes_preview": notes_preview
        }
    
    def get_presentation_version(self) -> str:
        """
        Compute a deterministic version hash for the presentation.
        
        The version is based on:
        - Slide count & Layouts
        - Shape counts per slide
        - Text content (SHA-256)
        - Shape Geometry (Position/Size) to detect layout changes
        
        Returns:
            SHA-256 hash prefix (16 characters)
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        # Build version components
        components = []
        
        # Slide count
        components.append(f"slides:{len(self.prs.slides)}")
        
        # Per-slide information
        for idx, slide in enumerate(self.prs.slides):
            slide_components = [
                f"slide:{idx}",
                f"layout:{slide.slide_layout.name}",
                f"shapes:{len(slide.shapes)}"
            ]
            
            # Add text content hash
            text_content = []
            for shape in slide.shapes:
                # Add Geometry hash to detect moves/resizes
                geo_hash = f"{shape.left}:{shape.top}:{shape.width}:{shape.height}"
                slide_components.append(f"geo:{geo_hash}")
                
                if hasattr(shape, 'text_frame') and shape.has_text_frame:
                    try:
                        text_content.append(shape.text_frame.text)
                    except Exception:
                        pass
            
            if text_content:
                # Use SHA-256 for content
                text_hash = hashlib.sha256("".join(text_content).encode()).hexdigest()[:8]
                slide_components.append(f"text:{text_hash}")
            
            components.extend(slide_components)
        
        # Compute final hash
        version_string = "|".join(components)
        full_hash = hashlib.sha256(version_string.encode()).hexdigest()
        
        return full_hash[:16]
    
    # ========================================================================
    # PRIVATE HELPER METHODS
    # ========================================================================
    
    def _get_slide(self, index: int):
        """
        Get slide by index with validation.
        
        Args:
            index: Slide index (0-based)
            
        Returns:
            Slide object
            
        Raises:
            PowerPointAgentError: If no presentation loaded
            SlideNotFoundError: If index is out of range
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        slide_count = len(self.prs.slides)
        
        if not 0 <= index < slide_count:
            raise SlideNotFoundError(
                f"Slide index {index} out of range (0-{slide_count-1})",
                details={"index": index, "slide_count": slide_count, "valid_range": f"0-{slide_count-1}"}
            )
        
        return self.prs.slides[index]
    
    def _get_shape(self, slide_index: int, shape_index: int):
        """
        Get shape by slide and shape index with validation.
        
        Args:
            slide_index: Slide index
            shape_index: Shape index on slide
            
        Returns:
            Shape object
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ShapeNotFoundError: If shape index is invalid
        """
        slide = self._get_slide(slide_index)
        
        shape_count = len(slide.shapes)
        
        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range on slide {slide_index}",
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "shape_count": shape_count,
                    "valid_range": f"0-{shape_count-1}" if shape_count > 0 else "no shapes"
                }
            )
        
        return slide.shapes[shape_index]
    
    def _get_chart_shape(self, slide_index: int, chart_index: int):
        """
        Get chart shape by slide and chart index.
        
        Args:
            slide_index: Slide index
            chart_index: Chart index on slide (0-based among charts only)
            
        Returns:
            Chart shape object
            
        Raises:
            ChartNotFoundError: If chart not found
        """
        slide = self._get_slide(slide_index)
        
        chart_count = 0
        for shape in slide.shapes:
            if hasattr(shape, 'has_chart') and shape.has_chart:
                if chart_count == chart_index:
                    return shape
                chart_count += 1
        
        raise ChartNotFoundError(
            f"Chart at index {chart_index} not found on slide {slide_index}",
            details={
                "slide_index": slide_index,
                "chart_index": chart_index,
                "charts_found": chart_count
            }
        )
    
    def _get_layout(self, layout_name: str):
        """
        Get layout by name with caching.
        
        Args:
            layout_name: Layout name
            
        Returns:
            Layout object
            
        Raises:
            LayoutNotFoundError: If layout doesn't exist
        """
        self._ensure_layout_cache()
        
        layout = self._layout_cache.get(layout_name)
        
        if layout is None:
            raise LayoutNotFoundError(
                f"Layout '{layout_name}' not found",
                details={"available_layouts": list(self._layout_cache.keys())}
            )
        
        return layout
    
    def _ensure_layout_cache(self) -> None:
        """Build layout cache if not already built."""
        if self._layout_cache is not None:
            return
        
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        self._layout_cache = {
            layout.name: layout
            for layout in self.prs.slide_layouts
        }
    
    def _copy_shape(self, source_shape, target_slide) -> None:
        """
        Copy shape to target slide.
        
        Args:
            source_shape: Shape to copy
            target_slide: Destination slide
        """
        # Handle pictures
        if source_shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                blob = source_shape.image.blob
                target_slide.shapes.add_picture(
                    BytesIO(blob),
                    source_shape.left, source_shape.top,
                    source_shape.width, source_shape.height
                )
            except Exception as e:
                logger.warning(f"Could not copy picture: {e}")
            return
        
        # Handle auto shapes and text boxes
        if source_shape.shape_type in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.TEXT_BOX):
            try:
                # Get auto shape type, default to rectangle
                try:
                    auto_shape_type = source_shape.auto_shape_type
                except Exception:
                    auto_shape_type = MSO_AUTO_SHAPE_TYPE.RECTANGLE
                
                new_shape = target_slide.shapes.add_shape(
                    auto_shape_type,
                    source_shape.left, source_shape.top,
                    source_shape.width, source_shape.height
                )
                
                # Copy text
                if source_shape.has_text_frame:
                    try:
                        new_shape.text_frame.text = source_shape.text_frame.text
                    except Exception:
                        pass
                
                # Copy fill
                try:
                    if source_shape.fill.type == 1:  # Solid fill
                        new_shape.fill.solid()
                        new_shape.fill.fore_color.rgb = source_shape.fill.fore_color.rgb
                except Exception:
                    pass
                
            except Exception as e:
                logger.warning(f"Could not copy shape: {e}")
            return
        
        # Log unsupported shape types
        logger.debug(f"Shape type {source_shape.shape_type} not copied (not supported)")


# ============================================================================
# MODULE EXPORTS
# ============================================================================

__all__ = [
    # Main class
    "PowerPointAgent",
    
    # Exceptions
    "PowerPointAgentError",
    "SlideNotFoundError",
    "ShapeNotFoundError",
    "ChartNotFoundError",
    "LayoutNotFoundError",
    "ImageNotFoundError",
    "InvalidPositionError",
    "TemplateError",
    "ThemeError",
    "AccessibilityError",
    "AssetValidationError",
    "FileLockError",
    "PathValidationError",
    "ApprovalTokenError",
    
    # Utility classes
    "FileLock",
    "PathValidator",
    "Position",
    "Size",
    "ColorHelper",
    "TemplateProfile",
    "AccessibilityChecker",
    "AssetValidator",
    
    # Enums
    "ShapeType",
    "ChartType",
    "TextAlignment",
    "VerticalAlignment",
    "BulletStyle",
    "ImageFormat",
    "ExportFormat",
    "ZOrderAction",
    "NotesMode",
    
    # Constants
    "SLIDE_WIDTH_INCHES",
    "SLIDE_HEIGHT_INCHES",
    "ANCHOR_POINTS",
    "CORPORATE_COLORS",
    "STANDARD_FONTS",
    "APPROVAL_SCOPE_DELETE_SLIDE",
    "APPROVAL_SCOPE_REMOVE_SHAPE",
    "WCAG_CONTRAST_NORMAL",
    "WCAG_CONTRAST_LARGE",
    "EMU_PER_INCH",
    
    # Functions
    "get_placeholder_type_name",
    
    # Module metadata
    "__version__",
    "__author__",
    "__license__",
]

```

# core/strict_validator.py
```py
#!/usr/bin/env python3
"""
Strict JSON Schema Validator
Production-grade JSON Schema validation with rich error reporting and caching.

This module provides comprehensive JSON Schema validation capabilities
for the PowerPoint Agent toolset, supporting manifest validation,
tool output validation, and configuration validation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.0.0

Features:
- Support for JSON Schema Draft-07, Draft-2019-09, and Draft-2020-12
- Schema caching for performance
- Rich error objects with JSON serialization
- ValidationResult objects for programmatic access
- Custom format checkers for presentation-specific formats
- Backward-compatible validate_against_schema() function

Usage:
    from core.strict_validator import (
        validate_against_schema,
        validate_dict,
        validate_json_file,
        ValidationResult,
        ValidationError
    )
    
    # Simple validation (raises on error)
    validate_against_schema(data, "schemas/manifest.schema.json")
    
    # Validation with result object
    result = validate_dict(data, schema)
    if not result.is_valid:
        for error in result.errors:
            print(f"{error.path}: {error.message}")

Changelog v3.0.0:
- NEW: ValidationResult class for structured validation results
- NEW: ValidationError exception with rich details and JSON serialization
- NEW: SchemaCache for performance optimization
- NEW: Support for multiple JSON Schema drafts
- NEW: validate_dict() returning ValidationResult
- NEW: validate_json_file() for file-based validation
- NEW: Custom format checkers (hex-color, percentage, file-path)
- IMPROVED: Error messages with full JSON paths
- IMPROVED: Graceful dependency handling
"""

import json
import re
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, Type
from dataclasses import dataclass, field
from datetime import datetime

# ============================================================================
# DEPENDENCY HANDLING
# ============================================================================

try:
    from jsonschema import (
        Draft7Validator,
        Draft201909Validator,
        Draft202012Validator,
        FormatChecker,
        ValidationError as JsonSchemaValidationError,
        SchemaError as JsonSchemaSchemaError
    )
    from jsonschema.protocols import Validator
    JSONSCHEMA_AVAILABLE = True
except ImportError:
    JSONSCHEMA_AVAILABLE = False
    Draft7Validator = None
    Draft201909Validator = None
    Draft202012Validator = None
    FormatChecker = None
    JsonSchemaValidationError = Exception
    JsonSchemaSchemaError = Exception
    Validator = None


# ============================================================================
# EXCEPTIONS
# ============================================================================

class ValidatorError(Exception):
    """Base exception for validator errors."""
    
    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to JSON-serializable dictionary."""
        return {
            "error": self.__class__.__name__,
            "message": self.message,
            "details": self.details
        }
    
    def to_json(self) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=2)


class ValidationError(ValidatorError):
    """
    Raised when validation fails.
    
    Contains detailed information about all validation errors.
    """
    
    def __init__(
        self,
        message: str,
        errors: Optional[List['ValidationErrorDetail']] = None,
        schema_path: Optional[str] = None
    ):
        details = {
            "error_count": len(errors) if errors else 0,
            "schema_path": schema_path
        }
        super().__init__(message, details)
        self.errors = errors or []
        self.schema_path = schema_path
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to JSON-serializable dictionary."""
        base = super().to_dict()
        base["errors"] = [e.to_dict() for e in self.errors]
        return base


class SchemaLoadError(ValidatorError):
    """Raised when schema cannot be loaded."""
    pass


class SchemaInvalidError(ValidatorError):
    """Raised when schema itself is invalid."""
    pass


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class ValidationErrorDetail:
    """
    Detailed information about a single validation error.
    """
    path: str
    message: str
    validator: str
    validator_value: Any = None
    instance: Any = None
    schema_path: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to JSON-serializable dictionary."""
        result = {
            "path": self.path,
            "message": self.message,
            "validator": self.validator,
            "schema_path": self.schema_path
        }
        
        # Include validator_value if it's JSON-serializable
        if self.validator_value is not None:
            try:
                json.dumps(self.validator_value)
                result["validator_value"] = self.validator_value
            except (TypeError, ValueError):
                result["validator_value"] = str(self.validator_value)
        
        return result
    
    def __str__(self) -> str:
        return f"{self.path or '<root>'}: {self.message}"


@dataclass
class ValidationResult:
    """
    Result of a validation operation.
    
    Provides structured access to validation outcome and any errors.
    """
    is_valid: bool
    errors: List[ValidationErrorDetail] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    schema_path: Optional[str] = None
    schema_draft: Optional[str] = None
    validated_at: str = field(default_factory=lambda: datetime.utcnow().isoformat() + "Z")
    
    @property
    def error_count(self) -> int:
        """Number of validation errors."""
        return len(self.errors)
    
    @property
    def warning_count(self) -> int:
        """Number of warnings."""
        return len(self.warnings)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to JSON-serializable dictionary."""
        return {
            "is_valid": self.is_valid,
            "error_count": self.error_count,
            "warning_count": self.warning_count,
            "errors": [e.to_dict() for e in self.errors],
            "warnings": self.warnings,
            "schema_path": self.schema_path,
            "schema_draft": self.schema_draft,
            "validated_at": self.validated_at
        }
    
    def to_json(self) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=2)
    
    def raise_if_invalid(self) -> None:
        """Raise ValidationError if validation failed."""
        if not self.is_valid:
            error_messages = [str(e) for e in self.errors]
            raise ValidationError(
                f"Validation failed with {self.error_count} error(s):\n" + 
                "\n".join(error_messages),
                errors=self.errors,
                schema_path=self.schema_path
            )


# ============================================================================
# SCHEMA CACHE
# ============================================================================

class SchemaCache:
    """
    Thread-safe schema cache for performance optimization.
    
    Caches loaded and compiled schemas to avoid repeated file I/O
    and schema compilation.
    """
    
    _instance: Optional['SchemaCache'] = None
    _schemas: Dict[str, Dict[str, Any]] = {}
    _validators: Dict[str, Any] = {}
    _mtimes: Dict[str, float] = {}
    
    def __new__(cls) -> 'SchemaCache':
        """Singleton pattern."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._schemas = {}
            cls._instance._validators = {}
            cls._instance._mtimes = {}
        return cls._instance
    
    def get_schema(self, schema_path: str, force_reload: bool = False) -> Dict[str, Any]:
        """
        Get schema from cache or load from file.
        
        Args:
            schema_path: Path to schema file
            force_reload: Force reload even if cached
            
        Returns:
            Parsed schema dictionary
        """
        path = Path(schema_path).resolve()
        path_str = str(path)
        
        # Check if reload needed
        if not force_reload and path_str in self._schemas:
            # Check if file was modified
            try:
                current_mtime = path.stat().st_mtime
                if current_mtime <= self._mtimes.get(path_str, 0):
                    return self._schemas[path_str]
            except OSError:
                pass
        
        # Load schema
        schema = self._load_schema_file(path)
        self._schemas[path_str] = schema
        
        try:
            self._mtimes[path_str] = path.stat().st_mtime
        except OSError:
            self._mtimes[path_str] = 0
        
        # Invalidate validator cache for this schema
        if path_str in self._validators:
            del self._validators[path_str]
        
        return schema
    
    def get_validator(
        self,
        schema_path: str,
        draft: Optional[str] = None
    ) -> Any:
        """
        Get compiled validator from cache or create new.
        
        Args:
            schema_path: Path to schema file
            draft: JSON Schema draft version (auto-detected if None)
            
        Returns:
            Compiled validator instance
        """
        if not JSONSCHEMA_AVAILABLE:
            raise ValidatorError(
                "jsonschema library is required for validation",
                details={"install": "pip install jsonschema"}
            )
        
        path = Path(schema_path).resolve()
        path_str = str(path)
        cache_key = f"{path_str}:{draft or 'auto'}"
        
        if cache_key in self._validators:
            return self._validators[cache_key]
        
        schema = self.get_schema(schema_path)
        validator_class = self._get_validator_class(schema, draft)
        
        # Create format checker with custom formats
        format_checker = self._create_format_checker()
        
        # Create validator
        validator = validator_class(schema, format_checker=format_checker)
        self._validators[cache_key] = validator
        
        return validator
    
    def clear(self) -> None:
        """Clear all cached schemas and validators."""
        self._schemas.clear()
        self._validators.clear()
        self._mtimes.clear()
    
    def _load_schema_file(self, path: Path) -> Dict[str, Any]:
        """Load schema from file."""
        if not path.exists():
            raise SchemaLoadError(
                f"Schema file not found: {path}",
                details={"path": str(path)}
            )
        
        try:
            content = path.read_text(encoding='utf-8')
            schema = json.loads(content)
            return schema
        except json.JSONDecodeError as e:
            raise SchemaLoadError(
                f"Invalid JSON in schema file: {path}",
                details={"path": str(path), "error": str(e)}
            )
        except OSError as e:
            raise SchemaLoadError(
                f"Cannot read schema file: {path}",
                details={"path": str(path), "error": str(e)}
            )
    
    def _get_validator_class(
        self,
        schema: Dict[str, Any],
        draft: Optional[str]
    ) -> Type:
        """Get appropriate validator class for schema."""
        if draft:
            draft_lower = draft.lower()
            if '2020' in draft_lower or '202012' in draft_lower:
                return Draft202012Validator
            elif '2019' in draft_lower or '201909' in draft_lower:
                return Draft201909Validator
            elif '7' in draft_lower or 'draft-07' in draft_lower:
                return Draft7Validator
        
        # Auto-detect from $schema
        schema_uri = schema.get('$schema', '')
        
        if '2020-12' in schema_uri or 'draft/2020-12' in schema_uri:
            return Draft202012Validator
        elif '2019-09' in schema_uri or 'draft/2019-09' in schema_uri:
            return Draft201909Validator
        elif 'draft-07' in schema_uri:
            return Draft7Validator
        
        # Default to latest
        return Draft202012Validator
    
    def _create_format_checker(self) -> FormatChecker:
        """Create format checker with custom formats."""
        checker = FormatChecker()
        
        # Hex color format
        @checker.checks('hex-color')
        def check_hex_color(value: str) -> bool:
            if not isinstance(value, str):
                return False
            pattern = r'^#?[0-9A-Fa-f]{6}$'
            return bool(re.match(pattern, value))
        
        # Percentage format
        @checker.checks('percentage')
        def check_percentage(value: str) -> bool:
            if not isinstance(value, str):
                return False
            pattern = r'^-?\d+(\.\d+)?%$'
            return bool(re.match(pattern, value))
        
        # File path format
        @checker.checks('file-path')
        def check_file_path(value: str) -> bool:
            if not isinstance(value, str):
                return False
            try:
                Path(value)
                return True
            except Exception:
                return False
        
        # Absolute path format
        @checker.checks('absolute-path')
        def check_absolute_path(value: str) -> bool:
            if not isinstance(value, str):
                return False
            return os.path.isabs(value)
        
        # Slide index format (non-negative integer)
        @checker.checks('slide-index')
        def check_slide_index(value: Any) -> bool:
            return isinstance(value, int) and value >= 0
        
        # Shape index format (non-negative integer)
        @checker.checks('shape-index')
        def check_shape_index(value: Any) -> bool:
            return isinstance(value, int) and value >= 0
        
        return checker


# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

def validate_against_schema(payload: Dict[str, Any], schema_path: str) -> None:
    """
    Strictly validate payload against JSON Schema.
    
    This is the backward-compatible function that raises ValueError on failure.
    
    Args:
        payload: Data to validate
        schema_path: Path to JSON Schema file
        
    Raises:
        ValueError: If validation fails (with detailed error messages)
        SchemaLoadError: If schema cannot be loaded
        
    Example:
        >>> validate_against_schema({"name": "test"}, "schemas/config.schema.json")
    """
    if not JSONSCHEMA_AVAILABLE:
        raise ImportError(
            "jsonschema library is required. Install with:\n"
            "  pip install jsonschema\n"
            "  or: uv pip install jsonschema"
        )
    
    result = validate_dict(payload, schema_path=schema_path)
    
    if not result.is_valid:
        error_messages = []
        for error in result.errors:
            loc = error.path or '<root>'
            error_messages.append(f"{loc}: {error.message}")
        
        raise ValueError(
            "Strict schema validation failed:\n" + "\n".join(error_messages)
        )


def validate_dict(
    data: Dict[str, Any],
    schema: Optional[Dict[str, Any]] = None,
    schema_path: Optional[str] = None,
    draft: Optional[str] = None,
    raise_on_error: bool = False
) -> ValidationResult:
    """
    Validate dictionary against JSON Schema.
    
    Either schema or schema_path must be provided.
    
    Args:
        data: Data to validate
        schema: JSON Schema dictionary
        schema_path: Path to JSON Schema file
        draft: JSON Schema draft version (auto-detected if None)
        raise_on_error: Raise ValidationError if validation fails
        
    Returns:
        ValidationResult with validation outcome
        
    Raises:
        ValidationError: If raise_on_error=True and validation fails
        SchemaLoadError: If schema cannot be loaded
        ValidatorError: If neither schema nor schema_path provided
        
    Example:
        >>> result = validate_dict(data, schema_path="schemas/manifest.json")
        >>> if not result.is_valid:
        ...     for error in result.errors:
        ...         print(error)
    """
    if not JSONSCHEMA_AVAILABLE:
        raise ValidatorError(
            "jsonschema library is required",
            details={"install": "pip install jsonschema"}
        )
    
    if schema is None and schema_path is None:
        raise ValidatorError(
            "Either schema or schema_path must be provided"
        )
    
    # Get or create validator
    cache = SchemaCache()
    
    if schema_path:
        validator = cache.get_validator(schema_path, draft)
        resolved_schema = cache.get_schema(schema_path)
    else:
        validator_class = cache._get_validator_class(schema, draft)
        format_checker = cache._create_format_checker()
        validator = validator_class(schema, format_checker=format_checker)
        resolved_schema = schema
    
    # Detect draft version
    schema_draft = resolved_schema.get('$schema', 'unknown')
    
    # Collect errors
    errors: List[ValidationErrorDetail] = []
    warnings: List[str] = []
    
    try:
        validation_errors = sorted(
            validator.iter_errors(data),
            key=lambda e: (list(e.absolute_path), e.message)
        )
        
        for error in validation_errors:
            path = "/".join(str(p) for p in error.absolute_path)
            schema_path_str = "/".join(str(p) for p in error.absolute_schema_path)
            
            errors.append(ValidationErrorDetail(
                path=path,
                message=error.message,
                validator=error.validator,
                validator_value=error.validator_value,
                instance=error.instance if _is_json_serializable(error.instance) else str(error.instance),
                schema_path=schema_path_str
            ))
    except JsonSchemaSchemaError as e:
        raise SchemaInvalidError(
            f"Invalid schema: {e.message}",
            details={"error": str(e)}
        )
    
    # Create result
    result = ValidationResult(
        is_valid=len(errors) == 0,
        errors=errors,
        warnings=warnings,
        schema_path=schema_path,
        schema_draft=schema_draft
    )
    
    if raise_on_error:
        result.raise_if_invalid()
    
    return result


def validate_json_file(
    file_path: str,
    schema_path: str,
    draft: Optional[str] = None,
    raise_on_error: bool = False
) -> ValidationResult:
    """
    Validate JSON file against schema.
    
    Args:
        file_path: Path to JSON file to validate
        schema_path: Path to JSON Schema file
        draft: JSON Schema draft version
        raise_on_error: Raise ValidationError if validation fails
        
    Returns:
        ValidationResult with validation outcome
        
    Raises:
        ValidationError: If raise_on_error=True and validation fails
        SchemaLoadError: If files cannot be loaded
    """
    path = Path(file_path)
    
    if not path.exists():
        raise SchemaLoadError(
            f"File not found: {file_path}",
            details={"path": file_path}
        )
    
    try:
        content = path.read_text(encoding='utf-8')
        data = json.loads(content)
    except json.JSONDecodeError as e:
        raise SchemaLoadError(
            f"Invalid JSON in file: {file_path}",
            details={"path": file_path, "error": str(e)}
        )
    except OSError as e:
        raise SchemaLoadError(
            f"Cannot read file: {file_path}",
            details={"path": file_path, "error": str(e)}
        )
    
    return validate_dict(
        data,
        schema_path=schema_path,
        draft=draft,
        raise_on_error=raise_on_error
    )


def load_schema(schema_path: str, force_reload: bool = False) -> Dict[str, Any]:
    """
    Load JSON Schema from file with caching.
    
    Args:
        schema_path: Path to schema file
        force_reload: Force reload from disk
        
    Returns:
        Parsed schema dictionary
    """
    cache = SchemaCache()
    return cache.get_schema(schema_path, force_reload=force_reload)


def clear_schema_cache() -> None:
    """Clear the schema cache."""
    cache = SchemaCache()
    cache.clear()


def is_valid(
    data: Dict[str, Any],
    schema: Optional[Dict[str, Any]] = None,
    schema_path: Optional[str] = None
) -> bool:
    """
    Quick validation check returning boolean.
    
    Args:
        data: Data to validate
        schema: JSON Schema dictionary
        schema_path: Path to JSON Schema file
        
    Returns:
        True if valid, False otherwise
    """
    try:
        result = validate_dict(data, schema=schema, schema_path=schema_path)
        return result.is_valid
    except Exception:
        return False


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def _is_json_serializable(value: Any) -> bool:
    """Check if value is JSON serializable."""
    try:
        json.dumps(value)
        return True
    except (TypeError, ValueError):
        return False


def get_schema_draft(schema: Dict[str, Any]) -> str:
    """
    Detect JSON Schema draft version from schema.
    
    Args:
        schema: Schema dictionary
        
    Returns:
        Draft identifier string
    """
    schema_uri = schema.get('$schema', '')
    
    if '2020-12' in schema_uri:
        return 'draft-2020-12'
    elif '2019-09' in schema_uri:
        return 'draft-2019-09'
    elif 'draft-07' in schema_uri:
        return 'draft-07'
    elif 'draft-06' in schema_uri:
        return 'draft-06'
    elif 'draft-04' in schema_uri:
        return 'draft-04'
    
    return 'unknown'


# ============================================================================
# MODULE METADATA
# ============================================================================

__version__ = "3.0.0"
__author__ = "PowerPoint Agent Team"
__license__ = "MIT"

__all__ = [
    # Main functions
    "validate_against_schema",
    "validate_dict",
    "validate_json_file",
    "load_schema",
    "clear_schema_cache",
    "is_valid",
    "get_schema_draft",
    
    # Classes
    "ValidationResult",
    "ValidationErrorDetail",
    "SchemaCache",
    
    # Exceptions
    "ValidatorError",
    "ValidationError",
    "SchemaLoadError",
    "SchemaInvalidError",
    
    # Constants
    "JSONSCHEMA_AVAILABLE",
    
    # Module metadata
    "__version__",
    "__author__",
    "__license__",
]

```

# tools/ppt_add_bullet_list.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Bullet List Tool v3.1.0
Add bullet or numbered list with 6×6 rule validation and accessibility checks.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 1 \\
        --items "Point 1,Point 2,Point 3" \\
        --position '{"left":"10%","top":"25%"}' \\
        --size '{"width":"80%","height":"60%"}' --json

Exit Codes:
    0: Success
    1: Error occurred

6×6 Rule (Best Practice):
    - Maximum 6 bullet points per slide
    - Maximum 6 words per line (~60 characters)
    - Ensures readability and audience engagement
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ColorHelper,
)
from pptx.dml.color import RGBColor

__version__ = "3.1.0"


def calculate_readability_score(items: List[str]) -> Dict[str, Any]:
    """Calculate readability metrics for bullet list."""
    total_chars = sum(len(item) for item in items)
    avg_chars = total_chars / len(items) if items else 0
    max_chars = max(len(item) for item in items) if items else 0
    
    total_words = sum(len(item.split()) for item in items)
    avg_words = total_words / len(items) if items else 0
    max_words = max(len(item.split()) for item in items) if items else 0
    
    score = 100
    issues = []
    
    if len(items) > 6:
        score -= (len(items) - 6) * 10
        issues.append(f"Exceeds 6×6 rule: {len(items)} items (recommended: ≤6)")
    
    if avg_chars > 60:
        score -= 20
        issues.append(f"Items too long: {avg_chars:.0f} chars average (recommended: ≤60)")
    
    if max_chars > 100:
        score -= 10
        issues.append(f"Longest item: {max_chars} chars (consider splitting)")
    
    if max_words > 12:
        score -= 15
        issues.append(f"Too many words per item: {max_words} max (recommended: ≤10)")
    
    score = max(0, score)
    
    return {
        "score": score,
        "grade": "A" if score >= 90 else "B" if score >= 75 else "C" if score >= 60 else "D" if score >= 50 else "F",
        "metrics": {
            "item_count": len(items),
            "avg_characters": round(avg_chars, 1),
            "max_characters": max_chars,
            "avg_words": round(avg_words, 1),
            "max_words": max_words
        },
        "issues": issues
    }


def add_bullet_list(
    filepath: Path,
    slide_index: int,
    items: List[str],
    position: Dict[str, Any],
    size: Dict[str, Any],
    bullet_style: str = "bullet",
    font_size: int = 18,
    font_name: str = "Calibri",
    color: str = None,
    line_spacing: float = 1.0,
    ignore_rules: bool = False
) -> Dict[str, Any]:
    """
    Add bullet or numbered list with validation.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Slide index (0-based)
        items: List of bullet items
        position: Position dict
        size: Size dict
        bullet_style: "bullet", "numbered", or "none"
        font_size: Font size in points
        font_name: Font name
        color: Optional text color (hex)
        line_spacing: Line spacing multiplier
        ignore_rules: Override 6×6 rule validation
        
    Returns:
        Dict with results and validation info
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If invalid parameters
        SlideNotFoundError: If slide index out of range
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    if not items:
        raise ValueError("At least one item required")
    
    warnings = []
    recommendations = []
    
    readability = calculate_readability_score(items)
    
    if len(items) > 6 and not ignore_rules:
        warnings.append(
            f"6×6 Rule violation: {len(items)} items exceeds recommended 6 per slide. "
            "This reduces readability and audience engagement."
        )
        recommendations.append(
            "Consider splitting into multiple slides or using --ignore-rules to override"
        )
    
    if len(items) > 10 and not ignore_rules:
        raise ValueError(
            f"Too many items: {len(items)} exceeds hard limit of 10 per slide. "
            "Split into multiple slides or use --ignore-rules to override."
        )
    
    for idx, item in enumerate(items):
        if len(item) > 100:
            warnings.append(
                f"Item {idx + 1} is {len(item)} characters (very long). "
                "Consider breaking into multiple bullets."
            )
    
    if font_size < 14:
        warnings.append(
            f"Font size {font_size}pt is below recommended minimum of 14pt."
        )
    
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            bg_color = RGBColor(255, 255, 255)
            is_large_text = font_size >= 18
            
            if not ColorHelper.meets_wcag(text_color, bg_color, is_large_text):
                contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
                required_ratio = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color contrast {contrast_ratio:.2f}:1 may not meet WCAG accessibility "
                    f"(required: {required_ratio}:1)."
                )
        except Exception:
            pass
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        agent.add_bullet_list(
            slide_index=slide_index,
            items=items,
            position=position,
            size=size,
            bullet_style=bullet_style,
            font_size=font_size
        )
        
        slide_info = agent.get_slide_info(slide_index)
        last_shape_idx = slide_info["shape_count"] - 1
        
        if color:
            try:
                agent.format_text(
                    slide_index=slide_index,
                    shape_index=last_shape_idx,
                    color=color
                )
            except Exception as e:
                warnings.append(f"Could not apply color: {str(e)}")
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    if readability["score"] < 75:
        recommendations.append(
            f"Readability score is {readability['grade']} ({readability['score']}/100). "
            "Consider simplifying content."
        )
    
    status = "success"
    if warnings:
        status = "warning"
    
    result: Dict[str, Any] = {
        "status": status,
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "items_added": len(items),
        "items": items,
        "bullet_style": bullet_style,
        "formatting": {
            "font_size": font_size,
            "font_name": font_name,
            "color": color,
            "line_spacing": line_spacing
        },
        "readability": readability,
        "validation": {
            "six_six_rule": {
                "compliant": len(items) <= 6 and readability["metrics"]["max_words"] <= 10,
                "item_count_ok": len(items) <= 6,
                "word_count_ok": readability["metrics"]["max_words"] <= 10
            },
            "accessibility": {
                "font_size_ok": font_size >= 14,
                "color_contrast_checked": color is not None
            }
        },
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if warnings:
        result["warnings"] = warnings
    
    if recommendations:
        result["recommendations"] = recommendations
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Add bullet/numbered list with 6×6 rule validation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
6×6 Rule (Best Practice):
  - Maximum 6 bullet points per slide
  - Maximum 6 words per line (~60 characters)
  - Ensures readability and audience engagement

Examples:
  # Simple bullet list
  uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 1 \\
    --items "Revenue up 45%,Customer growth 60%,Market share increased" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' --json

  # Numbered list
  uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 2 \\
    --items "Define objectives,Analyze market,Execute plan" \\
    --bullet-style numbered --font-size 20 --color "#0070C0" \\
    --position '{"left":"15%","top":"30%"}' \\
    --size '{"width":"70%","height":"50%"}' --json

  # From JSON file
  echo '["First point", "Second point"]' > items.json
  uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 3 \\
    --items-file items.json --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' --json
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path (.pptx)')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--items', help='Comma-separated list items')
    parser.add_argument('--items-file', type=Path, help='JSON file with array of items')
    parser.add_argument('--position', required=True, type=json.loads, help='Position dict (JSON)')
    parser.add_argument('--size', type=json.loads, help='Size dict (JSON)')
    parser.add_argument('--bullet-style', choices=['bullet', 'numbered', 'none'], default='bullet')
    parser.add_argument('--font-size', type=int, default=18, help='Font size (default: 18)')
    parser.add_argument('--font-name', default='Calibri', help='Font name')
    parser.add_argument('--color', help='Text color hex (e.g., #0070C0)')
    parser.add_argument('--line-spacing', type=float, default=1.0, help='Line spacing')
    parser.add_argument('--ignore-rules', action='store_true', help='Override 6×6 validation')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON (default: true)')
    
    args = parser.parse_args()
    
    try:
        if args.items_file:
            if not args.items_file.exists():
                raise FileNotFoundError(f"Items file not found: {args.items_file}")
            with open(args.items_file, 'r', encoding='utf-8') as f:
                items = json.load(f)
            if not isinstance(items, list):
                raise ValueError("Items file must contain JSON array")
        elif args.items:
            if '\\n' in args.items:
                items = args.items.split('\\n')
            else:
                items = args.items.split(',')
            items = [item.strip() for item in items if item.strip()]
        else:
            raise ValueError("Either --items or --items-file required")
        
        size = args.size if args.size else {}
        position = args.position
        
        if "width" not in size:
            size["width"] = position.get("width", "80%")
        if "height" not in size:
            size["height"] = position.get("height", "50%")
        
        result = add_bullet_list(
            filepath=args.file,
            slide_index=args.slide,
            items=items,
            position=position,
            size=size,
            bullet_style=args.bullet_style,
            font_size=args.font_size,
            font_name=args.font_name,
            color=args.color,
            line_spacing=args.line_spacing,
            ignore_rules=args.ignore_rules
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check items format and file extension (.pptx required)."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_chart.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Chart Tool v3.1.0
Add data visualization chart to slide

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_chart.py --file presentation.pptx --slide 1 --chart-type column --data chart_data.json --position '{"left":"10%","top":"20%"}' --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Supported Chart Types:
    column, column_stacked, bar, bar_stacked, line, line_markers,
    pie, area, scatter, doughnut
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Supported chart types
CHART_TYPES = [
    'column', 'column_stacked', 'bar', 'bar_stacked',
    'line', 'line_markers', 'pie', 'area', 'scatter', 'doughnut'
]


def add_chart(
    filepath: Path,
    slide_index: int,
    chart_type: str,
    data: Dict[str, Any],
    position: Dict[str, Any],
    size: Dict[str, Any],
    chart_title: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add a data visualization chart to a slide.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the target slide (0-based)
        chart_type: Type of chart (column, bar, line, pie, etc.)
        data: Chart data dict with 'categories' and 'series' keys
        position: Position specification dict
        size: Size specification dict
        chart_title: Optional chart title
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the added chart shape
            - chart_type: Type of chart added
            - chart_title: Title if provided
            - categories: Number of categories
            - series: Number of data series
            - data_points: Total number of data points
            - presentation_version_before: State hash before addition
            - presentation_version_after: State hash after addition
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ValueError: If data format is invalid
        
    Example:
        >>> data = {
        ...     "categories": ["Q1", "Q2", "Q3", "Q4"],
        ...     "series": [{"name": "Revenue", "values": [100, 120, 140, 160]}]
        ... }
        >>> result = add_chart(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=1,
        ...     chart_type="column",
        ...     data=data,
        ...     position={"left": "10%", "top": "20%"},
        ...     size={"width": "80%", "height": "60%"},
        ...     chart_title="Revenue Growth"
        ... )
        >>> print(result["shape_index"])
        5
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate chart type
    if chart_type not in CHART_TYPES:
        raise ValueError(
            f"Invalid chart type: {chart_type}. "
            f"Supported types: {', '.join(CHART_TYPES)}"
        )
    
    # Validate data structure
    if "categories" not in data:
        raise ValueError(
            "Data must contain 'categories' key. "
            "Example: {\"categories\": [\"Q1\", \"Q2\"], \"series\": [...]}"
        )
    
    if "series" not in data or not data["series"]:
        raise ValueError(
            "Data must contain at least one series. "
            "Example: {\"series\": [{\"name\": \"Sales\", \"values\": [10, 20]}]}"
        )
    
    # Validate all series have same length as categories
    cat_len = len(data["categories"])
    for i, series in enumerate(data["series"]):
        if "values" not in series:
            raise ValueError(f"Series {i} missing 'values' key")
        if len(series.get("values", [])) != cat_len:
            raise ValueError(
                f"Series '{series.get('name', f'[{i}]')}' has {len(series['values'])} values, "
                f"but there are {cat_len} categories. Counts must match."
            )
    
    # Validate pie chart has only one series
    if chart_type in ['pie', 'doughnut'] and len(data["series"]) > 1:
        raise ValueError(
            f"{chart_type.capitalize()} charts support only one data series. "
            f"Found {len(data['series'])} series."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE addition
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Add chart
        result = agent.add_chart(
            slide_index=slide_index,
            chart_type=chart_type,
            data=data,
            position=position,
            size=size,
            chart_title=chart_title
        )
        
        # Extract shape index from result (handle v3.0.x and v3.1.x)
        if isinstance(result, dict):
            shape_index = result.get("shape_index")
        else:
            # Fallback: get last shape index
            slide_info = agent.get_slide_info(slide_index)
            shape_index = slide_info.get("shape_count", 1) - 1
        
        # Save changes
        agent.save()
        
        # Capture version AFTER addition
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "chart_type": chart_type,
        "chart_title": chart_title,
        "categories": len(data["categories"]),
        "series": len(data["series"]),
        "data_points": sum(len(s["values"]) for s in data["series"]),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add data visualization chart to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Chart Types:
  column          Vertical bars (compare across categories)
  column_stacked  Stacked vertical bars (show composition)
  bar             Horizontal bars (compare items)
  bar_stacked     Stacked horizontal bars
  line            Line chart (show trends over time)
  line_markers    Line with data point markers
  pie             Pie chart (show proportions, single series only)
  doughnut        Doughnut chart (pie with hole, single series only)
  area            Area chart (emphasize magnitude of change)
  scatter         Scatter plot (show relationships)

Data Format (JSON file or inline):
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "Revenue", "values": [100, 120, 140, 160]},
    {"name": "Costs", "values": [80, 90, 100, 110]}
  ]
}

Examples:
  # Revenue growth chart from JSON file
  uv run tools/ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart-type column \\
    --data revenue_data.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --title "Revenue Growth Trajectory" \\
    --json
  
  # Inline data (short example)
  uv run tools/ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --chart-type column \\
    --data-string '{"categories":["A","B","C"],"series":[{"name":"Sales","values":[10,20,15]}]}' \\
    --position '{"left":"20%","top":"25%"}' \\
    --size '{"width":"60%","height":"50%"}' \\
    --json
  
  # Pie chart (single series)
  uv run tools/ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --chart-type pie \\
    --data-string '{"categories":["Us","Competitor A","Others"],"series":[{"name":"Share","values":[35,40,25]}]}' \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"60%","height":"60%"}' \\
    --title "Market Share" \\
    --json
  
  # Line chart for trends
  uv run tools/ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --chart-type line_markers \\
    --data trend_data.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"65%"}' \\
    --title "Monthly Trends" \\
    --json

Chart Selection Guide:
  Compare values across categories  → column or bar
  Show trends over time             → line or line_markers
  Show proportions/percentages      → pie or doughnut
  Show composition over time        → column_stacked or area
  Show correlation between values   → scatter

Best Practices:
  - Use column charts for most comparisons
  - Limit pie charts to 5-7 slices maximum
  - Use line charts for time series data
  - Keep series count to 3-5 for readability
  - Always include a descriptive title
  - Round numbers for better readability

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 1,
    "shape_index": 5,
    "chart_type": "column",
    "chart_title": "Revenue Growth",
    "categories": 4,
    "series": 2,
    "data_points": 8,
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--chart-type',
        required=True,
        choices=CHART_TYPES,
        help='Chart type'
    )
    
    parser.add_argument(
        '--data',
        type=Path,
        help='JSON file with chart data'
    )
    
    parser.add_argument(
        '--data-string',
        help='Inline JSON data string'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=str,
        help='Position dict as JSON string'
    )
    
    parser.add_argument(
        '--size',
        type=str,
        help='Size dict as JSON string'
    )
    
    parser.add_argument(
        '--title',
        help='Chart title'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse position JSON
        try:
            position = json.loads(args.position)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in --position: {e}")
        
        # Load chart data
        if args.data:
            if not args.data.exists():
                raise FileNotFoundError(f"Data file not found: {args.data}")
            with open(args.data, 'r') as f:
                data = json.load(f)
        elif args.data_string:
            try:
                data = json.loads(args.data_string)
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSON in --data-string: {e}")
        else:
            raise ValueError("Either --data or --data-string is required")
        
        # Parse size JSON or set defaults
        size: Dict[str, Any] = {}
        if args.size:
            try:
                size = json.loads(args.size)
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSON in --size: {e}")
        
        # Handle size from position if not specified
        if "width" in position and "width" not in size:
            size["width"] = position["width"]
        if "height" in position and "height" not in size:
            size["height"] = position["height"]
        
        # Apply defaults
        if "width" not in size:
            size["width"] = "50%"
        if "height" not in size:
            size["height"] = "50%"
        
        result = add_chart(
            filepath=args.file,
            slide_index=args.slide,
            chart_type=args.chart_type,
            data=data,
            position=position,
            size=size,
            chart_title=args.title
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file paths exist and are accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check data format and JSON syntax"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_connector.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Connector Tool v3.1.0
Draw a line/connector between two shapes on a slide

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_connector.py --file deck.pptx --slide 0 --from-shape 0 --to-shape 1 --type straight --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Use Cases:
    - Flowcharts and process diagrams
    - Org charts
    - Network diagrams
    - Relationship mapping
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Supported connector types
CONNECTOR_TYPES = ['straight', 'elbow', 'curve']

# Define fallback exception
try:
    from core.powerpoint_agent_core import ShapeNotFoundError
except ImportError:
    class ShapeNotFoundError(PowerPointAgentError):
        """Exception raised when shape is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def add_connector(
    filepath: Path,
    slide_index: int,
    from_shape: int,
    to_shape: int,
    connector_type: str = "straight",
    line_color: Optional[str] = None,
    line_width: Optional[float] = None
) -> Dict[str, Any]:
    """
    Add a connector line between two shapes on a slide.
    
    Creates a line that visually connects two shapes, useful for
    flowcharts, org charts, and process diagrams.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the shapes (0-based)
        from_shape: Index of the starting shape (0-based)
        to_shape: Index of the ending shape (0-based)
        connector_type: Type of connector ('straight', 'elbow', 'curve')
        line_color: Optional line color in hex format (e.g., "#000000")
        line_width: Optional line width in points
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the new connector shape
            - connection: Dict with from, to, and type info
            - presentation_version_before: State hash before addition
            - presentation_version_after: State hash after addition
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ShapeNotFoundError: If from_shape or to_shape index is invalid
        ValueError: If connector type is invalid
        
    Example:
        >>> result = add_connector(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     from_shape=0,
        ...     to_shape=1,
        ...     connector_type="straight"
        ... )
        >>> print(result["shape_index"])
        5
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate connector type
    if connector_type not in CONNECTOR_TYPES:
        raise ValueError(
            f"Invalid connector type: {connector_type}. "
            f"Supported types: {', '.join(CONNECTOR_TYPES)}"
        )
    
    # Validate from and to are different
    if from_shape == to_shape:
        raise ValueError(
            "Cannot connect a shape to itself. "
            "from_shape and to_shape must be different."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE addition
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Get slide info to validate shape indices
        slide_info = agent.get_slide_info(slide_index)
        shape_count = slide_info.get("shape_count", 0)
        
        # Validate from_shape
        if not 0 <= from_shape < shape_count:
            raise ShapeNotFoundError(
                f"from_shape index {from_shape} out of range (0-{shape_count - 1})",
                details={
                    "requested_index": from_shape,
                    "available_shapes": shape_count,
                    "parameter": "from_shape"
                }
            )
        
        # Validate to_shape
        if not 0 <= to_shape < shape_count:
            raise ShapeNotFoundError(
                f"to_shape index {to_shape} out of range (0-{shape_count - 1})",
                details={
                    "requested_index": to_shape,
                    "available_shapes": shape_count,
                    "parameter": "to_shape"
                }
            )
        
        # Add connector
        result = agent.add_connector(
            slide_index=slide_index,
            from_shape=from_shape,
            to_shape=to_shape,
            connector_type=connector_type,
            line_color=line_color,
            line_width=line_width
        )
        
        # Extract shape index from result
        if isinstance(result, dict):
            connector_index = result.get("shape_index", result.get("connector_index"))
        else:
            # Fallback: new shape is at end
            updated_info = agent.get_slide_info(slide_index)
            connector_index = updated_info.get("shape_count", 1) - 1
        
        # Save changes
        agent.save()
        
        # Capture version AFTER addition
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": connector_index,
        "connection": {
            "from_shape": from_shape,
            "to_shape": to_shape,
            "type": connector_type,
            "line_color": line_color,
            "line_width": line_width
        },
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add connector line between shapes in PowerPoint",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Connector Types:
  straight  Direct line between shapes (default)
  elbow     Right-angle connector (90-degree bends)
  curve     Curved/bezier connector

Examples:
  # Simple straight connector
  uv run tools/ppt_add_connector.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --from-shape 0 \\
    --to-shape 1 \\
    --json
  
  # Elbow connector with styling
  uv run tools/ppt_add_connector.py \\
    --file flowchart.pptx \\
    --slide 2 \\
    --from-shape 3 \\
    --to-shape 5 \\
    --type elbow \\
    --color "#0070C0" \\
    --width 2.0 \\
    --json
  
  # Curved connector
  uv run tools/ppt_add_connector.py \\
    --file diagram.pptx \\
    --slide 1 \\
    --from-shape 0 \\
    --to-shape 2 \\
    --type curve \\
    --json

Finding Shape Indices:
  Use ppt_get_slide_info.py to identify shape indices:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Use Cases:
  - Flowcharts: Connect process steps
  - Org charts: Connect hierarchy levels
  - Network diagrams: Show connections
  - Mind maps: Connect ideas
  - Process flows: Show sequence

Best Practices:
  - Use straight connectors for simple diagrams
  - Use elbow connectors for flowcharts (cleaner appearance)
  - Use curved connectors for org charts
  - Keep connector colors consistent with theme
  - Add shapes before connecting them

⚠️ Shape Index Warning:
  After adding a connector, shape indices may change.
  Always refresh shape indices using ppt_get_slide_info.py
  before performing additional operations.

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "shape_index": 5,
    "connection": {
      "from_shape": 0,
      "to_shape": 1,
      "type": "straight"
    },
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    parser.add_argument(
        '--slide', 
        required=True, 
        type=int, 
        help='Slide index (0-based)'
    )
    parser.add_argument(
        '--from-shape', 
        required=True, 
        type=int, 
        help='Starting shape index (0-based)'
    )
    parser.add_argument(
        '--to-shape', 
        required=True, 
        type=int, 
        help='Ending shape index (0-based)'
    )
    parser.add_argument(
        '--type', 
        choices=CONNECTOR_TYPES,
        default='straight', 
        help='Connector type (default: straight)'
    )
    parser.add_argument(
        '--color',
        help='Line color in hex format (e.g., "#000000")'
    )
    parser.add_argument(
        '--width',
        type=float,
        help='Line width in points'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_connector(
            filepath=args.file,
            slide_index=args.slide,
            from_shape=args.from_shape,
            to_shape=args.to_shape,
            connector_type=args.type,
            line_color=args.color,
            line_width=args.width
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": f"Check connector type (supported: {', '.join(CONNECTOR_TYPES)})"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_notes.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Speaker Notes Tool v3.1.0
Add, append, or overwrite speaker notes for a specific slide.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "Key talking point" --json
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "New script" --mode overwrite --json
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "IMPORTANT:" --mode prepend --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
)

__version__ = "3.1.0"


def add_notes(
    filepath: Path,
    slide_index: int,
    text: str,
    mode: str = "append"
) -> Dict[str, Any]:
    """
    Add speaker notes to a slide.
    
    Args:
        filepath: Path to PowerPoint file (.pptx only)
        slide_index: Index of slide to modify (0-based)
        text: Text content to add to speaker notes
        mode: Insertion mode - 'append' (default), 'prepend', or 'overwrite'
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute path to file
            - slide_index: Target slide index
            - mode: Mode that was used
            - original_length: Character count of original notes
            - new_length: Character count of final notes
            - preview: First 100 characters of final notes
            - presentation_version_before: Version hash before changes
            - presentation_version_after: Version hash after changes
            - tool_version: Tool version string
            
    Raises:
        FileNotFoundError: If PowerPoint file doesn't exist
        ValueError: If file format is invalid, text is empty, or mode is invalid
        SlideNotFoundError: If slide index is out of range
        PowerPointAgentError: If notes slide cannot be accessed
    """
    if filepath.suffix.lower() != '.pptx':
        raise ValueError(
            f"Invalid file format '{filepath.suffix}'. Only .pptx files are supported."
        )

    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not text or not text.strip():
        raise ValueError("Notes text cannot be empty")
    
    if mode not in ('append', 'prepend', 'overwrite'):
        raise ValueError(
            f"Invalid mode '{mode}'. Must be 'append', 'prepend', or 'overwrite'."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        slide_count = agent.get_slide_count()
        
        if slide_count == 0:
            raise PowerPointAgentError("Presentation has no slides")
        
        if not 0 <= slide_index < slide_count:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{slide_count - 1})",
                details={"requested": slide_index, "available": slide_count}
            )
            
        slide = agent.prs.slides[slide_index]
        
        try:
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
        except Exception as e:
            raise PowerPointAgentError(f"Failed to access notes slide: {str(e)}")
        
        original_text = text_frame.text if text_frame.text else ""
        
        if mode == "overwrite":
            final_text = text
        elif mode == "append":
            if original_text and original_text.strip():
                final_text = original_text + "\n" + text
            else:
                final_text = text
        elif mode == "prepend":
            if original_text and original_text.strip():
                final_text = text + "\n" + original_text
            else:
                final_text = text
        
        text_frame.text = final_text
                
        agent.save()
        
        version_after = agent.get_presentation_version()
        
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "mode": mode,
        "original_length": len(original_text),
        "new_length": len(final_text),
        "preview": final_text[:100] + "..." if len(final_text) > 100 else final_text,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add speaker notes to a PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Append notes (default mode)
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 \\
        --text "Key talking point: Emphasize Q4 growth." --json
    
    # Overwrite existing notes
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 \\
        --text "Complete new script for this slide." --mode overwrite --json
    
    # Prepend notes (add before existing)
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 \\
        --text "IMPORTANT: Start with customer story." --mode prepend --json

Modes:
    append    - Add text after existing notes (default)
    prepend   - Add text before existing notes
    overwrite - Replace all existing notes with new text

Use Cases:
    - Presentation scripting and speaker preparation
    - Accessibility: text alternatives for complex visuals
    - Documentation: embedding context for future editors
    - Training: detailed explanations not shown on slides
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='Path to PowerPoint file (.pptx)'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--text',
        required=True,
        help='Notes content to add'
    )
    
    parser.add_argument(
        '--mode',
        choices=['append', 'prepend', 'overwrite'],
        default='append',
        help='Insertion mode: append (default), prepend, or overwrite'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output as JSON (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_notes(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            mode=args.mode
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide count."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check that file is .pptx format, text is not empty, and mode is valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Verify the file is not corrupted and the slide structure is valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_shape.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Shape Tool v3.1.0
Add shapes (rectangle, circle, arrow, etc.) to slides with comprehensive styling options.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --position '{"left":"20%","top":"30%"}' \\
        --size '{"width":"60%","height":"40%"}' --fill-color "#0070C0" --json

    # Overlay with opacity
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --position '{"left":"0%","top":"0%"}' \\
        --size '{"width":"100%","height":"100%"}' \\
        --fill-color "#000000" --fill-opacity 0.15 --json

    # Quick overlay preset
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --overlay --fill-color "#FFFFFF" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
    ColorHelper,
)

__version__ = "3.1.0"

AVAILABLE_SHAPES = [
    "rectangle",
    "rounded_rectangle",
    "ellipse",
    "oval",
    "triangle",
    "arrow_right",
    "arrow_left",
    "arrow_up",
    "arrow_down",
    "diamond",
    "pentagon",
    "hexagon",
    "star",
    "heart",
    "lightning",
    "sun",
    "moon",
    "cloud",
]

SHAPE_ALIASES = {
    "rect": "rectangle",
    "round_rect": "rounded_rectangle",
    "circle": "ellipse",
    "arrow": "arrow_right",
    "right_arrow": "arrow_right",
    "left_arrow": "arrow_left",
    "up_arrow": "arrow_up",
    "down_arrow": "arrow_down",
    "5_point_star": "star",
}

COLOR_PRESETS = {
    "primary": "#0070C0",
    "secondary": "#595959",
    "accent": "#ED7D31",
    "success": "#70AD47",
    "warning": "#FFC000",
    "danger": "#C00000",
    "white": "#FFFFFF",
    "black": "#000000",
    "light_gray": "#D9D9D9",
    "dark_gray": "#404040",
}

OVERLAY_DEFAULTS = {
    "position": {"left": "0%", "top": "0%"},
    "size": {"width": "100%", "height": "100%"},
    "fill_opacity": 0.15,
    "z_order_action": "send_to_back",
}


def resolve_shape_type(shape_type: str) -> str:
    """Resolve shape type, handling aliases."""
    shape_lower = shape_type.lower().strip()
    
    if shape_lower in SHAPE_ALIASES:
        return SHAPE_ALIASES[shape_lower]
    
    if shape_lower in AVAILABLE_SHAPES:
        return shape_lower
    
    for available in AVAILABLE_SHAPES:
        if shape_lower in available or available in shape_lower:
            return available
    
    return shape_lower


def resolve_color(color: Optional[str]) -> Optional[str]:
    """Resolve color, handling presets and validation."""
    if color is None:
        return None
    
    color_lower = color.lower().strip()
    
    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]
    
    if not color.startswith('#') and len(color) == 6:
        try:
            int(color, 16)
            return f"#{color}"
        except ValueError:
            pass
    
    return color


def validate_opacity(
    fill_opacity: float,
    line_opacity: float
) -> Tuple[List[str], List[str]]:
    """Validate opacity values and return warnings/recommendations."""
    warnings: List[str] = []
    recommendations: List[str] = []
    
    if not 0.0 <= fill_opacity <= 1.0:
        raise ValueError(
            f"fill_opacity must be between 0.0 and 1.0, got {fill_opacity}"
        )
    
    if not 0.0 <= line_opacity <= 1.0:
        raise ValueError(
            f"line_opacity must be between 0.0 and 1.0, got {line_opacity}"
        )
    
    if fill_opacity == 0.0:
        warnings.append(
            "Fill opacity is 0.0 (fully transparent). Shape fill will be invisible."
        )
    elif fill_opacity < 0.05:
        warnings.append(
            f"Fill opacity {fill_opacity} is extremely low (<5%). Shape may be nearly invisible."
        )
    
    if line_opacity == 0.0 and fill_opacity == 0.0:
        warnings.append(
            "Both fill and line opacity are 0.0. Shape will be completely invisible."
        )
    
    if 0.1 <= fill_opacity <= 0.3:
        recommendations.append(
            f"Opacity {fill_opacity} is appropriate for overlay backgrounds. "
            "Remember to use ppt_set_z_order.py --action send_to_back after adding."
        )
    
    return warnings, recommendations


def validate_shape_params(
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: Optional[str] = None,
    fill_opacity: float = 1.0,
    line_color: Optional[str] = None,
    line_opacity: float = 1.0,
    text: Optional[str] = None,
    allow_offslide: bool = False,
    is_overlay: bool = False
) -> Dict[str, Any]:
    """Validate shape parameters and return warnings/recommendations."""
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    opacity_warnings, opacity_recommendations = validate_opacity(fill_opacity, line_opacity)
    warnings.extend(opacity_warnings)
    recommendations.extend(opacity_recommendations)
    
    validation_results["fill_opacity"] = fill_opacity
    validation_results["line_opacity"] = line_opacity
    validation_results["effective_fill_transparency"] = round(1.0 - fill_opacity, 2)
    
    if position:
        try:
            for key in ["left", "top"]:
                if key in position:
                    value_str = str(position[key])
                    if value_str.endswith('%'):
                        pct = float(value_str.rstrip('%'))
                        if not allow_offslide and (pct < 0 or pct > 100):
                            warnings.append(
                                f"Position '{key}' is {pct}% which is outside slide bounds (0-100%)."
                            )
        except (ValueError, TypeError):
            pass
    
    if size:
        try:
            for key in ["width", "height"]:
                if key in size:
                    value_str = str(size[key])
                    if value_str.endswith('%'):
                        pct = float(value_str.rstrip('%'))
                        if pct <= 0:
                            warnings.append(f"Size '{key}' is {pct}% which is invalid (must be > 0%).")
                        elif pct < 1:
                            warnings.append(f"Size '{key}' is {pct}% which is extremely small (<1%).")
        except (ValueError, TypeError):
            pass
    
    if fill_color:
        try:
            from pptx.dml.color import RGBColor
            shape_rgb = ColorHelper.from_hex(fill_color)
            validation_results["fill_color_hex"] = fill_color
            
            if fill_opacity < 1.0:
                effective_r = int(fill_opacity * shape_rgb.red + (1 - fill_opacity) * 255)
                effective_g = int(fill_opacity * shape_rgb.green + (1 - fill_opacity) * 255)
                effective_b = int(fill_opacity * shape_rgb.blue + (1 - fill_opacity) * 255)
                validation_results["effective_color_on_white"] = {
                    "hex": f"#{effective_r:02X}{effective_g:02X}{effective_b:02X}"
                }
        except Exception as e:
            validation_results["color_validation_error"] = str(e)
    
    if text and fill_opacity < 0.5:
        warnings.append(
            f"Shape has text but fill opacity is only {fill_opacity}. "
            "Text may be hard to read against varied backgrounds."
        )
    
    if is_overlay:
        validation_results["is_overlay"] = True
        
        if fill_opacity > 0.3:
            warnings.append(
                f"Overlay opacity {fill_opacity} is relatively high (>30%). "
                "System prompt recommends 0.15 for subtle overlays."
            )
        
        recommendations.append(
            "IMPORTANT: After adding this overlay, run ppt_set_z_order.py "
            "--action send_to_back to ensure overlay appears behind content."
        )
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def add_shape(
    filepath: Path,
    slide_index: int,
    shape_type: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: Optional[str] = None,
    fill_opacity: float = 1.0,
    line_color: Optional[str] = None,
    line_opacity: float = 1.0,
    line_width: float = 1.0,
    text: Optional[str] = None,
    allow_offslide: bool = False,
    is_overlay: bool = False
) -> Dict[str, Any]:
    """
    Add shape to slide with comprehensive validation and opacity support.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        shape_type: Type of shape to add
        position: Position specification dict
        size: Size specification dict
        fill_color: Fill color (hex or preset name)
        fill_opacity: Fill opacity (0.0=transparent to 1.0=opaque)
        line_color: Line/border color (hex or preset name)
        line_opacity: Line/border opacity (0.0=transparent to 1.0=opaque)
        line_width: Line width in points
        text: Optional text to add inside shape
        allow_offslide: Allow positioning outside slide bounds
        is_overlay: Whether this is an overlay shape
        
    Returns:
        Result dict with shape details and validation info
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format invalid or opacity out of range
        SlideNotFoundError: If slide index is invalid
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    resolved_shape = resolve_shape_type(shape_type)
    resolved_fill = resolve_color(fill_color)
    resolved_line = resolve_color(line_color)
    
    if is_overlay:
        if not position:
            position = OVERLAY_DEFAULTS["position"].copy()
        if not size:
            size = OVERLAY_DEFAULTS["size"].copy()
        if fill_opacity == 1.0:
            fill_opacity = OVERLAY_DEFAULTS["fill_opacity"]
    
    validation = validate_shape_params(
        position=position,
        size=size,
        fill_color=resolved_fill,
        fill_opacity=fill_opacity,
        line_color=resolved_line,
        line_opacity=line_opacity,
        text=text,
        allow_offslide=allow_offslide,
        is_overlay=is_overlay
    )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        version_before = agent.get_presentation_version()
        
        add_result = agent.add_shape(
            slide_index=slide_index,
            shape_type=resolved_shape,
            position=position,
            size=size,
            fill_color=resolved_fill,
            fill_opacity=fill_opacity,
            line_color=resolved_line,
            line_opacity=line_opacity,
            line_width=line_width,
            text=text
        )
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    shape_index = add_result.get("shape_index") if isinstance(add_result, dict) else add_result
    
    result: Dict[str, Any] = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_type": resolved_shape,
        "shape_type_requested": shape_type,
        "shape_index": shape_index,
        "position": add_result.get("position", position) if isinstance(add_result, dict) else position,
        "size": add_result.get("size", size) if isinstance(add_result, dict) else size,
        "styling": {
            "fill_color": resolved_fill,
            "fill_opacity": fill_opacity,
            "fill_transparency": round(1.0 - fill_opacity, 2),
            "line_color": resolved_line,
            "line_opacity": line_opacity,
            "line_width": line_width
        },
        "text": text,
        "is_overlay": is_overlay,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if validation["validation_results"]:
        result["validation"] = validation["validation_results"]
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
    
    notes = [
        "Shape added to top of z-order (in front of existing shapes).",
        f"Shape index {shape_index} may change if other shapes are added/removed."
    ]
    
    if is_overlay or fill_opacity < 1.0:
        notes.insert(1, "Use ppt_set_z_order.py --action send_to_back to move overlay behind content.")
    
    result["notes"] = notes
    
    if is_overlay:
        result["next_step"] = {
            "command": "ppt_set_z_order.py",
            "args": {
                "--file": str(filepath.resolve()),
                "--slide": slide_index,
                "--shape": shape_index,
                "--action": "send_to_back"
            },
            "description": "Send overlay to back so it appears behind content"
        }
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Add shape to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
AVAILABLE SHAPES:
  Basic:        rectangle, rounded_rectangle, ellipse/oval, triangle, diamond
  Arrows:       arrow_right, arrow_left, arrow_up, arrow_down
  Polygons:     pentagon, hexagon
  Decorative:   star, heart, lightning, sun, moon, cloud

SHAPE ALIASES:
  rect -> rectangle, circle -> ellipse, arrow -> arrow_right

OPACITY/TRANSPARENCY:
  --fill-opacity 1.0    Fully opaque (default)
  --fill-opacity 0.5    50% transparent
  --fill-opacity 0.15   85% transparent (subtle overlay, recommended)
  --fill-opacity 0.0    Fully transparent (invisible)

OVERLAY MODE (--overlay):
  Quick preset for creating background overlays:
  - Full-slide position and size
  - 15% opacity (subtle, non-competing)
  - Reminder to use ppt_set_z_order.py after

COLOR PRESETS:
  primary (#0070C0)    secondary (#595959)    accent (#ED7D31)
  success (#70AD47)    warning (#FFC000)      danger (#C00000)
  white (#FFFFFF)      black (#000000)

EXAMPLES:

  # Semi-transparent callout box
  uv run tools/ppt_add_shape.py --file deck.pptx --slide 0 --shape rounded_rectangle \\
    --position '{"left":"10%","top":"15%"}' --size '{"width":"30%","height":"15%"}' \\
    --fill-color primary --fill-opacity 0.8 --text "Key Point" --json

  # Subtle white overlay for text readability
  uv run tools/ppt_add_shape.py --file deck.pptx --slide 2 --shape rectangle \\
    --position '{"left":"0%","top":"0%"}' --size '{"width":"100%","height":"100%"}' \\
    --fill-color "#FFFFFF" --fill-opacity 0.15 --json

  # Quick overlay using --overlay preset
  uv run tools/ppt_add_shape.py --file deck.pptx --slide 3 --shape rectangle \\
    --overlay --fill-color black --json

Z-ORDER (LAYERING):
  Shapes are added on TOP of existing shapes by default.
  For overlays, you MUST send them to back:
    1. Add the overlay shape
    2. Note the shape_index from the output
    3. Run: ppt_set_z_order.py --file FILE --slide N --shape INDEX --action send_to_back
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path (.pptx)')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, help='Shape type')
    parser.add_argument('--position', type=json.loads, default={}, help='Position dict as JSON')
    parser.add_argument('--size', type=json.loads, help='Size dict as JSON')
    parser.add_argument('--fill-color', help='Fill color: hex or preset name')
    parser.add_argument('--fill-opacity', type=float, default=1.0, help='Fill opacity (0.0-1.0)')
    parser.add_argument('--line-color', help='Line/border color')
    parser.add_argument('--line-opacity', type=float, default=1.0, help='Line opacity (0.0-1.0)')
    parser.add_argument('--line-width', type=float, default=1.0, help='Line width in points')
    parser.add_argument('--text', help='Text to add inside shape')
    parser.add_argument('--overlay', action='store_true', help='Overlay preset: full-slide, 15% opacity')
    parser.add_argument('--allow-offslide', action='store_true', help='Allow off-slide positioning')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON (default: true)')
    
    args = parser.parse_args()
    
    try:
        size = args.size if args.size else {}
        position = args.position if args.position else {}
        
        if "width" in position and "width" not in size:
            size["width"] = position.pop("width")
        if "height" in position and "height" not in size:
            size["height"] = position.pop("height")
        
        if args.overlay:
            if "left" not in position:
                position["left"] = "0%"
            if "top" not in position:
                position["top"] = "0%"
            if "width" not in size:
                size["width"] = "100%"
            if "height" not in size:
                size["height"] = "100%"
        else:
            if "width" not in size:
                size["width"] = "20%"
            if "height" not in size:
                size["height"] = "20%"
        
        result = add_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_type=args.shape,
            position=position,
            size=size,
            fill_color=args.fill_color,
            fill_opacity=args.fill_opacity,
            line_color=args.line_color,
            line_opacity=args.line_opacity,
            line_width=args.line_width,
            text=args.text,
            allow_offslide=args.allow_offslide,
            is_overlay=args.overlay
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check opacity values are between 0.0 and 1.0, and file is .pptx format."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON: {e}",
            "error_type": "JSONDecodeError",
            "suggestion": "Ensure --position and --size are valid JSON strings."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check the presentation file is valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_slide.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Slide Tool v3.1.0
Add new slide to existing presentation with specific layout

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Compatible with PowerPoint Agent Core v3.1.0 (Dictionary Returns)

Usage:
    uv run tools/ppt_add_slide.py --file presentation.pptx --layout "Title and Content" --index 2 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError
)

__version__ = "3.1.0"

# Define fallback exceptions if not in core
try:
    from core.powerpoint_agent_core import LayoutNotFoundError
except ImportError:
    class LayoutNotFoundError(PowerPointAgentError):
        """Exception raised when layout is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def add_slide(
    filepath: Path,
    layout: str,
    index: Optional[int] = None,
    set_title: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add a new slide to a presentation.
    
    Handles the v3.1.0 Core API where add_slide returns a dictionary
    with slide_index and version information.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        layout: Layout name for the new slide (fuzzy matching supported)
        index: Position to insert slide (0-based, default: end of presentation)
        set_title: Optional title text to set on the new slide
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the new slide
            - layout: Actual layout name used
            - title_set: Title text if provided
            - title_set_success: Whether title was set successfully
            - total_slides: Total slide count after addition
            - slide_info: Shape count and notes info
            - presentation_version_before: State hash before addition
            - presentation_version_after: State hash after addition
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        LayoutNotFoundError: If layout is not found
        
    Example:
        >>> result = add_slide(
        ...     filepath=Path("presentation.pptx"),
        ...     layout="Title and Content",
        ...     set_title="Q4 Results"
        ... )
        >>> print(result["slide_index"])
        5
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Get available layouts for validation
        available_layouts = agent.get_available_layouts()
        
        # Validate layout with fuzzy matching
        matched_layout = layout
        if layout not in available_layouts:
            layout_lower = layout.lower()
            match_found = False
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    matched_layout = avail
                    match_found = True
                    break
            
            if not match_found:
                raise LayoutNotFoundError(
                    f"Layout '{layout}' not found. Available layouts: {available_layouts}",
                    details={
                        "requested_layout": layout,
                        "available_layouts": available_layouts
                    }
                )
        
        # Add slide (Core v3.1.0 returns a dict)
        add_result = agent.add_slide(layout_name=matched_layout, index=index)
        
        # Extract the integer index from the returned dictionary
        # Core v3.1.0 returns dict, older versions may return int
        if isinstance(add_result, dict):
            slide_index = add_result["slide_index"]
            version_before = add_result.get("presentation_version_before")
        else:
            slide_index = add_result
            version_before = None
        
        # Set title if provided
        title_set_result = None
        title_set_success = False
        if set_title:
            try:
                title_set_result = agent.set_title(slide_index, set_title)
                if isinstance(title_set_result, dict):
                    title_set_success = title_set_result.get("title_set", False)
                else:
                    title_set_success = True
            except Exception:
                title_set_success = False
        
        # Get slide info before saving (for verification)
        slide_info = agent.get_slide_info(slide_index)
        
        # Save the file
        agent.save()
        
        # Get updated presentation info (includes final version hash)
        prs_info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "layout": matched_layout,
        "title_set": set_title,
        "title_set_success": title_set_success,
        "total_slides": prs_info["slide_count"],
        "slide_info": {
            "shape_count": slide_info.get("shape_count", 0),
            "has_notes": slide_info.get("has_notes", False)
        },
        "presentation_version_before": version_before,
        "presentation_version_after": prs_info.get("presentation_version"),
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add new slide to PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Add slide at end
  uv run tools/ppt_add_slide.py \\
    --file presentation.pptx \\
    --layout "Title and Content" \\
    --json
  
  # Add slide at specific position
  uv run tools/ppt_add_slide.py \\
    --file deck.pptx \\
    --layout "Section Header" \\
    --index 2 \\
    --json
  
  # Add slide with title
  uv run tools/ppt_add_slide.py \\
    --file presentation.pptx \\
    --layout "Title Slide" \\
    --title "Q4 Results" \\
    --json

Common Layouts:
  - Title Slide
  - Title and Content
  - Section Header
  - Two Content
  - Comparison
  - Title Only
  - Blank

Layout Matching:
  The tool supports fuzzy matching:
  - Exact match first
  - Then substring match (case-insensitive)
  
  Example: "content" will match "Title and Content"

Finding Available Layouts:
  Use ppt_get_info.py to list layouts:
  uv run tools/ppt_get_info.py --file presentation.pptx --json | jq '.layouts'

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 5,
    "layout": "Title and Content",
    "title_set": "Q4 Results",
    "title_set_success": true,
    "total_slides": 6,
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--layout',
        required=True,
        help='Layout name for new slide (fuzzy matching supported)'
    )
    
    parser.add_argument(
        '--index',
        type=int,
        help='Position to insert slide (0-based, default: end)'
    )
    
    parser.add_argument(
        '--title',
        help='Optional title text to set on new slide'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_slide(
            filepath=args.file,
            layout=args.layout,
            index=args.index,
            set_title=args.title
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except LayoutNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "LayoutNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to list available layouts"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_table.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Table Tool v3.1.0
Add data table to slide with comprehensive validation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_table.py --file presentation.pptx --slide 1 --rows 5 --cols 3 \\
        --data table_data.json --position '{"left":"10%","top":"25%"}' \\
        --size '{"width":"80%","height":"50%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
)

__version__ = "3.1.0"


def validate_table_params(
    rows: int,
    cols: int,
    position: Dict[str, Any],
    size: Dict[str, Any],
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Validate table parameters and return warnings/recommendations.
    
    Args:
        rows: Number of rows
        cols: Number of columns
        position: Position specification dict
        size: Size specification dict
        allow_offslide: Whether to allow off-slide positioning
        
    Returns:
        Dict with warnings, recommendations, and validation_results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    if position:
        try:
            if "left" in position:
                left_str = str(position["left"])
                if left_str.endswith('%'):
                    left_pct = float(left_str.rstrip('%'))
                    if (left_pct < 0 or left_pct > 100) and not allow_offslide:
                        warnings.append(
                            f"Left position {left_pct}% is outside slide bounds (0-100%). "
                            "Table may not be visible. Use --allow-offslide if intentional."
                        )
            
            if "top" in position:
                top_str = str(position["top"])
                if top_str.endswith('%'):
                    top_pct = float(top_str.rstrip('%'))
                    if (top_pct < 0 or top_pct > 100) and not allow_offslide:
                        warnings.append(
                            f"Top position {top_pct}% is outside slide bounds (0-100%). "
                            "Table may not be visible. Use --allow-offslide if intentional."
                        )
        except (ValueError, TypeError):
            pass
    
    if size:
        try:
            if "height" in size:
                height_str = str(size["height"])
                if height_str.endswith('%'):
                    height_pct = float(height_str.rstrip('%'))
                    min_height = rows * 2
                    if height_pct < min_height:
                        warnings.append(
                            f"Table height {height_pct}% is very small for {rows} rows "
                            f"(recommended: >{min_height}%). Text may be unreadable."
                        )
            
            if "width" in size:
                width_str = str(size["width"])
                if width_str.endswith('%'):
                    width_pct = float(width_str.rstrip('%'))
                    min_width = cols * 5
                    if width_pct < min_width:
                        warnings.append(
                            f"Table width {width_pct}% is very small for {cols} columns "
                            f"(recommended: >{min_width}%). Text may be unreadable."
                        )
        except (ValueError, TypeError):
            pass
            
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results
    }


def add_table(
    filepath: Path,
    slide_index: int,
    rows: int,
    cols: int,
    position: Dict[str, Any],
    size: Dict[str, Any],
    data: Optional[List[List[Any]]] = None,
    headers: Optional[List[str]] = None,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Add table to slide with validation.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        rows: Number of rows (including header row if headers provided)
        cols: Number of columns
        position: Position specification dict
        size: Size specification dict
        data: Optional 2D list of cell values
        headers: Optional list of header strings
        allow_offslide: Allow positioning outside slide bounds
        
    Returns:
        Dict with operation results, validation info, and version tracking
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If parameters are invalid
        SlideNotFoundError: If slide index out of range
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    validation = validate_table_params(rows, cols, position, size, allow_offslide)
    
    if rows < 1 or cols < 1:
        raise ValueError("Table must have at least 1 row and 1 column")
    
    if rows > 50 or cols > 20:
        raise ValueError("Maximum table size: 50 rows × 20 columns (readability limit)")
    
    table_data: List[List[Any]] = []
    
    if headers:
        if len(headers) != cols:
            raise ValueError(f"Headers count ({len(headers)}) must match columns ({cols})")
        table_data.append(headers)
        data_rows = rows - 1
    else:
        data_rows = rows
    
    if data:
        if len(data) > data_rows:
            raise ValueError(
                f"Too many data rows ({len(data)}) for table size ({data_rows} data rows)"
            )
        
        for row_idx, row in enumerate(data):
            if len(row) != cols:
                raise ValueError(
                    f"Data row {row_idx} has {len(row)} items, expected {cols}"
                )
            table_data.append([str(cell) for cell in row])
        
        while len(table_data) < rows:
            table_data.append([""] * cols)
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        agent.add_table(
            slide_index=slide_index,
            rows=rows,
            cols=cols,
            position=position,
            size=size,
            data=table_data if table_data else None
        )
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    result: Dict[str, Any] = {
        "status": "success" if not validation["warnings"] else "warning",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "rows": rows,
        "cols": cols,
        "has_headers": headers is not None,
        "data_rows_filled": len(data) if data else 0,
        "total_cells": rows * cols,
        "validation": validation["validation_results"],
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
        
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
        
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Add data table to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Data Format (JSON):
  - 2D array: [["A1","B1","C1"], ["A2","B2","C2"]]
  - CSV file: converted to 2D array
  - Pandas DataFrame: exported to JSON array

Examples:
  # Simple pricing table
  cat > pricing.json << 'EOF'
[
  ["Starter", "$9/mo", "Basic features"],
  ["Pro", "$29/mo", "Advanced features"],
  ["Enterprise", "$99/mo", "All features + support"]
]
EOF
  
  uv run tools/ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --rows 4 \\
    --cols 3 \\
    --headers "Plan,Price,Features" \\
    --data pricing.json \\
    --position '{"left":"15%","top":"25%"}' \\
    --size '{"width":"70%","height":"50%"}' \\
    --json
  
  # Quarterly results table
  cat > results.json << 'EOF'
[
  ["Q1", "10.5", "8.2", "2.3"],
  ["Q2", "12.8", "9.1", "3.7"],
  ["Q3", "15.2", "10.5", "4.7"],
  ["Q4", "18.6", "12.1", "6.5"]
]
EOF
  
  uv run tools/ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --rows 5 \\
    --cols 4 \\
    --headers "Quarter,Revenue,Costs,Profit" \\
    --data results.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"55%"}' \\
    --json
  
  # Comparison table (centered)
  cat > comparison.json << 'EOF'
[
  ["Speed", "Fast", "Very Fast"],
  ["Security", "Standard", "Enterprise"],
  ["Support", "Email", "24/7 Phone"]
]
EOF
  
  uv run tools/ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --rows 4 \\
    --cols 3 \\
    --headers "Feature,Basic,Premium" \\
    --data comparison.json \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"60%","height":"40%"}' \\
    --json
  
  # Empty table (for manual filling)
  uv run tools/ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 6 \\
    --rows 6 \\
    --cols 4 \\
    --headers "Name,Role,Department,Email" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json

Best Practices:
  - Keep tables under 10 rows for readability
  - Use headers for all tables
  - Align numbers right, text left
  - Use consistent decimal places
  - Highlight key values with color
  - Leave white space around table
  - Use alternating row colors for large tables

Table Size Guidelines:
  - 3-5 columns: Optimal for most presentations
  - 6-10 rows: Maximum for comfortable reading
  - Font size: 12-16pt for body, 14-18pt for headers
  - Cell padding: Leave breathing room

When to Use Tables vs Charts:
  - Use tables: Exact values matter, detailed data
  - Use charts: Show trends, comparisons, patterns
  - Use both: Table with summary chart
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path (.pptx)'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--rows',
        required=True,
        type=int,
        help='Number of rows (including header if present)'
    )
    
    parser.add_argument(
        '--cols',
        required=True,
        type=int,
        help='Number of columns'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        required=True,
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--data',
        type=Path,
        help='JSON file with 2D array of cell values'
    )
    
    parser.add_argument(
        '--data-string',
        help='Inline JSON 2D array string'
    )
    
    parser.add_argument(
        '--headers',
        help='Comma-separated header row (will be row 0)'
    )
    
    parser.add_argument(
        '--allow-offslide',
        action='store_true',
        help='Allow positioning outside slide bounds'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        headers = None
        if args.headers:
            headers = [h.strip() for h in args.headers.split(',')]
        
        data = None
        if args.data:
            if not args.data.exists():
                raise FileNotFoundError(f"Data file not found: {args.data}")
            with open(args.data, 'r', encoding='utf-8') as f:
                data = json.load(f)
        elif args.data_string:
            data = json.loads(args.data_string)
        
        if data is not None:
            if not isinstance(data, list):
                raise ValueError("Data must be a 2D array (list of lists)")
            if data and not isinstance(data[0], list):
                raise ValueError("Data must be a 2D array (list of lists)")
        
        result = add_table(
            filepath=args.file,
            slide_index=args.slide,
            rows=args.rows,
            cols=args.cols,
            position=args.position,
            size=args.size,
            data=data,
            headers=headers,
            allow_offslide=args.allow_offslide
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check table dimensions, data format, and position/size JSON."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON: {str(e)}",
            "error_type": "JSONDecodeError",
            "suggestion": "Validate JSON syntax. Use single quotes around JSON strings."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_text_box.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Text Box Tool v3.1.0
Add text box with flexible positioning, comprehensive validation, and accessibility checking.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_text_box.py --file deck.pptx --slide 0 \\
        --text "Revenue: $1.5M" --position '{"left":"20%","top":"30%"}' \\
        --size '{"width":"60%","height":"10%"}' --json

Exit Codes:
    0: Success
    1: Error occurred

Position Formats:
  1. Percentage: {"left": "20%", "top": "30%"}
  2. Inches: {"left": 2.0, "top": 3.0}
  3. Anchor: {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
  4. Grid: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ColorHelper,
)

__version__ = "3.1.0"

COLOR_PRESETS = {
    "black": "#000000",
    "white": "#FFFFFF",
    "primary": "#0070C0",
    "secondary": "#595959",
    "accent": "#ED7D31",
    "success": "#70AD47",
    "warning": "#FFC000",
    "danger": "#C00000",
    "dark_gray": "#333333",
    "light_gray": "#808080",
    "muted": "#808080",
}

FONT_PRESETS = {
    "default": "Calibri",
    "heading": "Calibri Light",
    "body": "Calibri",
    "code": "Consolas",
    "serif": "Georgia",
    "sans": "Arial",
}


def resolve_color(color: Optional[str]) -> Optional[str]:
    """
    Resolve color value, handling presets and hex formats.
    
    Args:
        color: Color specification (hex, preset name, or None)
        
    Returns:
        Resolved hex color or None
    """
    if color is None:
        return None
    
    color_lower = color.lower().strip()
    
    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]
    
    if not color.startswith('#') and len(color) == 6:
        try:
            int(color, 16)
            return f"#{color}"
        except ValueError:
            pass
    
    return color


def resolve_font(font: Optional[str]) -> str:
    """
    Resolve font name, handling presets.
    
    Args:
        font: Font name or preset
        
    Returns:
        Resolved font name
    """
    if font is None:
        return "Calibri"
    
    font_lower = font.lower().strip()
    if font_lower in FONT_PRESETS:
        return FONT_PRESETS[font_lower]
    
    return font


def validate_text_box(
    text: str,
    font_size: int,
    color: Optional[str] = None,
    position: Optional[Dict[str, Any]] = None,
    size: Optional[Dict[str, Any]] = None,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Validate text box parameters and return warnings/recommendations.
    
    Args:
        text: Text content
        font_size: Font size in points
        color: Text color hex
        position: Position specification
        size: Size specification
        allow_offslide: Allow off-slide positioning
        
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    text_length = len(text)
    line_count = text.count('\n') + 1
    
    validation_results["text_length"] = text_length
    validation_results["line_count"] = line_count
    validation_results["is_multiline"] = line_count > 1
    
    if line_count == 1 and text_length > 100:
        warnings.append(
            f"Text is {text_length} characters for single line (recommended: ≤100). "
            "Long single-line text may be hard to read."
        )
        recommendations.append("Consider breaking into multiple lines or shortening text")
    
    if line_count > 1 and text_length > 500:
        warnings.append(
            f"Multi-line text is {text_length} characters. Very long text blocks reduce readability."
        )
    
    validation_results["font_size"] = font_size
    validation_results["font_size_accessible"] = font_size >= 14
    
    if font_size < 10:
        warnings.append(
            f"Font size {font_size}pt is below minimum (10pt). Text will be very hard to read."
        )
    elif font_size < 12:
        warnings.append(
            f"Font size {font_size}pt is very small. Consider 14pt+ for projected presentations."
        )
        recommendations.append("Use 14pt or larger for projected content")
    elif font_size < 14:
        recommendations.append(
            f"Font size {font_size}pt is below recommended 14pt for projected content"
        )
    
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            from pptx.dml.color import RGBColor
            bg_color = RGBColor(255, 255, 255)
            
            is_large_text = font_size >= 18
            contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
            meets_wcag = ColorHelper.meets_wcag(text_color, bg_color, is_large_text)
            
            validation_results["color_contrast"] = {
                "ratio": round(contrast_ratio, 2),
                "wcag_aa": meets_wcag,
                "required_ratio": 3.0 if is_large_text else 4.5,
                "is_large_text": is_large_text
            }
            
            if not meets_wcag:
                required = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color contrast {contrast_ratio:.2f}:1 may not meet WCAG AA "
                    f"(required: {required}:1). Consider darker color."
                )
                recommendations.append(
                    "Use black (#000000), dark_gray (#333333), or primary (#0070C0) for better contrast"
                )
        except Exception as e:
            validation_results["color_error"] = str(e)
    
    if position:
        try:
            for key in ["left", "top"]:
                if key in position:
                    value_str = str(position[key])
                    if value_str.endswith('%'):
                        pct = float(value_str.rstrip('%'))
                        if not allow_offslide and (pct < 0 or pct > 100):
                            warnings.append(
                                f"Position '{key}' is {pct}% which is outside slide bounds (0-100%). "
                                f"Text box may not be visible."
                            )
        except (ValueError, TypeError):
            pass
    
    if size:
        try:
            if "height" in size:
                height_str = str(size["height"])
                if height_str.endswith('%'):
                    height_pct = float(height_str.rstrip('%'))
                    min_height = font_size * 0.15
                    if height_pct < min_height:
                        warnings.append(
                            f"Height {height_pct}% may be too small for {font_size}pt text. "
                            f"Consider at least {min_height:.1f}%."
                        )
            
            if "width" in size:
                width_str = str(size["width"])
                if width_str.endswith('%'):
                    width_pct = float(width_str.rstrip('%'))
                    if width_pct < 5:
                        warnings.append(
                            f"Width {width_pct}% is very narrow. Text may be excessively wrapped."
                        )
        except (ValueError, TypeError):
            pass
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def add_text_box(
    filepath: Path,
    slide_index: int,
    text: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    font_name: str = "Calibri",
    font_size: int = 18,
    bold: bool = False,
    italic: bool = False,
    color: Optional[str] = None,
    alignment: str = "left",
    vertical_alignment: str = "top",
    word_wrap: bool = True,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Add text box with comprehensive validation and formatting.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Slide index (0-based)
        text: Text content
        position: Position dict (supports %, inches, anchor, grid)
        size: Size dict
        font_name: Font name or preset
        font_size: Font size in points
        bold: Bold text
        italic: Italic text
        color: Text color (hex or preset)
        alignment: Horizontal alignment (left, center, right, justify)
        vertical_alignment: Vertical alignment (top, middle, bottom)
        word_wrap: Enable word wrap
        allow_offslide: Allow off-slide positioning
        
    Returns:
        Result dict with shape_index and validation info
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format invalid
        SlideNotFoundError: If slide index out of range
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    resolved_color = resolve_color(color)
    resolved_font = resolve_font(font_name)
    
    validation = validate_text_box(
        text=text,
        font_size=font_size,
        color=resolved_color,
        position=position,
        size=size,
        allow_offslide=allow_offslide
    )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        version_before = agent.get_presentation_version()
        
        add_result = agent.add_text_box(
            slide_index=slide_index,
            text=text,
            position=position,
            size=size,
            font_name=resolved_font,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=resolved_color,
            alignment=alignment
        )
        
        agent.save()
        
        version_after = agent.get_presentation_version()
        slide_info = agent.get_slide_info(slide_index)
    
    result = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": add_result.get("shape_index") if isinstance(add_result, dict) else add_result,
        "text": text,
        "text_length": len(text),
        "position": add_result.get("position", position) if isinstance(add_result, dict) else position,
        "size": add_result.get("size", size) if isinstance(add_result, dict) else size,
        "formatting": {
            "font_name": resolved_font,
            "font_size": font_size,
            "bold": bold,
            "italic": italic,
            "color": resolved_color,
            "alignment": alignment,
            "vertical_alignment": vertical_alignment,
            "word_wrap": word_wrap
        },
        "slide_shape_count": slide_info.get("shape_count", 0),
        "validation": validation["validation_results"],
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Add text box to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
POSITION FORMATS:

  Percentage (recommended):
    {"left": "20%", "top": "30%"}
    
  Absolute inches:
    {"left": 2.0, "top": 3.0}
    
  Anchor-based:
    {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
    Anchors: top_left, top_center, top_right,
             center_left, center, center_right,
             bottom_left, bottom_center, bottom_right
    
  Grid (12-column):
    {"grid_row": 2, "grid_col": 3, "grid_size": 12}

COLOR PRESETS:

  black (#000000)      white (#FFFFFF)      primary (#0070C0)
  secondary (#595959)  accent (#ED7D31)     success (#70AD47)
  warning (#FFC000)    danger (#C00000)     dark_gray (#333333)
  light_gray (#808080) muted (#808080)

FONT PRESETS:

  default (Calibri)    heading (Calibri Light)   body (Calibri)
  code (Consolas)      serif (Georgia)           sans (Arial)

EXAMPLES:

  # Simple text box
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 0 \\
    --text "Revenue: \\$1.5M" \\
    --position '{"left":"20%","top":"30%"}' \\
    --size '{"width":"60%","height":"10%"}' \\
    --font-size 24 --bold --json

  # Centered headline
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 1 \\
    --text "Key Results" \\
    --position '{"anchor":"center","offset_y":-2}' \\
    --size '{"width":"80%","height":"15%"}' \\
    --font-size 48 --bold --color primary --alignment center --json

  # Copyright notice (bottom-right)
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 0 \\
    --text "© 2024 Company Inc." \\
    --position '{"anchor":"bottom_right","offset_x":-0.5,"offset_y":-0.3}' \\
    --size '{"width":"20%","height":"5%"}' \\
    --font-size 10 --color muted --json

ACCESSIBILITY GUIDELINES:

  Font Size:
    - Minimum: 10pt (absolute minimum)
    - Recommended: 14pt+ for projected presentations
    - Large text: 18pt+ (relaxed contrast requirements)

  Color Contrast (WCAG 2.1 AA):
    - Normal text (<18pt): 4.5:1 minimum
    - Large text (>=18pt): 3.0:1 minimum
    - Best colors: black, dark_gray, primary
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path (.pptx)'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--text',
        required=True,
        help='Text content (use \\n for line breaks)'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict as JSON'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict as JSON (defaults to 40%% x 20%% if omitted)'
    )
    
    parser.add_argument(
        '--font-name',
        default='Calibri',
        help='Font name or preset (default, heading, body, code, serif, sans)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18, recommended: >=14)'
    )
    
    parser.add_argument(
        '--bold',
        action='store_true',
        help='Make text bold'
    )
    
    parser.add_argument(
        '--italic',
        action='store_true',
        help='Make text italic'
    )
    
    parser.add_argument(
        '--color',
        help='Text color: hex (#0070C0) or preset (primary, danger, etc.)'
    )
    
    parser.add_argument(
        '--alignment',
        choices=['left', 'center', 'right', 'justify'],
        default='left',
        help='Horizontal text alignment (default: left)'
    )
    
    parser.add_argument(
        '--vertical-alignment',
        choices=['top', 'middle', 'bottom'],
        default='top',
        help='Vertical text alignment (default: top)'
    )
    
    parser.add_argument(
        '--no-word-wrap',
        action='store_true',
        help='Disable word wrap'
    )
    
    parser.add_argument(
        '--allow-offslide',
        action='store_true',
        help='Allow positioning outside slide bounds'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        size = args.size if args.size else {}
        position = args.position
        
        if "width" in position and "width" not in size:
            size["width"] = position.get("width")
        if "height" in position and "height" not in size:
            size["height"] = position.get("height")
        
        if "width" not in size:
            size["width"] = "40%"
        if "height" not in size:
            size["height"] = "20%"
        
        result = add_text_box(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            position=position,
            size=size,
            font_name=args.font_name,
            font_size=args.font_size,
            bold=args.bold,
            italic=args.italic,
            color=args.color,
            alignment=args.alignment,
            vertical_alignment=args.vertical_alignment,
            word_wrap=not args.no_word_wrap,
            allow_offslide=args.allow_offslide
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON: {e}",
            "error_type": "JSONDecodeError",
            "suggestion": "Use single quotes around JSON: '{\"left\":\"20%\",\"top\":\"30%\"}'"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check file format is .pptx and parameters are valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check the presentation file is valid and not corrupted."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_capability_probe.py
```py
#!/usr/bin/env python3
"""
PowerPoint Capability Probe Tool v3.1.0
Detect and report presentation template capabilities, layouts, and theme properties.

This tool provides comprehensive introspection of PowerPoint presentations to detect:
- Available layouts and their placeholders (with accurate runtime positions)
- Slide dimensions and aspect ratios
- Theme colors and fonts (using proper font scheme API)
- Template capabilities (footer support, slide numbers, dates)
- Multiple master slide support

Critical for AI agents and automation workflows to understand template capabilities
before generating content.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    # Basic probe (essential info)
    uv run tools/ppt_capability_probe.py --file template.pptx --json
    
    # Deep probe (accurate positions via transient instantiation)
    uv run tools/ppt_capability_probe.py --file template.pptx --deep --json
    
    # Human-friendly summary
    uv run tools/ppt_capability_probe.py --file template.pptx --summary

Exit Codes:
    0: Success
    1: Error occurred

Design Principles:
    - Read-only operation (atomic, no file mutation)
    - JSON-first output with consistent contract
    - Accurate data via transient slide instantiation
    - Graceful degradation for missing features
    - Performance-optimized with timeout protection
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that `jq` or other parsers only see valid JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import hashlib
import uuid
import time
import logging
import importlib.metadata
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
from datetime import datetime

# Configure logging to null handler
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"
SCHEMA_VERSION = "capability_probe.v3.1.0"

# ============================================================================
# IMPORTS WITH ERROR HANDLING
# ============================================================================

try:
    from pptx import Presentation
    from pptx.enum.shapes import PP_PLACEHOLDER
except ImportError as e:
    error_output = {
        "status": "error",
        "error": f"python-pptx not installed: {e}",
        "error_type": "ImportError",
        "suggestion": "Install python-pptx: pip install python-pptx"
    }
    sys.stdout.write(json.dumps(error_output, indent=2) + "\n")
    sys.exit(1)

try:
    from core.powerpoint_agent_core import PowerPointAgentError
except ImportError:
    class PowerPointAgentError(Exception):
        """Fallback exception class if core not available."""
        pass

try:
    from core.strict_validator import validate_against_schema
    STRICT_VALIDATION_AVAILABLE = True
except ImportError:
    STRICT_VALIDATION_AVAILABLE = False
    def validate_against_schema(data, schema_path):
        pass


# ============================================================================
# LIBRARY VERSION DETECTION
# ============================================================================

def get_library_versions() -> Dict[str, str]:
    """
    Detect versions of key libraries.
    
    Returns:
        Dict mapping library name to version string
    """
    versions = {}
    
    try:
        versions["python-pptx"] = importlib.metadata.version("python-pptx")
    except importlib.metadata.PackageNotFoundError:
        versions["python-pptx"] = "not_installed"
    except Exception:
        versions["python-pptx"] = "unknown"
    
    try:
        versions["Pillow"] = importlib.metadata.version("Pillow")
    except importlib.metadata.PackageNotFoundError:
        versions["Pillow"] = "not_installed"
    except Exception:
        versions["Pillow"] = "unknown"
    
    return versions


# ============================================================================
# PLACEHOLDER TYPE MAPPING
# ============================================================================

def build_placeholder_type_map() -> Dict[int, str]:
    """
    Build mapping from PP_PLACEHOLDER enum values to human-readable names.
    
    Uses actual python-pptx enum values, not guessed numbers.
    
    Returns:
        Dict mapping type code to name
    """
    type_map = {}
    
    for name in dir(PP_PLACEHOLDER):
        if name.isupper() and not name.startswith('_'):
            try:
                member = getattr(PP_PLACEHOLDER, name)
                if isinstance(member, int):
                    code = member
                elif hasattr(member, 'value'):
                    code = member.value
                else:
                    continue
                if code is not None:
                    type_map[int(code)] = name
            except (AttributeError, TypeError, ValueError):
                pass
    
    return type_map


PLACEHOLDER_TYPE_MAP = build_placeholder_type_map()


def get_placeholder_type_name(ph_type_code: int) -> str:
    """
    Get human-readable name for placeholder type code.
    
    Args:
        ph_type_code: Numeric type code from placeholder
        
    Returns:
        Type name or UNKNOWN_X if not recognized
    """
    return PLACEHOLDER_TYPE_MAP.get(ph_type_code, f"UNKNOWN_{ph_type_code}")


# ============================================================================
# FILE UTILITIES
# ============================================================================

def calculate_file_checksum(filepath: Path) -> str:
    """
    Calculate MD5 checksum of file to verify no mutation.
    
    Args:
        filepath: Path to file
        
    Returns:
        Hex digest of file contents
    """
    md5 = hashlib.md5()
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            md5.update(chunk)
    return md5.hexdigest()


# ============================================================================
# COLOR UTILITIES
# ============================================================================

def rgb_to_hex(rgb_color) -> str:
    """
    Convert RGBColor to hex string.
    
    Args:
        rgb_color: RGBColor object from python-pptx
        
    Returns:
        Hex color string like "#0070C0"
    """
    try:
        return f"#{rgb_color.r:02X}{rgb_color.g:02X}{rgb_color.b:02X}"
    except (AttributeError, TypeError):
        return "#000000"


# ============================================================================
# DIMENSION DETECTION
# ============================================================================

def detect_slide_dimensions(prs) -> Dict[str, Any]:
    """
    Detect slide dimensions and calculate aspect ratio.
    
    Args:
        prs: Presentation object
        
    Returns:
        Dict with width, height, aspect ratio, DPI estimate
    """
    width_inches = prs.slide_width.inches
    height_inches = prs.slide_height.inches
    
    width_emu = int(prs.slide_width)
    height_emu = int(prs.slide_height)
    
    dpi_estimate = 96
    width_pixels = int(width_inches * dpi_estimate)
    height_pixels = int(height_inches * dpi_estimate)
    
    ratio = width_inches / height_inches if height_inches > 0 else 1.0
    if abs(ratio - 16/9) < 0.01:
        aspect_ratio = "16:9"
    elif abs(ratio - 4/3) < 0.01:
        aspect_ratio = "4:3"
    else:
        from fractions import Fraction
        frac = Fraction(width_pixels, height_pixels).limit_denominator(20)
        aspect_ratio = f"{frac.numerator}:{frac.denominator}"
    
    return {
        "width_inches": round(width_inches, 2),
        "height_inches": round(height_inches, 2),
        "width_emu": width_emu,
        "height_emu": height_emu,
        "width_pixels": width_pixels,
        "height_pixels": height_pixels,
        "aspect_ratio": aspect_ratio,
        "aspect_ratio_float": round(ratio, 4),
        "dpi_estimate": dpi_estimate
    }


# ============================================================================
# PLACEHOLDER ANALYSIS
# ============================================================================

def analyze_placeholder(shape, slide_width: float, slide_height: float, instantiated: bool = False) -> Dict[str, Any]:
    """
    Analyze a single placeholder and return comprehensive info.
    
    Args:
        shape: Placeholder shape to analyze
        slide_width: Slide width in inches
        slide_height: Slide height in inches
        instantiated: Whether this is from an instantiated slide (accurate) or template
        
    Returns:
        Dict with type, position, size information
    """
    ph_format = shape.placeholder_format
    ph_type = ph_format.type if hasattr(ph_format.type, '__int__') else int(ph_format.type)
    ph_type_name = get_placeholder_type_name(ph_type)
    
    try:
        left_emu = shape.left if hasattr(shape, 'left') else 0
        top_emu = shape.top if hasattr(shape, 'top') else 0
        width_emu = shape.width if hasattr(shape, 'width') else 0
        height_emu = shape.height if hasattr(shape, 'height') else 0
        
        left_inches = left_emu / 914400
        top_inches = top_emu / 914400
        width_inches = width_emu / 914400
        height_inches = height_emu / 914400
        
        left_percent = (left_inches / slide_width * 100) if slide_width > 0 else 0
        top_percent = (top_inches / slide_height * 100) if slide_height > 0 else 0
        width_percent = (width_inches / slide_width * 100) if slide_width > 0 else 0
        height_percent = (height_inches / slide_height * 100) if slide_height > 0 else 0
        
        return {
            "type": ph_type_name,
            "type_code": ph_type,
            "idx": ph_format.idx,
            "name": shape.name,
            "position_inches": {
                "left": round(left_inches, 2),
                "top": round(top_inches, 2)
            },
            "position_percent": {
                "left": f"{left_percent:.1f}%",
                "top": f"{top_percent:.1f}%"
            },
            "position_emu": {
                "left": left_emu,
                "top": top_emu
            },
            "size_inches": {
                "width": round(width_inches, 2),
                "height": round(height_inches, 2)
            },
            "size_percent": {
                "width": f"{width_percent:.1f}%",
                "height": f"{height_percent:.1f}%"
            },
            "size_emu": {
                "width": width_emu,
                "height": height_emu
            },
            "position_source": "instantiated" if instantiated else "template"
        }
    except Exception as e:
        return {
            "type": ph_type_name,
            "type_code": ph_type,
            "idx": ph_format.idx,
            "error": str(e),
            "position_source": "error"
        }


# ============================================================================
# TRANSIENT SLIDE PATTERN
# ============================================================================

def _add_transient_slide(prs, layout):
    """
    Helper to safely add and remove a transient slide for deep analysis.
    Yields the slide object, then ensures cleanup in finally block.
    
    This is the key pattern for getting accurate placeholder positions:
    template positions are theoretical until instantiated.
    
    Args:
        prs: Presentation object
        layout: Layout to instantiate
        
    Yields:
        The instantiated slide for analysis
        
    Note:
        NEVER call save() while a transient slide exists.
        The slide is automatically cleaned up, and since no save occurs,
        the file remains unmodified (atomic read guarantee).
    """
    slide = None
    added_index = -1
    try:
        slide = prs.slides.add_slide(layout)
        added_index = len(prs.slides) - 1
        yield slide
    finally:
        if added_index != -1 and added_index < len(prs.slides):
            try:
                rId = prs.slides._sldIdLst[added_index].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[added_index]
            except Exception:
                # Suppress cleanup errors to avoid masking analysis errors
                # File is not saved, so transient slide disappears anyway
                pass


# ============================================================================
# LAYOUT DETECTION
# ============================================================================

def detect_layouts_with_instantiation(
    prs, 
    slide_width: float, 
    slide_height: float, 
    deep: bool, 
    warnings: List[str], 
    timeout_start: Optional[float] = None, 
    timeout_seconds: Optional[int] = None,
    max_layouts: Optional[int] = None
) -> List[Dict[str, Any]]:
    """
    Detect all layouts, optionally instantiating them for accurate positions.
    
    In deep mode, creates transient slides in-memory to get runtime positions,
    then discards them without saving (maintains atomic read guarantee).
    
    Args:
        prs: Presentation object
        slide_width: Slide width in inches
        slide_height: Slide height in inches
        deep: If True, instantiate layouts for accurate positions
        warnings: List to append warnings to
        timeout_start: Start time for timeout check
        timeout_seconds: Max seconds allowed
        max_layouts: Maximum number of layouts to analyze
        
    Returns:
        List of layout information dicts
    """
    layouts = []
    
    # Build mapping: layout partname -> master index
    master_map = {}
    try:
        for m_idx, master in enumerate(prs.slide_masters):
            for layout in master.slide_layouts:
                try:
                    key = layout.part.partname
                except AttributeError:
                    key = id(layout)
                master_map[key] = m_idx
    except Exception:
        pass

    layouts_to_process = list(prs.slide_layouts)
    total_layouts = len(layouts_to_process)
    
    if max_layouts and len(layouts_to_process) > max_layouts:
        layouts_to_process = layouts_to_process[:max_layouts]

    for idx, layout in enumerate(layouts_to_process):
        # Timeout check at each iteration
        if timeout_start and timeout_seconds:
            elapsed = time.perf_counter() - timeout_start
            if elapsed > timeout_seconds:
                warnings.append(
                    f"Probe timeout at layout {idx} ({elapsed:.1f}s > {timeout_seconds}s) - "
                    "returning partial results"
                )
                break

        # Get original index in case we sliced
        try:
            original_idx = list(prs.slide_layouts).index(layout)
        except ValueError:
            original_idx = idx

        # Determine master index
        try:
            key = layout.part.partname
        except AttributeError:
            key = id(layout)
            
        layout_info = {
            "index": idx,
            "original_index": original_idx, 
            "name": layout.name,
            "placeholder_count": len(layout.placeholders),
            "master_index": master_map.get(key)
        }
        
        if deep:
            try:
                instantiation_success = False
                for temp_slide in _add_transient_slide(prs, layout):
                    instantiation_success = True
                    
                    # Map instantiated placeholders by idx for lookup
                    instantiated_map = {}
                    for shape in temp_slide.placeholders:
                        try:
                            instantiated_map[shape.placeholder_format.idx] = shape
                        except (AttributeError, TypeError):
                            pass
                    
                    placeholders = []
                    for layout_ph in layout.placeholders:
                        try:
                            ph_idx = layout_ph.placeholder_format.idx
                            if ph_idx in instantiated_map:
                                ph_info = analyze_placeholder(
                                    instantiated_map[ph_idx], 
                                    slide_width, 
                                    slide_height, 
                                    instantiated=True
                                )
                            else:
                                ph_info = analyze_placeholder(
                                    layout_ph, 
                                    slide_width, 
                                    slide_height, 
                                    instantiated=False
                                )
                            placeholders.append(ph_info)
                        except Exception:
                            pass
                    
                    layout_info["placeholders"] = placeholders
                    layout_info["instantiation_complete"] = len(placeholders) == len(layout.placeholders)
                    layout_info["placeholder_expected"] = len(layout.placeholders)
                    layout_info["placeholder_instantiated"] = len(placeholders)

                if not instantiation_success:
                    raise Exception("Transient slide creation failed")
                
            except Exception as e:
                warnings.append(f"Could not instantiate layout '{layout.name}': {str(e)}")
                
                placeholders = []
                for shape in layout.placeholders:
                    try:
                        ph_info = analyze_placeholder(shape, slide_width, slide_height, instantiated=False)
                        placeholders.append(ph_info)
                    except Exception:
                        pass
                
                layout_info["placeholders"] = placeholders
                layout_info["instantiation_complete"] = False
                layout_info["placeholder_expected"] = len(layout.placeholders)
                layout_info["placeholder_instantiated"] = len(placeholders)
                layout_info["_warning"] = "Using template positions (instantiation failed)"
        
        # Build placeholder type summary
        placeholder_map = {}
        placeholder_types = []
        for shape in layout.placeholders:
            try:
                ph_type = shape.placeholder_format.type
                if hasattr(ph_type, '__int__'):
                    ph_type = int(ph_type)
                ph_type_name = get_placeholder_type_name(ph_type)
                
                placeholder_map[ph_type_name] = placeholder_map.get(ph_type_name, 0) + 1
                
                if ph_type_name not in placeholder_types:
                    placeholder_types.append(ph_type_name)
            except Exception:
                pass
        
        layout_info["placeholder_types"] = placeholder_types
        layout_info["placeholder_map"] = placeholder_map
        
        layouts.append(layout_info)
    
    return layouts


# ============================================================================
# THEME EXTRACTION
# ============================================================================

def extract_theme_colors(master_or_prs, warnings: List[str]) -> Dict[str, str]:
    """
    Extract theme colors from presentation or master using proper color scheme API.
    
    Args:
        master_or_prs: Presentation or SlideMaster object
        warnings: List to append warnings to
        
    Returns:
        Dict mapping color names to hex codes or scheme references
    """
    colors = {}
    
    try:
        if hasattr(master_or_prs, 'slide_masters'):
            slide_master = master_or_prs.slide_masters[0]
        else:
            slide_master = master_or_prs

        theme = getattr(slide_master, 'theme', None)
        if not theme:
            warnings.append("Theme object unavailable")
            return {}
            
        color_scheme = getattr(theme, 'theme_color_scheme', None)
        if not color_scheme:
            warnings.append("Theme color scheme unavailable")
            return {}
        
        color_attrs = [
            'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
            'background1', 'background2', 'text1', 'text2', 'hyperlink', 'followed_hyperlink'
        ]
        
        non_rgb_found = False
        for color_name in color_attrs:
            try:
                color = getattr(color_scheme, color_name, None)
                if color:
                    if hasattr(color, 'r'):
                        colors[color_name] = rgb_to_hex(color)
                    else:
                        colors[color_name] = f"schemeColor:{color_name}"
                        non_rgb_found = True
            except Exception:
                pass
        
        if not colors:
            warnings.append("Theme color scheme unavailable or empty")
        elif non_rgb_found:
            warnings.append("Theme colors include non-RGB scheme references")
            
    except Exception as e:
        warnings.append(f"Theme color extraction failed: {str(e)}")
    
    return colors


def _font_name(font_obj) -> Optional[str]:
    """Helper to safely get typeface from font object."""
    if font_obj is None:
        return None
    return getattr(font_obj, 'typeface', str(font_obj))


def extract_theme_fonts(master_or_prs, warnings: List[str]) -> Dict[str, str]:
    """
    Extract theme fonts from presentation or master using proper font scheme API.
    
    Args:
        master_or_prs: Presentation or SlideMaster object
        warnings: List to append warnings to
        
    Returns:
        Dict with heading and body font names
    """
    fonts = {}
    fallback_used = False
    
    try:
        if hasattr(master_or_prs, 'slide_masters'):
            slide_master = master_or_prs.slide_masters[0]
        else:
            slide_master = master_or_prs

        theme = getattr(slide_master, 'theme', None)
        
        if theme:
            font_scheme = getattr(theme, 'font_scheme', None)
            if font_scheme:
                major = getattr(font_scheme, 'major_font', None)
                minor = getattr(font_scheme, 'minor_font', None)
                
                if major:
                    latin = getattr(major, 'latin', None)
                    ea = getattr(major, 'east_asian', None)
                    cs = getattr(major, 'complex_script', None)
                    
                    heading_font = _font_name(latin) or _font_name(ea) or _font_name(cs)
                    if heading_font:
                        fonts['heading'] = heading_font
                    
                    if ea and _font_name(ea):
                        fonts['heading_east_asian'] = _font_name(ea)
                    if cs and _font_name(cs):
                        fonts['heading_complex'] = _font_name(cs)
                
                if minor:
                    latin = getattr(minor, 'latin', None)
                    ea = getattr(minor, 'east_asian', None)
                    cs = getattr(minor, 'complex_script', None)
                    
                    body_font = _font_name(latin) or _font_name(ea) or _font_name(cs)
                    if body_font:
                        fonts['body'] = body_font
                    
                    if ea and _font_name(ea):
                        fonts['body_east_asian'] = _font_name(ea)
                    if cs and _font_name(cs):
                        fonts['body_complex'] = _font_name(cs)

        if not fonts:
            for shape in slide_master.shapes:
                if hasattr(shape, 'text_frame') and shape.text_frame.paragraphs:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.font.name and 'heading' not in fonts:
                            fonts['heading'] = paragraph.font.name
                            break
                    if 'heading' in fonts:
                        break
        
        if not fonts:
            fallback_used = True
            fonts = {"heading": "Calibri", "body": "Calibri"}
            
    except Exception as e:
        fallback_used = True
        fonts = {"heading": "Calibri", "body": "Calibri"}
        warnings.append(f"Theme font extraction failed: {str(e)}")
    
    if fallback_used and hasattr(master_or_prs, 'slide_masters'):
        warnings.append("Theme fonts unavailable - using Calibri defaults")
    
    return fonts


# ============================================================================
# CAPABILITY ANALYSIS
# ============================================================================

def analyze_capabilities(layouts: List[Dict[str, Any]], prs) -> Dict[str, Any]:
    """
    Analyze template capabilities based on detected layouts.
    
    Args:
        layouts: List of layout information dicts
        prs: Presentation object
        
    Returns:
        Dict with capability flags, layout mappings, and recommendations
    """
    has_footer = False
    has_slide_number = False
    has_date = False
    layouts_with_footer = []
    layouts_with_slide_number = []
    layouts_with_date = []
    
    footer_type_code = None
    slide_number_type_code = None
    date_type_code = None
    
    for type_code, type_name in PLACEHOLDER_TYPE_MAP.items():
        if type_name == 'FOOTER':
            footer_type_code = type_code
        elif type_name == 'SLIDE_NUMBER':
            slide_number_type_code = type_code
        elif type_name == 'DATE':
            date_type_code = type_code
    
    per_master_stats = {}
    
    for layout in layouts:
        layout_ref = {
            "index": layout['index'],
            "original_index": layout.get('original_index', layout['index']),
            "name": layout['name'],
            "master_index": layout.get('master_index')
        }
        m_idx = layout.get('master_index')
        
        if m_idx is not None:
            if m_idx not in per_master_stats:
                per_master_stats[m_idx] = {
                    "master_index": m_idx,
                    "layout_count": 0,
                    "has_footer_layouts": 0,
                    "has_slide_number_layouts": 0,
                    "has_date_layouts": 0
                }
            per_master_stats[m_idx]["layout_count"] += 1
        
        layout_has_footer = False
        layout_has_slide_number = False
        layout_has_date = False

        if 'placeholders' in layout:
            for ph in layout['placeholders']:
                if footer_type_code and ph.get('type_code') == footer_type_code:
                    has_footer = True
                    layout_has_footer = True
                    if layout_ref not in layouts_with_footer:
                        layouts_with_footer.append(layout_ref)
                
                if slide_number_type_code and ph.get('type_code') == slide_number_type_code:
                    has_slide_number = True
                    layout_has_slide_number = True
                    if layout_ref not in layouts_with_slide_number:
                        layouts_with_slide_number.append(layout_ref)
                
                if date_type_code and ph.get('type_code') == date_type_code:
                    has_date = True
                    layout_has_date = True
                    if layout_ref not in layouts_with_date:
                        layouts_with_date.append(layout_ref)
                        
        elif 'placeholder_types' in layout:
            if 'FOOTER' in layout['placeholder_types']:
                has_footer = True
                layout_has_footer = True
                layouts_with_footer.append(layout_ref)
            
            if 'SLIDE_NUMBER' in layout['placeholder_types']:
                has_slide_number = True
                layout_has_slide_number = True
                layouts_with_slide_number.append(layout_ref)
            
            if 'DATE' in layout['placeholder_types']:
                has_date = True
                layout_has_date = True
                layouts_with_date.append(layout_ref)
        
        if m_idx is not None:
            if layout_has_footer:
                per_master_stats[m_idx]["has_footer_layouts"] += 1
            if layout_has_slide_number:
                per_master_stats[m_idx]["has_slide_number_layouts"] += 1
            if layout_has_date:
                per_master_stats[m_idx]["has_date_layouts"] += 1
    
    recommendations = []
    
    if not has_footer:
        recommendations.append(
            "No footer placeholders found - ppt_set_footer.py will use text box fallback strategy"
        )
    else:
        layout_names = [l['name'] for l in layouts_with_footer]
        recommendations.append(
            f"Footer placeholders available on {len(layouts_with_footer)} layout(s): {', '.join(layout_names)}"
        )
    
    if not has_slide_number:
        recommendations.append(
            "No slide number placeholders - recommend manual text box for slide numbers"
        )
    else:
        layout_names = [l['name'] for l in layouts_with_slide_number]
        recommendations.append(
            f"Slide number placeholders available on {len(layouts_with_slide_number)} layout(s): {', '.join(layout_names)}"
        )
    
    if not has_date:
        recommendations.append(
            "No date placeholders - dates must be added manually if needed"
        )
    else:
        layout_names = [l['name'] for l in layouts_with_date]
        recommendations.append(
            f"Date placeholders available on {len(layouts_with_date)} layout(s): {', '.join(layout_names)}"
        )
    
    return {
        "has_footer_placeholders": has_footer,
        "has_slide_number_placeholders": has_slide_number,
        "has_date_placeholders": has_date,
        "layouts_with_footer": layouts_with_footer,
        "layouts_with_slide_number": layouts_with_slide_number,
        "layouts_with_date": layouts_with_date,
        "total_layouts": len(layouts),
        "total_master_slides": len(prs.slide_masters),
        "per_master": list(per_master_stats.values()),
        "footer_support_mode": "placeholder" if has_footer else "fallback_textbox",
        "slide_number_strategy": "placeholder" if has_slide_number else "textbox",
        "recommendations": recommendations
    }


# ============================================================================
# OUTPUT VALIDATION
# ============================================================================

def validate_output(result: Dict[str, Any]) -> Tuple[bool, List[str]]:
    """
    Validate probe result has all required fields.
    
    Args:
        result: Probe result dict
        
    Returns:
        Tuple of (is_valid, list of missing fields)
    """
    required_fields = [
        "status",
        "metadata",
        "metadata.file",
        "metadata.probed_at",
        "metadata.tool_version",
        "metadata.operation_id",
        "metadata.duration_ms",
        "slide_dimensions",
        "layouts",
        "theme",
        "capabilities",
        "warnings"
    ]
    
    missing = []
    
    for field_path in required_fields:
        parts = field_path.split('.')
        current = result
        
        for part in parts:
            if not isinstance(current, dict) or part not in current:
                missing.append(field_path)
                break
            current = current[part]
    
    return (len(missing) == 0, missing)


# ============================================================================
# MAIN PROBE FUNCTION
# ============================================================================

def probe_presentation(
    filepath: Path,
    deep: bool = False,
    verify_atomic: bool = True,
    max_layouts: Optional[int] = None,
    timeout_seconds: Optional[int] = None
) -> Dict[str, Any]:
    """
    Probe presentation and return comprehensive capability report.
    
    Args:
        filepath: Path to PowerPoint file
        deep: If True, perform deep analysis with transient slide instantiation
        verify_atomic: If True, verify no file mutation occurred
        max_layouts: Maximum layouts to analyze (None = all)
        timeout_seconds: Maximum seconds for analysis (None = no limit)
        
    Returns:
        Dict with complete capability report
        
    Raises:
        FileNotFoundError: If file doesn't exist
        PermissionError: If file is locked
        PowerPointAgentError: If atomic verification fails
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not filepath.is_file():
        raise ValueError(f"Path is not a file: {filepath}")
    
    try:
        with open(filepath, 'rb') as f:
            f.read(1)
    except PermissionError:
        raise PermissionError(f"File is locked or permission denied: {filepath}")
    
    start_time = time.perf_counter()
    operation_id = str(uuid.uuid4())
    warnings = []
    info = []
    
    checksum_before = None
    if verify_atomic:
        checksum_before = calculate_file_checksum(filepath)
    
    prs = Presentation(str(filepath))
    
    dimensions = detect_slide_dimensions(prs)
    slide_width = dimensions['width_inches']
    slide_height = dimensions['height_inches']
    
    all_layouts = list(prs.slide_layouts)
    if max_layouts and len(all_layouts) > max_layouts:
        info.append(f"Limited analysis to first {max_layouts} of {len(all_layouts)} layouts")
    
    layouts = detect_layouts_with_instantiation(
        prs, 
        slide_width, 
        slide_height, 
        deep, 
        warnings, 
        timeout_start=start_time, 
        timeout_seconds=timeout_seconds,
        max_layouts=max_layouts
    )
    
    analysis_complete = True
    if timeout_seconds and (time.perf_counter() - start_time) > timeout_seconds:
        analysis_complete = False
    
    theme_colors = extract_theme_colors(prs, warnings)
    theme_fonts = extract_theme_fonts(prs, warnings)
    
    theme_per_master = []
    try:
        for m_idx, master in enumerate(prs.slide_masters):
            m_warnings = []
            m_colors = extract_theme_colors(master, m_warnings)
            m_fonts = extract_theme_fonts(master, m_warnings)
            theme_per_master.append({
                "master_index": m_idx,
                "colors": m_colors,
                "fonts": m_fonts
            })
    except Exception:
        pass
    
    capabilities = analyze_capabilities(layouts, prs)
    capabilities["analysis_complete"] = analysis_complete
    
    duration_ms = int((time.perf_counter() - start_time) * 1000)
    
    if verify_atomic:
        checksum_after = calculate_file_checksum(filepath)
        if checksum_before != checksum_after:
            raise PowerPointAgentError(
                "File was modified during probe operation! "
                "This should never happen (atomic read violation). "
                f"Checksum before: {checksum_before}, after: {checksum_after}"
            )
    
    masters_info = []
    try:
        for m_idx, m in enumerate(prs.slide_masters):
            masters_info.append({
                "master_index": m_idx,
                "layout_count": len(m.slide_layouts),
                "name": getattr(m, 'name', f"Master {m_idx}")
            })
    except Exception:
        pass
    
    result = {
        "status": "success",
        "metadata": {
            "file": str(filepath.resolve()),
            "probed_at": datetime.now().isoformat(),
            "tool_version": __version__,
            "schema_version": SCHEMA_VERSION,
            "operation_id": operation_id,
            "deep_analysis": deep,
            "analysis_mode": "deep" if deep else "essential",
            "atomic_verified": verify_atomic,
            "duration_ms": duration_ms,
            "timeout_seconds": timeout_seconds,
            "layout_count_total": len(all_layouts),
            "layout_count_analyzed": len(layouts),
            "warnings_count": len(warnings),
            "masters": masters_info,
            "library_versions": get_library_versions(),
            "checksum": checksum_before if verify_atomic else "verification_skipped"
        },
        "slide_dimensions": dimensions,
        "layouts": layouts,
        "theme": {
            "colors": theme_colors,
            "fonts": theme_fonts,
            "per_master": theme_per_master
        },
        "capabilities": capabilities,
        "warnings": warnings,
        "info": info
    }
    
    is_valid, missing_fields = validate_output(result)
    if not is_valid:
        warnings.append(f"Output validation found missing fields: {', '.join(missing_fields)}")
    
    if STRICT_VALIDATION_AVAILABLE:
        try:
            schema_path = Path(__file__).parent.parent / "schemas" / "capability_probe.v3.1.0.schema.json"
            if schema_path.exists():
                validate_against_schema(result, str(schema_path))
        except FileNotFoundError:
            info.append("Schema file not found - strict validation skipped")
        except Exception as e:
            warnings.append(f"Strict schema validation warning: {str(e)}")
    
    return result


# ============================================================================
# HUMAN-READABLE SUMMARY
# ============================================================================

def format_summary(probe_result: Dict[str, Any]) -> str:
    """
    Format probe result as human-readable summary.
    
    Args:
        probe_result: Result from probe_presentation()
        
    Returns:
        Formatted string summary
    """
    lines = []
    
    lines.append("═══════════════════════════════════════════════════════════════")
    lines.append(f"PowerPoint Capability Probe Report v{__version__}")
    lines.append("═══════════════════════════════════════════════════════════════")
    lines.append("")
    
    meta = probe_result['metadata']
    lines.append(f"File: {meta['file']}")
    lines.append(f"Probed: {meta['probed_at']}")
    lines.append(f"Operation ID: {meta['operation_id']}")
    lines.append(f"Analysis Mode: {'Deep (instantiated positions)' if meta['deep_analysis'] else 'Essential (template positions)'}")
    lines.append(f"Duration: {meta['duration_ms']}ms")
    lines.append(f"Atomic Verified: {'✓' if meta['atomic_verified'] else '✗'}")
    lines.append("")
    
    if meta.get('library_versions'):
        lines.append("Library Versions:")
        for lib, ver in meta['library_versions'].items():
            lines.append(f"  {lib}: {ver}")
        lines.append("")
    
    dims = probe_result['slide_dimensions']
    lines.append("Slide Dimensions:")
    lines.append(f"  Size: {dims['width_inches']}\" × {dims['height_inches']}\" ({dims['width_pixels']}×{dims['height_pixels']}px)")
    lines.append(f"  Aspect Ratio: {dims['aspect_ratio']}")
    lines.append(f"  DPI Estimate: {dims['dpi_estimate']}")
    lines.append("")
    
    caps = probe_result['capabilities']
    lines.append("Template Capabilities:")
    lines.append(f"  ✓ Total Layouts: {caps['total_layouts']}")
    lines.append(f"  ✓ Master Slides: {caps['total_master_slides']}")
    lines.append(f"  {'✓' if caps['has_footer_placeholders'] else '✗'} Footer Placeholders: {len(caps['layouts_with_footer'])} layout(s)")
    lines.append(f"  {'✓' if caps['has_slide_number_placeholders'] else '✗'} Slide Number Placeholders: {len(caps['layouts_with_slide_number'])} layout(s)")
    lines.append(f"  {'✓' if caps['has_date_placeholders'] else '✗'} Date Placeholders: {len(caps['layouts_with_date'])} layout(s)")
    lines.append("")

    if 'per_master' in caps and caps['per_master']:
        lines.append("Master Slides:")
        for m in caps['per_master']:
            lines.append(f"  Master {m['master_index']}: {m['layout_count']} layouts")
            footer = 'Yes' if m['has_footer_layouts'] else 'No'
            slide_num = 'Yes' if m['has_slide_number_layouts'] else 'No'
            date = 'Yes' if m['has_date_layouts'] else 'No'
            lines.append(f"    Footer: {footer} | Slide #: {slide_num} | Date: {date}")
        lines.append("")
    
    lines.append("Available Layouts:")
    for layout in probe_result['layouts']:
        ph_count = layout['placeholder_count']
        display_idx = layout.get('original_index', layout['index'])
        lines.append(f"  [{display_idx}] {layout['name']} ({ph_count} placeholder{'s' if ph_count != 1 else ''})")
        
        if 'placeholder_types' in layout and layout['placeholder_types']:
            types_str = ', '.join(layout['placeholder_types'])
            lines.append(f"      Types: {types_str}")
    lines.append("")
    
    theme = probe_result['theme']
    if theme.get('fonts'):
        lines.append("Theme Fonts:")
        for key, value in theme['fonts'].items():
            if not key.startswith('_'):
                lines.append(f"  {key.replace('_', ' ').title()}: {value}")
        lines.append("")
    
    if theme.get('colors'):
        color_count = len([k for k in theme['colors'].keys() if not k.startswith('_')])
        lines.append(f"Theme Colors: {color_count} defined")
        lines.append("")
    
    if caps.get('recommendations'):
        lines.append("Recommendations:")
        for rec in caps['recommendations']:
            lines.append(f"  • {rec}")
        lines.append("")
    
    if probe_result.get('warnings'):
        lines.append("⚠️  Warnings:")
        for warning in probe_result['warnings']:
            lines.append(f"  • {warning}")
        lines.append("")
    
    if probe_result.get('info'):
        lines.append("ℹ️  Information:")
        for info_msg in probe_result['info']:
            lines.append(f"  • {info_msg}")
        lines.append("")
    
    lines.append("═══════════════════════════════════════════════════════════════")
    
    return "\n".join(lines)


# ============================================================================
# CLI ENTRY POINT
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description=f"Probe PowerPoint presentation capabilities (v{__version__})",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic probe (essential info, fast)
  uv run tools/ppt_capability_probe.py --file template.pptx --json
  
  # Deep probe (accurate positions via transient instantiation)
  uv run tools/ppt_capability_probe.py --file template.pptx --deep --json
  
  # Human-friendly summary
  uv run tools/ppt_capability_probe.py --file template.pptx --summary
  
  # Skip atomic verification for speed
  uv run tools/ppt_capability_probe.py --file template.pptx --no-verify-atomic --json
  
  # Large template with layout limit
  uv run tools/ppt_capability_probe.py --file big_template.pptx --max-layouts 20 --json

Version: """ + __version__
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to probe'
    )
    
    parser.add_argument(
        '--deep',
        action='store_true',
        help='Perform deep analysis with transient slide instantiation for accurate positions (slower)'
    )
    
    parser.add_argument(
        '--summary',
        action='store_true',
        help='Output human-friendly summary instead of JSON'
    )
    
    parser.add_argument(
        '--verify-atomic',
        action='store_true',
        default=True,
        dest='verify_atomic',
        help='Verify no file mutation occurred (default: true)'
    )
    
    parser.add_argument(
        '--no-verify-atomic',
        action='store_false',
        dest='verify_atomic',
        help='Skip atomic verification (faster, less safe)'
    )
    
    parser.add_argument(
        '--max-layouts',
        type=int,
        help='Maximum layouts to analyze (for large templates)'
    )

    parser.add_argument(
        '--timeout',
        type=int,
        default=30,
        help='Timeout in seconds for analysis (default: 30)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        dest='output_json',
        help='Output JSON format (default if --summary not used)'
    )
    
    args = parser.parse_args()
    
    if not args.summary and not args.output_json:
        args.output_json = True
        
    if args.summary and args.output_json:
        error_output = {
            "status": "error",
            "error": "Cannot use both --summary and --json",
            "error_type": "ArgumentError"
        }
        sys.stdout.write(json.dumps(error_output, indent=2) + "\n")
        sys.exit(1)
    
    try:
        result = probe_presentation(
            filepath=args.file,
            deep=args.deep,
            verify_atomic=args.verify_atomic,
            max_layouts=args.max_layouts,
            timeout_seconds=args.timeout
        )
        
        if args.summary:
            print(format_summary(result))
        else:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check file path and permissions",
            "metadata": {
                "file": str(args.file) if args.file else None,
                "tool_version": __version__,
                "operation_id": str(uuid.uuid4()),
                "probed_at": datetime.now().isoformat()
            }
        }
        
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_check_accessibility.py
```py
#!/usr/bin/env python3
"""
PowerPoint Check Accessibility Tool v3.1.0
Run WCAG 2.1 accessibility checks on presentation.

This tool performs comprehensive accessibility validation including:
- Alt text presence for images
- Color contrast ratios
- Reading order verification
- Font size compliance

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json

Exit Codes:
    0: Success (check completed, see 'passed' field for result)
    1: Error occurred (file not found, crash)

Design Principles:
    - Read-only operation (acquire_lock=False)
    - JSON-first output with consistent contract
    - Strict output hygiene (stderr suppressed)
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately.
# This prevents libraries (pptx, warnings) from printing non-JSON text
# which corrupts pipelines that capture 2>&1.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import logging
import time
import uuid
from pathlib import Path
from typing import Dict, Any
from datetime import datetime

# Configure logging to null handler
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"

# ============================================================================
# IMPORTS WITH ERROR HANDLING
# ============================================================================

try:
    from core.powerpoint_agent_core import (
        PowerPointAgent,
        PowerPointAgentError
    )
    CORE_AVAILABLE = True
except ImportError as e:
    CORE_AVAILABLE = False
    IMPORT_ERROR = str(e)


# ============================================================================
# MAIN LOGIC
# ============================================================================

def check_accessibility(filepath: Path) -> Dict[str, Any]:
    """
    Run accessibility checks on a PowerPoint presentation.
    
    Args:
        filepath: Path to PowerPoint file
        
    Returns:
        Dict with accessibility check results including:
        - status: "success"
        - passed: bool indicating if all checks passed
        - issues: dict of issue categories
        - summary: counts by category
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ImportError: If core module not available
        PowerPointAgentError: If presentation cannot be opened
    """
    if not CORE_AVAILABLE:
        raise ImportError(f"Core module not available: {IMPORT_ERROR}")
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    start_time = time.perf_counter()
    operation_id = str(uuid.uuid4())
    
    with PowerPointAgent(filepath) as agent:
        # acquire_lock=False because validation is read-only
        agent.open(filepath, acquire_lock=False)
        
        # Capture presentation version for audit trail
        try:
            presentation_version = agent.get_presentation_version()
        except AttributeError:
            presentation_version = "unknown"
        
        # Get presentation info for context
        try:
            pres_info = agent.get_presentation_info()
            slide_count = pres_info.get("slide_count", 0)
        except Exception:
            slide_count = 0
        
        # Run accessibility check
        result = agent.check_accessibility()
    
    duration_ms = int((time.perf_counter() - start_time) * 1000)
    
    # Extract issues for summary
    issues = result.get("issues", {})
    missing_alt_text = issues.get("missing_alt_text", [])
    low_contrast = issues.get("low_contrast", [])
    reading_order = issues.get("reading_order_issues", [])
    small_fonts = issues.get("small_fonts", [])
    
    total_issues = (
        len(missing_alt_text) + 
        len(low_contrast) + 
        len(reading_order) + 
        len(small_fonts)
    )
    
    # Determine pass/fail
    passed = result.get("passed", total_issues == 0)
    
    return {
        "status": "success",
        "passed": passed,
        "file": str(filepath.resolve()),
        "tool_version": __version__,
        "presentation_version": presentation_version,
        "validated_at": datetime.utcnow().isoformat() + "Z",
        "operation_id": operation_id,
        "duration_ms": duration_ms,
        "summary": {
            "slide_count": slide_count,
            "total_issues": total_issues,
            "missing_alt_text_count": len(missing_alt_text),
            "low_contrast_count": len(low_contrast),
            "reading_order_issues_count": len(reading_order),
            "small_fonts_count": len(small_fonts)
        },
        "issues": {
            "missing_alt_text": missing_alt_text,
            "low_contrast": low_contrast,
            "reading_order_issues": reading_order,
            "small_fonts": small_fonts
        },
        "wcag_level": "AA",
        "recommendations": _generate_recommendations(issues)
    }


def _generate_recommendations(issues: Dict[str, Any]) -> list:
    """
    Generate actionable recommendations based on issues found.
    
    Args:
        issues: Dict of issue categories
        
    Returns:
        List of recommendation dicts
    """
    recommendations = []
    
    missing_alt = issues.get("missing_alt_text", [])
    if missing_alt:
        recommendations.append({
            "priority": "high",
            "category": "accessibility",
            "action": f"Add alt text to {len(missing_alt)} image(s)",
            "fix_command": "ppt_set_image_properties.py --alt-text"
        })
    
    low_contrast = issues.get("low_contrast", [])
    if low_contrast:
        recommendations.append({
            "priority": "high",
            "category": "accessibility",
            "action": f"Fix contrast on {len(low_contrast)} element(s)",
            "fix_command": "ppt_format_text.py --color"
        })
    
    small_fonts = issues.get("small_fonts", [])
    if small_fonts:
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "action": f"Increase font size on {len(small_fonts)} element(s)",
            "fix_command": "ppt_format_text.py --font-size"
        })
    
    reading_order = issues.get("reading_order_issues", [])
    if reading_order:
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "action": f"Review reading order on {len(reading_order)} slide(s)",
            "fix_command": "Manual review required"
        })
    
    return recommendations


# ============================================================================
# CLI ENTRY POINT
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description=f"Check PowerPoint accessibility (WCAG 2.1) - v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Check accessibility
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json
    
Output includes:
    - passed: Boolean indicating if accessibility checks passed
    - summary: Counts of issues by category
    - issues: Detailed list of accessibility violations
    - recommendations: Suggested fixes with commands

Version: """ + __version__
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = check_accessibility(filepath=args.file)
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ImportError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ImportError",
            "suggestion": "Ensure core module is properly installed"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_clone_presentation.py
```py
#!/usr/bin/env python3
"""
PowerPoint Clone Presentation Tool v3.1.0
Create an exact copy of a presentation for safe editing

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

⚠️ GOVERNANCE FOUNDATION - Clone-Before-Edit Principle

This tool implements the foundational safety principle: NEVER modify source
files directly. Always create a working copy first using this tool.

Usage:
    uv run tools/ppt_clone_presentation.py --source original.pptx --output work_copy.pptx --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Safety Workflow:
    1. Clone: ppt_clone_presentation.py --source original.pptx --output work.pptx
    2. Edit: Use other tools on work.pptx
    3. Validate: ppt_validate_presentation.py --file work.pptx
    4. Deliver: Rename/move work.pptx when approved
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError
)

__version__ = "3.1.0"


def clone_presentation(
    source: Path, 
    output: Path
) -> Dict[str, Any]:
    """
    Create an exact copy of a PowerPoint presentation.
    
    This is the foundational tool for the Clone-Before-Edit governance
    principle. Always use this before modifying any presentation to:
    
    1. Protect source files from accidental modification
    2. Enable rollback to original if needed
    3. Create audit-safe work copies
    4. Allow parallel editing without conflicts
    
    Args:
        source: Path to the source presentation to clone
        output: Path where the clone will be saved
        
    Returns:
        Dict containing:
            - status: "success"
            - source: Absolute path to source file
            - output: Absolute path to cloned file
            - source_size_bytes: Size of source file
            - output_size_bytes: Size of cloned file
            - slide_count: Number of slides in presentation
            - presentation_version: State hash of the cloned presentation
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If source file doesn't exist
        PermissionError: If output location is not writable
        
    Example:
        >>> result = clone_presentation(
        ...     source=Path("template.pptx"),
        ...     output=Path("work/project.pptx")
        ... )
        >>> print(result["presentation_version"])
        'a1b2c3d4e5f6g7h8'
    """
    # Validate source exists
    if not source.exists():
        raise FileNotFoundError(f"Source file not found: {source}")
    
    # Validate source is a PowerPoint file
    if source.suffix.lower() not in {'.pptx', '.pptm', '.potx'}:
        raise ValueError(
            f"Source must be a PowerPoint file (.pptx, .pptm, .potx), got: {source.suffix}"
        )
    
    # Ensure output has correct extension
    if not output.suffix.lower() == '.pptx':
        output = output.with_suffix('.pptx')
    
    # Create output directory if needed
    output.parent.mkdir(parents=True, exist_ok=True)
    
    # Get source file size
    source_size = source.stat().st_size
    
    # Open source (read-only, no lock) and save to output
    with PowerPointAgent(source) as agent:
        agent.open(source, acquire_lock=False)  # Read-only, don't lock source
        
        # Get presentation info before saving
        info = agent.get_presentation_info()
        
        # Save to new location (creates the clone)
        agent.save(output)
        
        # Get the cloned presentation's version
        presentation_version = info.get("presentation_version")
        slide_count = info.get("slide_count", 0)
    
    # Get output file size (should match source)
    output_size = output.stat().st_size
    
    return {
        "status": "success",
        "source": str(source.resolve()),
        "output": str(output.resolve()),
        "source_size_bytes": source_size,
        "output_size_bytes": output_size,
        "slide_count": slide_count,
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Clone PowerPoint presentation (⚠️ GOVERNANCE FOUNDATION)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️ GOVERNANCE FOUNDATION: Clone-Before-Edit Principle

This tool implements the first and most important safety rule:
NEVER modify source files directly. Always create a working copy.

Examples:
  # Basic clone
  uv run tools/ppt_clone_presentation.py \\
    --source original.pptx \\
    --output work_copy.pptx \\
    --json
  
  # Clone to work directory
  uv run tools/ppt_clone_presentation.py \\
    --source templates/corporate.pptx \\
    --output work/q4_report.pptx \\
    --json
  
  # Clone for parallel editing
  uv run tools/ppt_clone_presentation.py \\
    --source shared/presentation.pptx \\
    --output my_edits/presentation_v2.pptx \\
    --json

Safety Workflow:
  1. CLONE the source file:
     uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx
  
  2. PROBE the clone:
     uv run tools/ppt_capability_probe.py --file work.pptx --deep --json
  
  3. EDIT the clone (not the original!):
     uv run tools/ppt_add_slide.py --file work.pptx --layout "Title Slide" --json
  
  4. VALIDATE before delivery:
     uv run tools/ppt_validate_presentation.py --file work.pptx --json
     uv run tools/ppt_check_accessibility.py --file work.pptx --json
  
  5. DELIVER when approved:
     mv work.pptx final_presentation.pptx

Why Clone-Before-Edit?
  - Protects original files from accidental modification
  - Enables rollback if edits go wrong
  - Creates audit trail (original preserved)
  - Allows concurrent work without conflicts
  - Required by governance framework

Version Tracking:
  The presentation_version in the output is a state hash that can be used
  to track changes. After editing the clone, the version will change.
  Compare versions to detect modifications.

Output Format:
  {
    "status": "success",
    "source": "/path/to/original.pptx",
    "output": "/path/to/work_copy.pptx",
    "source_size_bytes": 1234567,
    "output_size_bytes": 1234567,
    "slide_count": 15,
    "presentation_version": "a1b2c3d4e5f6g7h8",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--source', 
        required=True, 
        type=Path, 
        help='Source PowerPoint file to clone'
    )
    parser.add_argument(
        '--output', 
        required=True, 
        type=Path, 
        help='Destination path for the cloned file'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = clone_presentation(source=args.source, output=args.output)
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the source file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Source must be a PowerPoint file (.pptx, .pptm, .potx)"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PermissionError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PermissionError",
            "suggestion": "Check write permissions for the output directory"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_create_from_structure.py
```py
#!/usr/bin/env python3
"""
PowerPoint Create From Structure Tool v3.1.0
Create a complete presentation from a JSON structure definition.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_create_from_structure.py --structure deck.json --output presentation.pptx --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
)

__version__ = "3.1.0"


def validate_structure(structure: Dict[str, Any]) -> None:
    """
    Validate the JSON structure schema.
    
    Args:
        structure: Dictionary containing presentation structure
        
    Raises:
        ValueError: If structure is invalid
    """
    if "slides" not in structure:
        raise ValueError("Structure must contain 'slides' array")
    
    if not isinstance(structure["slides"], list):
        raise ValueError("'slides' must be an array")
    
    if len(structure["slides"]) == 0:
        raise ValueError("Must have at least one slide")
    
    if len(structure["slides"]) > 100:
        raise ValueError("Maximum 100 slides supported (performance limit)")


def create_from_structure(
    structure: Dict[str, Any],
    output: Path
) -> Dict[str, Any]:
    """
    Create a PowerPoint presentation from a JSON structure definition.
    
    Args:
        structure: Dictionary defining presentation structure with slides and content
        output: Output path for the created presentation (.pptx)
        
    Returns:
        Dict containing:
            - status: 'success' or 'success_with_errors'
            - file: Absolute path to created file
            - presentation_version: Version hash of created presentation
            - slides_created: Number of slides created
            - content_added: Dict with counts per content type
            - errors: List of error messages encountered
            - error_count: Total number of errors
            - file_size_bytes: Size of created file
            - tool_version: Tool version string
            
    Raises:
        ValueError: If structure is invalid
        PowerPointAgentError: If presentation creation fails
    """
    validate_structure(structure)
    
    stats = {
        "slides_created": 0,
        "text_boxes_added": 0,
        "images_inserted": 0,
        "charts_added": 0,
        "tables_added": 0,
        "shapes_added": 0,
        "errors": []
    }
    
    with PowerPointAgent() as agent:
        template = structure.get("template")
        if template and Path(template).exists():
            agent.create_new(template=Path(template))
        else:
            agent.create_new()
        
        for slide_idx, slide_def in enumerate(structure["slides"]):
            try:
                layout = slide_def.get("layout", "Title and Content")
                agent.add_slide(layout_name=layout)
                stats["slides_created"] += 1
                
                if "title" in slide_def:
                    agent.set_title(
                        slide_index=slide_idx,
                        title=slide_def["title"],
                        subtitle=slide_def.get("subtitle")
                    )
                
                for item in slide_def.get("content", []):
                    try:
                        item_type = item.get("type")
                        
                        if item_type == "text_box":
                            agent.add_text_box(
                                slide_index=slide_idx,
                                text=item["text"],
                                position=item["position"],
                                size=item["size"],
                                font_name=item.get("font_name", "Calibri"),
                                font_size=item.get("font_size", 18),
                                bold=item.get("bold", False),
                                italic=item.get("italic", False),
                                color=item.get("color"),
                                alignment=item.get("alignment", "left")
                            )
                            stats["text_boxes_added"] += 1
                        
                        elif item_type == "image":
                            image_path = Path(item["path"])
                            if image_path.exists():
                                agent.insert_image(
                                    slide_index=slide_idx,
                                    image_path=image_path,
                                    position=item["position"],
                                    size=item.get("size"),
                                    compress=item.get("compress", False)
                                )
                                stats["images_inserted"] += 1
                            else:
                                stats["errors"].append(f"Image not found: {item['path']}")
                        
                        elif item_type == "chart":
                            agent.add_chart(
                                slide_index=slide_idx,
                                chart_type=item["chart_type"],
                                data=item["data"],
                                position=item["position"],
                                size=item["size"],
                                chart_title=item.get("title")
                            )
                            stats["charts_added"] += 1
                        
                        elif item_type == "table":
                            agent.add_table(
                                slide_index=slide_idx,
                                rows=item["rows"],
                                cols=item["cols"],
                                position=item["position"],
                                size=item["size"],
                                data=item.get("data")
                            )
                            stats["tables_added"] += 1
                        
                        elif item_type == "shape":
                            agent.add_shape(
                                slide_index=slide_idx,
                                shape_type=item["shape_type"],
                                position=item["position"],
                                size=item["size"],
                                fill_color=item.get("fill_color"),
                                line_color=item.get("line_color"),
                                line_width=item.get("line_width", 1.0)
                            )
                            stats["shapes_added"] += 1
                        
                        elif item_type == "bullet_list":
                            agent.add_bullet_list(
                                slide_index=slide_idx,
                                items=item["items"],
                                position=item["position"],
                                size=item["size"],
                                bullet_style=item.get("bullet_style", "bullet"),
                                font_size=item.get("font_size", 18)
                            )
                            stats["text_boxes_added"] += 1
                        
                        else:
                            stats["errors"].append(f"Unknown content type: {item_type}")
                    
                    except Exception as e:
                        stats["errors"].append(f"Error adding {item.get('type', 'unknown')}: {str(e)}")
            
            except Exception as e:
                stats["errors"].append(f"Error processing slide {slide_idx}: {str(e)}")
        
        agent.save(output)
        
        presentation_version = agent.get_presentation_version()
    
    file_size = output.stat().st_size if output.exists() else 0
    
    return {
        "status": "success" if len(stats["errors"]) == 0 else "success_with_errors",
        "file": str(output.resolve()),
        "presentation_version": presentation_version,
        "slides_created": stats["slides_created"],
        "content_added": {
            "text_boxes": stats["text_boxes_added"],
            "images": stats["images_inserted"],
            "charts": stats["charts_added"],
            "tables": stats["tables_added"],
            "shapes": stats["shapes_added"]
        },
        "errors": stats["errors"],
        "error_count": len(stats["errors"]),
        "file_size_bytes": file_size,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Create PowerPoint presentation from JSON structure definition",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
JSON Structure Schema:
{
  "template": "optional_template.pptx",
  "slides": [
    {
      "layout": "Title Slide",
      "title": "Presentation Title",
      "subtitle": "Subtitle",
      "content": [
        {
          "type": "text_box",
          "text": "Content here",
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "10%"},
          "font_size": 18,
          "color": "#000000"
        },
        {
          "type": "image",
          "path": "image.png",
          "position": {"left": "20%", "top": "30%"},
          "size": {"width": "60%", "height": "auto"}
        },
        {
          "type": "chart",
          "chart_type": "column",
          "data": {
            "categories": ["Q1", "Q2", "Q3"],
            "series": [{"name": "Revenue", "values": [100, 120, 140]}]
          },
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "60%"}
        },
        {
          "type": "table",
          "rows": 3,
          "cols": 3,
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "50%"},
          "data": [["A", "B", "C"], ["1", "2", "3"]]
        },
        {
          "type": "shape",
          "shape_type": "rectangle",
          "position": {"left": "10%", "top": "10%"},
          "size": {"width": "30%", "height": "15%"},
          "fill_color": "#0070C0"
        },
        {
          "type": "bullet_list",
          "items": ["Item 1", "Item 2", "Item 3"],
          "position": {"left": "10%", "top": "25%"},
          "size": {"width": "80%", "height": "60%"}
        }
      ]
    }
  ]
}

Examples:
    # Create simple presentation
    cat > structure.json << 'EOF'
{
  "slides": [
    {
      "layout": "Title Slide",
      "title": "My Presentation",
      "subtitle": "Created from Structure"
    },
    {
      "layout": "Title and Content",
      "title": "Agenda",
      "content": [
        {
          "type": "bullet_list",
          "items": ["Introduction", "Main Content", "Conclusion"],
          "position": {"left": "10%", "top": "25%"},
          "size": {"width": "80%", "height": "60%"}
        }
      ]
    }
  ]
}
EOF
    
    uv run tools/ppt_create_from_structure.py \\
        --structure structure.json \\
        --output presentation.pptx \\
        --json

    # Create presentation with charts
    cat > complex.json << 'EOF'
{
  "slides": [
    {
      "layout": "Title Slide",
      "title": "Q4 Report"
    },
    {
      "layout": "Title and Content",
      "title": "Revenue Growth",
      "content": [
        {
          "type": "chart",
          "chart_type": "column",
          "data": {
            "categories": ["Q1", "Q2", "Q3", "Q4"],
            "series": [
              {"name": "2023", "values": [100, 110, 120, 130]},
              {"name": "2024", "values": [120, 135, 145, 160]}
            ]
          },
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "65%"},
          "title": "Year over Year Comparison"
        }
      ]
    }
  ]
}
EOF
    
    uv run tools/ppt_create_from_structure.py \\
        --structure complex.json \\
        --output q4_report.pptx \\
        --json

Content Types:
    text_box    - Text container with formatting options
    image       - Image from file path
    chart       - Data visualization (column, bar, line, pie)
    table       - Grid of cells
    shape       - Geometric shape (rectangle, arrow, etc.)
    bullet_list - Bulleted or numbered list

Use Cases:
    - Automated report generation
    - Template-based presentations from data
    - Batch presentation creation
    - AI-generated presentations
    - Programmatic deck building
        """
    )
    
    parser.add_argument(
        '--structure',
        required=True,
        type=Path,
        help='Path to JSON structure file'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output path for created presentation (.pptx)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output as JSON (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        if not args.structure.exists():
            raise FileNotFoundError(f"Structure file not found: {args.structure}")
        
        with open(args.structure, 'r', encoding='utf-8') as f:
            structure = json.load(f)
        
        output_path = args.output
        if output_path.suffix.lower() != '.pptx':
            output_path = output_path.with_suffix('.pptx')
        
        result = create_from_structure(
            structure=structure,
            output=output_path
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the structure file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON in structure file: {str(e)}",
            "error_type": "JSONDecodeError",
            "suggestion": "Validate JSON syntax using a JSON linter."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check structure file matches the required schema."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Check template file exists and is valid if specified."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_create_from_template.py
```py
#!/usr/bin/env python3
"""
PowerPoint Create From Template Tool v3.1.1
Create new presentation from existing .pptx template.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_create_from_template.py --template corporate_template.pptx --output new_presentation.pptx --slides 10 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
    - Added get_available_layouts() fallback for compatibility
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError,
    LayoutNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# MAIN LOGIC
# ============================================================================

def create_from_template(
    template: Path,
    output: Path,
    slides: int = 1,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """
    Create a new PowerPoint presentation from an existing template.
    
    This tool copies the template (including its theme, master slides, and
    any existing content) and optionally adds additional slides using the
    specified layout.
    
    Args:
        template: Path to the source template .pptx file
        output: Path where the new presentation will be saved
        slides: Total number of slides desired in the output (default: 1)
        layout: Layout name for additional slides (default: "Title and Content")
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to created file
            - template_used: Path to source template
            - total_slides: Final slide count
            - slides_requested: Number of slides requested
            - template_slides: Number of slides in original template
            - slides_added: Number of slides added
            - layout_used: Layout name used for added slides
            - available_layouts: List of all available layouts
            - file_size_bytes: Size of created file
            - slide_dimensions: Width, height, and aspect ratio
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If template file does not exist
        ValueError: If template is not .pptx or slide count invalid
        LayoutNotFoundError: If specified layout not found (falls back)
        
    Example:
        >>> result = create_from_template(
        ...     template=Path("templates/corporate.pptx"),
        ...     output=Path("q4_report.pptx"),
        ...     slides=15,
        ...     layout="Title and Content"
        ... )
        >>> print(result["total_slides"])
        15
    """
    if not template.exists():
        raise FileNotFoundError(f"Template file not found: {template}")
    
    if not template.suffix.lower() == '.pptx':
        raise ValueError(f"Template must be .pptx file, got: {template.suffix}")
    
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
    with PowerPointAgent() as agent:
        agent.create_new(template=template)
        
        try:
            available_layouts = agent.get_available_layouts()
        except AttributeError:
            info = agent.get_presentation_info()
            available_layouts = info.get("layouts", [])
        
        resolved_layout = layout
        if layout not in available_layouts:
            layout_lower = layout.lower()
            matched = False
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    resolved_layout = avail
                    matched = True
                    break
            
            if not matched:
                resolved_layout = available_layouts[0] if available_layouts else "Title Slide"
        
        current_slides = agent.get_slide_count()
        
        slides_to_add = max(0, slides - current_slides)
        
        slide_indices: List[int] = list(range(current_slides))
        
        for i in range(slides_to_add):
            result = agent.add_slide(layout_name=resolved_layout)
            if isinstance(result, dict):
                idx = result.get("slide_index", result.get("index", len(slide_indices)))
            else:
                idx = result
            slide_indices.append(idx)
        
        agent.save(output)
        
        info = agent.get_presentation_info()
        presentation_version = info.get("presentation_version", None)
    
    file_size = output.stat().st_size if output.exists() else 0
    
    return {
        "status": "success",
        "file": str(output.resolve()),
        "template_used": str(template.resolve()),
        "total_slides": info["slide_count"],
        "slides_requested": slides,
        "template_slides": current_slides,
        "slides_added": slides_to_add,
        "layout_used": resolved_layout,
        "available_layouts": info.get("layouts", available_layouts),
        "file_size_bytes": file_size,
        "slide_dimensions": {
            "width_inches": info.get("slide_width_inches", 13.333),
            "height_inches": info.get("slide_height_inches", 7.5),
            "aspect_ratio": info.get("aspect_ratio", "16:9")
        },
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Create PowerPoint presentation from template",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create from corporate template with 15 slides
  uv run tools/ppt_create_from_template.py \\
    --template templates/corporate.pptx \\
    --output q4_report.pptx \\
    --slides 15 \\
    --json
  
  # Create presentation using specific layout
  uv run tools/ppt_create_from_template.py \\
    --template templates/minimal.pptx \\
    --output demo.pptx \\
    --slides 5 \\
    --layout "Section Header" \\
    --json
  
  # Quick presentation from template (uses template's existing slides)
  uv run tools/ppt_create_from_template.py \\
    --template templates/branded.pptx \\
    --output quick_deck.pptx \\
    --json

Use Cases:
  - Corporate presentations with consistent branding
  - Team presentations with shared theme
  - Pre-formatted layouts (fonts, colors, logos)
  - Department-specific templates
  - Client-specific branded decks

Template Benefits:
  - Consistent branding across organization
  - Pre-configured master slides
  - Corporate colors and fonts
  - Logo placements
  - Standard layouts
  - Accessibility features built-in

Creating Templates:
  1. Design in PowerPoint with desired theme
  2. Configure master slides
  3. Set up color scheme
  4. Define standard layouts
  5. Save as .pptx template
  6. Use with this tool

Best Practices:
  - Maintain template library for different purposes
  - Version control templates
  - Document template usage guidelines
  - Test templates before distribution
  - Include variety of layouts in template

Output Format:
  {
    "status": "success",
    "file": "/path/to/output.pptx",
    "template_used": "/path/to/template.pptx",
    "total_slides": 15,
    "template_slides": 1,
    "slides_added": 14,
    "layout_used": "Title and Content",
    "presentation_version": "a1b2c3d4...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--template',
        required=True,
        type=Path,
        help='Path to template .pptx file'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output presentation path'
    )
    
    parser.add_argument(
        '--slides',
        type=int,
        default=1,
        help='Total number of slides desired (default: 1)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout for additional slides (default: "Title and Content")'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        output_path = args.output
        if not output_path.suffix.lower() == '.pptx':
            output_path = output_path.with_suffix('.pptx')
        
        result = create_from_template(
            template=args.template.resolve(),
            output=output_path.resolve(),
            slides=args.slides,
            layout=args.layout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            sys.stdout.write(f"Created presentation from template: {result['file']}\n")
            sys.stdout.write(f"  Template: {result['template_used']}\n")
            sys.stdout.write(f"  Total slides: {result['total_slides']}\n")
            sys.stdout.write(f"  Template had: {result['template_slides']} slides\n")
            sys.stdout.write(f"  Added: {result['slides_added']} slides\n")
            sys.stdout.write(f"  Layout: {result['layout_used']}\n")
            sys.stdout.flush()
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the template file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check that template is .pptx and slide count is 1-100",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except LayoutNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "LayoutNotFoundError",
            "suggestion": "Use ppt_capability_probe.py to discover available layouts",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check template file integrity and available layouts",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_create_new.py
```py
#!/usr/bin/env python3
"""
PowerPoint Create New Tool v3.1.1
Create a new PowerPoint presentation with specified slides.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_create_new.py --output presentation.pptx --slides 5 --layout "Title and Content" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Note:
    For creating presentations from existing templates with branding,
    consider using ppt_create_from_template.py instead.

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
    - Added get_available_layouts() fallback for compatibility
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError,
    LayoutNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# MAIN LOGIC
# ============================================================================

def create_new_presentation(
    output: Path,
    slides: int,
    template: Optional[Path] = None,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """
    Create a new PowerPoint presentation with specified number of slides.
    
    Creates a blank presentation (or from optional template) and populates
    it with the requested number of slides. The first slide uses "Title Slide"
    layout if available, subsequent slides use the specified layout.
    
    Args:
        output: Path where the new presentation will be saved
        slides: Number of slides to create (1-100)
        template: Optional path to template .pptx file (default: None for blank)
        layout: Layout name for slides after the first (default: "Title and Content")
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to created file
            - slides_created: Number of slides created
            - slide_indices: List of slide indices
            - file_size_bytes: Size of created file
            - slide_dimensions: Width, height, and aspect ratio
            - available_layouts: List of all available layouts
            - layout_used: Layout name used for non-title slides
            - template_used: Path to template if used, else None
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Raises:
        ValueError: If slide count is invalid (not 1-100)
        FileNotFoundError: If template specified but not found
        
    Example:
        >>> result = create_new_presentation(
        ...     output=Path("pitch_deck.pptx"),
        ...     slides=10,
        ...     layout="Title and Content"
        ... )
        >>> print(result["slides_created"])
        10
    """
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
    if template is not None:
        if not template.exists():
            raise FileNotFoundError(f"Template file not found: {template}")
        if not template.suffix.lower() == '.pptx':
            raise ValueError(f"Template must be .pptx file, got: {template.suffix}")
    
    with PowerPointAgent() as agent:
        agent.create_new(template=template)
        
        try:
            available_layouts = agent.get_available_layouts()
        except AttributeError:
            info = agent.get_presentation_info()
            available_layouts = info.get("layouts", [])
        
        resolved_layout = layout
        if layout not in available_layouts:
            layout_lower = layout.lower()
            matched = False
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    resolved_layout = avail
                    matched = True
                    break
            
            if not matched:
                resolved_layout = available_layouts[0] if available_layouts else "Title Slide"
        
        slide_indices: List[int] = []
        
        for i in range(slides):
            if i == 0 and "Title Slide" in available_layouts:
                slide_layout = "Title Slide"
            else:
                slide_layout = resolved_layout
            
            result = agent.add_slide(layout_name=slide_layout)
            if isinstance(result, dict):
                idx = result.get("slide_index", result.get("index", i))
            else:
                idx = result
            slide_indices.append(idx)
        
        agent.save(output)
        
        info = agent.get_presentation_info()
        presentation_version = info.get("presentation_version", None)
    
    file_size = output.stat().st_size if output.exists() else 0
    
    return {
        "status": "success",
        "file": str(output.resolve()),
        "slides_created": slides,
        "slide_indices": slide_indices,
        "file_size_bytes": file_size,
        "slide_dimensions": {
            "width_inches": info.get("slide_width_inches", 13.333),
            "height_inches": info.get("slide_height_inches", 7.5),
            "aspect_ratio": info.get("aspect_ratio", "16:9")
        },
        "available_layouts": info.get("layouts", available_layouts),
        "layout_used": resolved_layout,
        "template_used": str(template.resolve()) if template else None,
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Create new PowerPoint presentation with specified slides",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create presentation with 5 blank slides
  uv run tools/ppt_create_new.py --output presentation.pptx --slides 5 --json
  
  # Create with specific layout
  uv run tools/ppt_create_new.py --output pitch_deck.pptx --slides 10 --layout "Title and Content" --json
  
  # Create from template (for simple cases; use ppt_create_from_template.py for advanced)
  uv run tools/ppt_create_new.py --output new_deck.pptx --slides 3 --template corporate_template.pptx --json
  
  # Create single title slide
  uv run tools/ppt_create_new.py --output title.pptx --slides 1 --layout "Title Slide" --json

Available Layouts (typical):
  - Title Slide
  - Title and Content
  - Section Header
  - Two Content
  - Comparison
  - Title Only
  - Blank
  - Content with Caption
  - Picture with Caption

First Slide Behavior:
  The first slide automatically uses "Title Slide" layout if available,
  regardless of the --layout parameter. Subsequent slides use --layout.

For Template-Based Creation:
  If you need to preserve template content or work with branded templates,
  use ppt_create_from_template.py instead. This tool is optimized for
  creating presentations from scratch.

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slides_created": 5,
    "slide_indices": [0, 1, 2, 3, 4],
    "file_size_bytes": 28432,
    "slide_dimensions": {
      "width_inches": 13.333,
      "height_inches": 7.5,
      "aspect_ratio": "16:9"
    },
    "available_layouts": ["Title Slide", "Title and Content", ...],
    "layout_used": "Title and Content",
    "presentation_version": "a1b2c3d4...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output PowerPoint file path (.pptx)'
    )
    
    parser.add_argument(
        '--slides',
        type=int,
        default=1,
        help='Number of slides to create (default: 1)'
    )
    
    parser.add_argument(
        '--template',
        type=Path,
        default=None,
        help='Optional template file to use (.pptx)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout to use for slides after the first (default: "Title and Content")'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        output_path = args.output
        if not output_path.suffix.lower() == '.pptx':
            output_path = output_path.with_suffix('.pptx')
        
        result = create_new_presentation(
            output=output_path.resolve(),
            slides=args.slides,
            template=args.template.resolve() if args.template else None,
            layout=args.layout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            sys.stdout.write(f"Created presentation: {result['file']}\n")
            sys.stdout.write(f"  Slides: {result['slides_created']}\n")
            sys.stdout.write(f"  Layout: {result['layout_used']}\n")
            sys.stdout.write(f"  Dimensions: {result['slide_dimensions']['aspect_ratio']}\n")
            if args.template:
                sys.stdout.write(f"  Template: {result['template_used']}\n")
            sys.stdout.flush()
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the template file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check slide count (1-100) and template file extension (.pptx)",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except LayoutNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "LayoutNotFoundError",
            "suggestion": "Use ppt_capability_probe.py to discover available layouts",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check file permissions and template compatibility",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_crop_image.py
```py
#!/usr/bin/env python3
"""
PowerPoint Crop Image Tool v3.1.0
Crop an existing image on a slide by trimming edges

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_crop_image.py --file deck.pptx --slide 0 --shape 1 --left 0.1 --right 0.1 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Notes:
    Crop values are percentages of the original image size (0.0 to 1.0).
    For example, --left 0.1 trims 10% from the left edge.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

# Import MSO_SHAPE_TYPE safely
try:
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    # Fallback if pptx not directly importable
    MSO_SHAPE_TYPE = None

__version__ = "3.1.0"


# Define ShapeNotFoundError if not available in core
try:
    from core.powerpoint_agent_core import ShapeNotFoundError
except ImportError:
    class ShapeNotFoundError(PowerPointAgentError):
        """Exception raised when shape is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def crop_image(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    left: float = 0.0,
    right: float = 0.0,
    top: float = 0.0,
    bottom: float = 0.0
) -> Dict[str, Any]:
    """
    Crop an image on a slide by trimming edges.
    
    Applies crop values to an existing image shape. Crop values represent
    the percentage of the original image to trim from each edge.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the image (0-based)
        shape_index: Index of the image shape to crop (0-based)
        left: Percentage to crop from left edge (0.0-1.0, default: 0.0)
        right: Percentage to crop from right edge (0.0-1.0, default: 0.0)
        top: Percentage to crop from top edge (0.0-1.0, default: 0.0)
        bottom: Percentage to crop from bottom edge (0.0-1.0, default: 0.0)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the cropped shape
            - crop_applied: Dict with applied crop values
            - presentation_version_before: State hash before crop
            - presentation_version_after: State hash after crop
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        ShapeNotFoundError: If the shape index is out of range
        ValueError: If crop values are invalid or shape is not an image
        
    Example:
        >>> result = crop_image(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     shape_index=1,
        ...     left=0.1,
        ...     right=0.1
        ... )
        >>> print(result["crop_applied"])
        {'left': 0.1, 'right': 0.1, 'top': 0.0, 'bottom': 0.0}
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate crop values
    crop_values = [left, right, top, bottom]
    for name, value in [("left", left), ("right", right), ("top", top), ("bottom", bottom)]:
        if not (0.0 <= value <= 1.0):
            raise ValueError(
                f"Crop value '{name}' must be between 0.0 and 1.0, got: {value}"
            )
    
    # Validate total crop doesn't exceed 100%
    if left + right >= 1.0:
        raise ValueError(
            f"Combined left ({left}) and right ({right}) crop cannot exceed 1.0"
        )
    if top + bottom >= 1.0:
        raise ValueError(
            f"Combined top ({top}) and bottom ({bottom}) crop cannot exceed 1.0"
        )

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE crop
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # NOTE: Direct prs access required because python-pptx crop API
        # requires accessing the shape's crop properties directly.
        # This is a necessary workaround for features not exposed via agent methods.
        slide = agent.prs.slides[slide_index]
        
        # Validate shape index
        if not 0 <= shape_index < len(slide.shapes):
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{len(slide.shapes) - 1})",
                details={
                    "requested_index": shape_index,
                    "available_shapes": len(slide.shapes)
                }
            )
        
        shape = slide.shapes[shape_index]
        
        # Validate shape is a picture
        if MSO_SHAPE_TYPE is not None:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                raise ValueError(
                    f"Shape at index {shape_index} is not an image (type: {shape.shape_type}). "
                    "Use ppt_get_slide_info.py to identify image shapes."
                )
        else:
            # Fallback check if MSO_SHAPE_TYPE not available
            if not hasattr(shape, 'crop_left'):
                raise ValueError(
                    f"Shape at index {shape_index} does not support cropping. "
                    "Ensure it is an image shape."
                )
        
        # Apply crop values (only set if > 0 to avoid unnecessary changes)
        if left > 0:
            shape.crop_left = left
        if right > 0:
            shape.crop_right = right
        if top > 0:
            shape.crop_top = top
        if bottom > 0:
            shape.crop_bottom = bottom
        
        # Save changes
        agent.save()
        
        # Capture version AFTER crop
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "crop_applied": {
            "left": left,
            "right": right,
            "top": top,
            "bottom": bottom
        },
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Crop an image in a PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Crop 10%% from left and right edges
  uv run tools/ppt_crop_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --left 0.1 \\
    --right 0.1 \\
    --json
  
  # Crop to focus on center (trim all edges)
  uv run tools/ppt_crop_image.py \\
    --file deck.pptx \\
    --slide 2 \\
    --shape 3 \\
    --left 0.15 \\
    --right 0.15 \\
    --top 0.1 \\
    --bottom 0.1 \\
    --json
  
  # Crop top portion only
  uv run tools/ppt_crop_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 0 \\
    --top 0.2 \\
    --json

Crop Values:
  - Values are percentages of original image size (0.0 to 1.0)
  - 0.0 = no crop, 0.1 = 10%% crop, 0.5 = 50%% crop
  - Combined opposite edges (left+right or top+bottom) must be < 1.0

Finding Images:
  Use ppt_get_slide_info.py to identify image shape indices:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "shape_index": 1,
    "crop_applied": {
      "left": 0.1,
      "right": 0.1,
      "top": 0.0,
      "bottom": 0.0
    },
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    parser.add_argument(
        '--slide', 
        required=True, 
        type=int, 
        help='Slide index (0-based)'
    )
    parser.add_argument(
        '--shape', 
        required=True, 
        type=int, 
        help='Shape index of image to crop (0-based)'
    )
    parser.add_argument(
        '--left', 
        type=float, 
        default=0.0, 
        help='Crop percentage from left edge (0.0-1.0, default: 0.0)'
    )
    parser.add_argument(
        '--right', 
        type=float, 
        default=0.0, 
        help='Crop percentage from right edge (0.0-1.0, default: 0.0)'
    )
    parser.add_argument(
        '--top', 
        type=float, 
        default=0.0, 
        help='Crop percentage from top edge (0.0-1.0, default: 0.0)'
    )
    parser.add_argument(
        '--bottom', 
        type=float, 
        default=0.0, 
        help='Crop percentage from bottom edge (0.0-1.0, default: 0.0)'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = crop_image(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            left=args.left,
            right=args.right,
            top=args.top,
            bottom=args.bottom
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check crop values (0.0-1.0) and ensure shape is an image"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_delete_slide.py
```py
#!/usr/bin/env python3
"""
PowerPoint Delete Slide Tool v3.1.1
Remove a slide from the presentation.

⚠️ DESTRUCTIVE OPERATION - Requires approval token with scope 'delete:slide'

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_delete_slide.py --file presentation.pptx --index 1 --approval-token "HMAC-SHA256:..." --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)
    4: Permission error (missing or invalid approval token)

Security:
    This tool performs a destructive operation and requires a valid approval
    token with scope 'delete:slide'. Generate tokens using the approval token
    system described in the governance documentation.

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"

# ============================================================================
# EXCEPTION FALLBACK
# ============================================================================

try:
    from core.powerpoint_agent_core import ApprovalTokenError
except ImportError:
    class ApprovalTokenError(PowerPointAgentError):
        """Exception raised when approval token is missing or invalid."""
        def __init__(self, message: str, details: Optional[Dict] = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)
        
        def __str__(self):
            return self.message


# ============================================================================
# MAIN LOGIC
# ============================================================================

def delete_slide(
    filepath: Path, 
    index: int,
    approval_token: str
) -> Dict[str, Any]:
    """
    Delete a slide at the specified index.
    
    ⚠️ DESTRUCTIVE OPERATION - This permanently removes the slide.
    
    This operation requires a valid approval token with scope 'delete:slide'
    to prevent accidental data loss. Always clone the presentation first
    using ppt_clone_presentation.py before performing destructive operations.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        index: Slide index to delete (0-based)
        approval_token: HMAC-SHA256 approval token with scope 'delete:slide'
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - deleted_index: Index of the deleted slide
            - remaining_slides: Number of slides after deletion
            - presentation_version_before: State hash before deletion
            - presentation_version_after: State hash after deletion
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        ApprovalTokenError: If approval token is missing or invalid
        
    Example:
        >>> result = delete_slide(
        ...     filepath=Path("presentation.pptx"),
        ...     index=2,
        ...     approval_token="HMAC-SHA256:eyJ..."
        ... )
        >>> print(result["remaining_slides"])
        9
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not approval_token:
        raise ApprovalTokenError(
            "Approval token required for slide deletion",
            details={
                "operation": "delete_slide",
                "slide_index": index,
                "required_scope": "delete:slide",
                "file": str(filepath)
            }
        )
    
    if not approval_token.startswith("HMAC-SHA256:"):
        raise ApprovalTokenError(
            "Invalid approval token format. Expected 'HMAC-SHA256:...'",
            details={
                "operation": "delete_slide",
                "slide_index": index,
                "required_scope": "delete:slide",
                "token_prefix_received": approval_token[:20] + "..." if len(approval_token) > 20 else approval_token
            }
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        total_slides = agent.get_slide_count()
        if not 0 <= index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": index,
                    "available_slides": total_slides,
                    "valid_range": f"0 to {total_slides - 1}"
                }
            )
        
        try:
            agent.delete_slide(index, approval_token=approval_token)
        except TypeError:
            agent.delete_slide(index)
        
        agent.save()
        
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
        new_count = info_after["slide_count"]
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "deleted_index": index,
        "remaining_slides": new_count,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Delete PowerPoint slide (⚠️ DESTRUCTIVE - requires approval token)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️ DESTRUCTIVE OPERATION ⚠️

This tool permanently removes a slide from the presentation.
An approval token with scope 'delete:slide' is REQUIRED.

Examples:
  # Delete slide at index 2 (third slide)
  uv run tools/ppt_delete_slide.py \\
    --file presentation.pptx \\
    --index 2 \\
    --approval-token "HMAC-SHA256:eyJzY29wZSI6ImRlbGV0ZTpzbGlkZSIsLi4ufQ==.abc123..." \\
    --json

Safety Workflow:
  1. CLONE the presentation first:
     uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx
  
  2. VERIFY slide count and content:
     uv run tools/ppt_get_info.py --file work.pptx --json
     uv run tools/ppt_get_slide_info.py --file work.pptx --slide 2 --json
  
  3. GENERATE approval token with scope 'delete:slide'
  
  4. DELETE the slide:
     uv run tools/ppt_delete_slide.py --file work.pptx --index 2 --approval-token "..." --json

Token Generation:
  Approval tokens must be generated by a trusted service using HMAC-SHA256.
  The token must have scope 'delete:slide' and not be expired.
  See governance documentation for token generation details.

Exit Codes:
  0: Success - slide deleted
  1: Error - check error_type in JSON output
  4: Permission Error - missing or invalid approval token

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "deleted_index": 2,
    "remaining_slides": 9,
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--index', 
        required=True, 
        type=int, 
        help='Slide index to delete (0-based)'
    )
    
    parser.add_argument(
        '--approval-token',
        required=True,
        type=str,
        help='Approval token with scope "delete:slide" (REQUIRED for this destructive operation)'
    )
    
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = delete_slide(
            filepath=args.file.resolve(), 
            index=args.index,
            approval_token=args.approval_token
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
    except ApprovalTokenError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ApprovalTokenError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Generate a valid approval token with scope 'delete:slide'",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(4)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check file integrity and slide index validity",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_duplicate_slide.py
```py
#!/usr/bin/env python3
"""
PowerPoint Duplicate Slide Tool v3.1.1
Clone an existing slide within the presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_duplicate_slide.py --file presentation.pptx --index 0 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# MAIN LOGIC
# ============================================================================

def duplicate_slide(
    filepath: Path, 
    index: int
) -> Dict[str, Any]:
    """
    Duplicate a slide at the specified index.
    
    Creates a deep copy of the slide including all shapes, text runs,
    formatting, and styles. The duplicated slide is inserted immediately
    after the source slide.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        index: Index of the slide to duplicate (0-based)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - source_index: Index of the original slide
            - new_slide_index: Index of the newly created duplicate
            - total_slides: Total slide count after duplication
            - layout: Layout name of the duplicated slide
            - presentation_version_before: State hash before duplication
            - presentation_version_after: State hash after duplication
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        
    Example:
        >>> result = duplicate_slide(
        ...     filepath=Path("presentation.pptx"),
        ...     index=0
        ... )
        >>> print(result["new_slide_index"])
        1
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        total = agent.get_slide_count()
        if not 0 <= index < total:
            raise SlideNotFoundError(
                f"Slide index {index} out of range (0-{total - 1})",
                details={
                    "requested_index": index,
                    "available_slides": total,
                    "valid_range": f"0 to {total - 1}"
                }
            )
        
        result = agent.duplicate_slide(index)
        
        if isinstance(result, dict):
            new_index = result.get("slide_index", result.get("new_slide_index", index + 1))
        else:
            new_index = result
        
        agent.save()
        
        slide_info = agent.get_slide_info(new_index)
        layout_name = slide_info.get("layout", "Unknown")
        
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
        final_count = info_after["slide_count"]
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "source_index": index,
        "new_slide_index": new_index,
        "total_slides": final_count,
        "layout": layout_name,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Duplicate a PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Duplicate the first slide
  uv run tools/ppt_duplicate_slide.py --file presentation.pptx --index 0 --json
  
  # Duplicate slide at index 5
  uv run tools/ppt_duplicate_slide.py --file deck.pptx --index 5 --json

Behavior:
  - Creates a deep copy of the slide at the specified index
  - The duplicate is inserted immediately after the source slide
  - All shapes, text, formatting, and styles are preserved
  - Returns the index of the newly created slide

Use Cases:
  - Creating similar slides with slight variations
  - Building slide sequences from a template slide
  - Backing up a slide before major changes

Important Notes:
  - Shape indices on the duplicated slide start fresh
  - The new slide gets the next available index
  - Use ppt_get_slide_info.py to inspect the new slide

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "source_index": 0,
    "new_slide_index": 1,
    "total_slides": 6,
    "layout": "Title and Content",
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--index', 
        required=True, 
        type=int, 
        help='Source slide index to duplicate (0-based)'
    )
    
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = duplicate_slide(
            filepath=args.file.resolve(),
            index=args.index
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check file integrity and slide index validity",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_export_images.py
```py
#!/usr/bin/env python3
"""
PowerPoint Export Images Tool v3.1.1
Export each slide as PNG or JPG image using LibreOffice.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_export_images.py --file presentation.pptx --output-dir output/ --format png --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for image export:
    - Linux: sudo apt install libreoffice-impress
    - macOS: brew install --cask libreoffice
    - Windows: Download from https://www.libreoffice.org/

Changelog v3.1.1:
    - Added hygiene block for JSON pipeline safety
    - Added presentation_version tracking
    - Added tool_version to output
    - Added --timeout argument
    - Fixed error response format with suggestions
    - Removed stderr print statements
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that JSON parsers only see valid JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import subprocess
import shutil
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def check_libreoffice() -> bool:
    """Check if LibreOffice is installed and accessible."""
    return shutil.which('soffice') is not None or shutil.which('libreoffice') is not None


def get_libreoffice_command() -> str:
    """Get the LibreOffice command for the current system."""
    if shutil.which('soffice'):
        return 'soffice'
    return 'libreoffice'


# ============================================================================
# MAIN LOGIC
# ============================================================================

def export_images(
    filepath: Path,
    output_dir: Path,
    image_format: str = "png",
    prefix: str = "slide_",
    timeout: int = 120
) -> Dict[str, Any]:
    """
    Export PowerPoint slides as images.
    
    Args:
        filepath: Path to PowerPoint file (must be .pptx)
        output_dir: Directory for output images
        image_format: Image format ('png', 'jpg', 'jpeg')
        prefix: Filename prefix for output images
        timeout: Subprocess timeout in seconds
        
    Returns:
        Dict with export results including file list and sizes
        
    Raises:
        FileNotFoundError: If input file doesn't exist
        ValueError: If invalid format or file type
        RuntimeError: If LibreOffice not installed
        PowerPointAgentError: If export process fails
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not filepath.suffix.lower() == '.pptx':
        raise ValueError(f"Input must be .pptx file, got: {filepath.suffix}")
    
    if image_format.lower() not in ['png', 'jpg', 'jpeg']:
        raise ValueError(f"Format must be png or jpg, got: {image_format}")
    
    format_ext = 'png' if image_format.lower() == 'png' else 'jpg'
    
    if not check_libreoffice():
        raise RuntimeError(
            "LibreOffice not found. Image export requires LibreOffice.\n"
            "Install:\n"
            "  Linux: sudo apt install libreoffice-impress\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    warnings_collected: List[str] = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only operation, no lock needed
        slide_count = agent.get_slide_count()
        presentation_version = agent.get_presentation_version()
    
    base_name = filepath.stem
    pdf_path = output_dir / f"{base_name}.pdf"
    lo_command = get_libreoffice_command()
    
    cmd_pdf = [
        lo_command,
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', str(output_dir),
        str(filepath)
    ]
    
    try:
        result_pdf = subprocess.run(
            cmd_pdf,
            capture_output=True,
            text=True,
            timeout=timeout
        )
    except subprocess.TimeoutExpired:
        raise PowerPointAgentError(
            f"PDF export timed out after {timeout} seconds. "
            f"Try increasing --timeout for large presentations."
        )
    
    if result_pdf.returncode != 0:
        raise PowerPointAgentError(
            f"PDF export failed: {result_pdf.stderr}\n"
            f"Command: {' '.join(cmd_pdf)}"
        )
    
    use_pdftoppm = shutil.which('pdftoppm') is not None
    
    if use_pdftoppm and pdf_path.exists():
        cmd_img = [
            'pdftoppm',
            f"-{format_ext}",
            '-r', '150',
            str(pdf_path),
            str(output_dir / base_name)
        ]
        
        try:
            result_img = subprocess.run(
                cmd_img,
                capture_output=True,
                text=True,
                timeout=timeout
            )
            
            if result_img.returncode != 0:
                warnings_collected.append(
                    f"pdftoppm failed, using LibreOffice direct export"
                )
                _export_direct(filepath, output_dir, format_ext, lo_command, timeout)
                
        except subprocess.TimeoutExpired:
            warnings_collected.append(
                f"pdftoppm timed out, using LibreOffice direct export"
            )
            _export_direct(filepath, output_dir, format_ext, lo_command, timeout)
        
        if pdf_path.exists():
            pdf_path.unlink()
    else:
        if not use_pdftoppm:
            warnings_collected.append(
                "pdftoppm not found, using LibreOffice direct export (may be incomplete)"
            )
        _export_direct(filepath, output_dir, format_ext, lo_command, timeout)
    
    result = _scan_and_process_results(filepath, output_dir, format_ext, prefix)
    
    result["presentation_version"] = presentation_version
    result["slide_count_source"] = slide_count
    result["tool_version"] = __version__
    
    if warnings_collected:
        result["warnings"] = warnings_collected
    
    return result


def _export_direct(
    filepath: Path,
    output_dir: Path,
    format_ext: str,
    lo_command: str,
    timeout: int
) -> None:
    """
    Direct export using LibreOffice (fallback method).
    
    Args:
        filepath: Input PowerPoint file
        output_dir: Output directory
        format_ext: Image format extension
        lo_command: LibreOffice command to use
        timeout: Subprocess timeout in seconds
    """
    cmd = [
        lo_command,
        '--headless',
        '--convert-to', format_ext,
        '--outdir', str(output_dir),
        str(filepath)
    ]
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout
        )
    except subprocess.TimeoutExpired:
        raise PowerPointAgentError(
            f"Direct image export timed out after {timeout} seconds"
        )
    
    if result.returncode != 0:
        raise PowerPointAgentError(
            f"Image export failed: {result.stderr}\n"
            f"Command: {' '.join(cmd)}"
        )


def _scan_and_process_results(
    filepath: Path,
    output_dir: Path,
    format_ext: str,
    prefix: str
) -> Dict[str, Any]:
    """
    Find, rename, and report exported images.
    
    Args:
        filepath: Original input file (for base name)
        output_dir: Directory containing exported images
        format_ext: Image format extension
        prefix: Prefix for renamed files
        
    Returns:
        Dict with export statistics and file list
    """
    base_name = filepath.stem
    
    candidates = sorted(output_dir.glob(f"{base_name}*.{format_ext}"))
    
    if not candidates:
        candidates = sorted(output_dir.glob(f"*.{format_ext}"))
    
    exported_files: List[Path] = []
    
    for i, old_file in enumerate(candidates):
        new_file = output_dir / f"{prefix}{i+1:03d}.{format_ext}"
        
        if old_file != new_file:
            if new_file.exists():
                new_file.unlink()
            old_file.rename(new_file)
            exported_files.append(new_file)
        else:
            exported_files.append(old_file)
    
    if len(exported_files) == 0:
        raise PowerPointAgentError(
            f"Export completed but no image files found in: {output_dir}"
        )
    
    total_size = sum(f.stat().st_size for f in exported_files)
    
    return {
        "status": "success",
        "input_file": str(filepath.resolve()),
        "output_dir": str(output_dir.resolve()),
        "format": format_ext,
        "slides_exported": len(exported_files),
        "files": [str(f.resolve()) for f in exported_files],
        "total_size_bytes": total_size,
        "total_size_mb": round(total_size / (1024 * 1024), 2),
        "average_size_mb": round(total_size / (1024 * 1024) / len(exported_files), 2) if exported_files else 0
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Export PowerPoint slides as images",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Export as PNG
  uv run tools/ppt_export_images.py \\
    --file presentation.pptx \\
    --output-dir slides/ \\
    --format png \\
    --json
  
  # Export as JPG with custom prefix and timeout
  uv run tools/ppt_export_images.py \\
    --file presentation.pptx \\
    --output-dir images/ \\
    --format jpg \\
    --prefix deck_ \\
    --timeout 300 \\
    --json

Output Files:
  Files are named: <prefix><number>.<format>
  Examples: slide_001.png, slide_002.png, deck_001.jpg

Requirements:
  LibreOffice must be installed:
  - Linux: sudo apt install libreoffice-impress
  - macOS: brew install --cask libreoffice
  - Windows: https://www.libreoffice.org/download/

Format Comparison:
  PNG: Lossless, better for text/diagrams, larger files
  JPG: Lossy, better for photos, smaller files
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to export'
    )
    
    parser.add_argument(
        '--output-dir',
        required=True,
        type=Path,
        help='Output directory for images'
    )
    
    parser.add_argument(
        '--format',
        choices=['png', 'jpg', 'jpeg'],
        default='png',
        help='Image format (default: png)'
    )
    
    parser.add_argument(
        '--prefix',
        default='slide_',
        help='Filename prefix (default: slide_)'
    )
    
    parser.add_argument(
        '--timeout',
        type=int,
        default=120,
        help='Subprocess timeout in seconds (default: 120)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = export_images(
            filepath=args.file.resolve(),
            output_dir=args.output_dir.resolve(),
            image_format=args.format,
            prefix=args.prefix,
            timeout=args.timeout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            print(f"✅ Exported {result['slides_exported']} slides to {result['output_dir']}")
            print(f"   Format: {result['format'].upper()}")
            print(f"   Total size: {result['total_size_mb']} MB")
            print(f"   Average: {result['average_size_mb']} MB per slide")
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check input file format (.pptx) and image format (png/jpg)",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except RuntimeError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "RuntimeError",
            "suggestion": "Install LibreOffice: sudo apt install libreoffice-impress",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Check LibreOffice installation and file integrity",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_export_pdf.py
```py
#!/usr/bin/env python3
"""
PowerPoint Export PDF Tool v3.1.1
Export presentation to PDF format using LibreOffice.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_export_pdf.py --file presentation.pptx --output presentation.pdf --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for PDF export:
    - Linux: sudo apt install libreoffice-impress
    - macOS: brew install --cask libreoffice
    - Windows: Download from https://www.libreoffice.org/

Changelog v3.1.1:
    - Added hygiene block for JSON pipeline safety
    - Added presentation_version tracking via PowerPointAgent
    - Added tool_version and slide_count to output
    - Added --timeout argument (default: 300s for large presentations)
    - Fixed cross-filesystem rename with shutil.move
    - Fixed error response format with suggestions
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that JSON parsers only see valid JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import subprocess
import shutil
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def check_libreoffice() -> bool:
    """Check if LibreOffice is installed and accessible."""
    return shutil.which('soffice') is not None or shutil.which('libreoffice') is not None


def get_libreoffice_command() -> str:
    """Get the LibreOffice command for the current system."""
    if shutil.which('soffice'):
        return 'soffice'
    return 'libreoffice'


# ============================================================================
# MAIN LOGIC
# ============================================================================

def export_pdf(
    filepath: Path,
    output: Path,
    timeout: int = 300
) -> Dict[str, Any]:
    """
    Export PowerPoint presentation to PDF.
    
    Args:
        filepath: Path to PowerPoint file (must be .pptx)
        output: Output PDF file path
        timeout: Subprocess timeout in seconds (default: 300)
        
    Returns:
        Dict with export results including file sizes and version info
        
    Raises:
        FileNotFoundError: If input file doesn't exist
        ValueError: If input is not a .pptx file
        RuntimeError: If LibreOffice not installed
        PowerPointAgentError: If export process fails
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not filepath.suffix.lower() == '.pptx':
        raise ValueError(f"Input must be .pptx file, got: {filepath.suffix}")
    
    if not check_libreoffice():
        raise RuntimeError(
            "LibreOffice not found. PDF export requires LibreOffice.\n"
            "Install:\n"
            "  Linux: sudo apt install libreoffice-impress\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only operation, no lock needed
        presentation_version = agent.get_presentation_version()
        slide_count = agent.get_slide_count()
        presentation_info = agent.get_presentation_info()
    
    output.parent.mkdir(parents=True, exist_ok=True)
    
    lo_command = get_libreoffice_command()
    
    cmd = [
        lo_command,
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', str(output.parent.resolve()),
        str(filepath.resolve())
    ]
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout
        )
    except subprocess.TimeoutExpired:
        raise PowerPointAgentError(
            f"PDF export timed out after {timeout} seconds. "
            f"Try increasing --timeout for large presentations (100+ slides may need 5+ minutes)."
        )
    
    if result.returncode != 0:
        raise PowerPointAgentError(
            f"PDF export failed: {result.stderr}\n"
            f"Command: {' '.join(cmd)}"
        )
    
    expected_pdf = output.parent / f"{filepath.stem}.pdf"
    
    if expected_pdf != output:
        if expected_pdf.exists():
            if output.exists():
                output.unlink()
            shutil.move(str(expected_pdf), str(output))
    
    if not output.exists():
        if expected_pdf.exists():
            shutil.move(str(expected_pdf), str(output))
    
    if not output.exists():
        raise PowerPointAgentError(
            f"PDF export completed but output file not found. "
            f"Expected at: {output}"
        )
    
    input_size = filepath.stat().st_size
    output_size = output.stat().st_size
    
    return {
        "status": "success",
        "tool_version": __version__,
        "input_file": str(filepath.resolve()),
        "output_file": str(output.resolve()),
        "presentation_version": presentation_version,
        "slide_count": slide_count,
        "input_size_bytes": input_size,
        "input_size_mb": round(input_size / (1024 * 1024), 2),
        "output_size_bytes": output_size,
        "output_size_mb": round(output_size / (1024 * 1024), 2),
        "size_ratio": round(output_size / input_size, 2) if input_size > 0 else 0,
        "compression_percent": round((1 - output_size / input_size) * 100, 1) if input_size > 0 else 0
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Export PowerPoint presentation to PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic export
  uv run tools/ppt_export_pdf.py \\
    --file presentation.pptx \\
    --output presentation.pdf \\
    --json
  
  # Large presentation with extended timeout
  uv run tools/ppt_export_pdf.py \\
    --file large_deck.pptx \\
    --output reports/output.pdf \\
    --timeout 600 \\
    --json

Requirements:
  LibreOffice must be installed:
  - Linux: sudo apt install libreoffice-impress
  - macOS: brew install --cask libreoffice
  - Windows: https://www.libreoffice.org/download/

Performance Notes:
  - Small decks (<20 slides): ~10-30 seconds
  - Medium decks (20-50 slides): ~1-2 minutes
  - Large decks (100+ slides): ~3-5 minutes
  - Adjust --timeout accordingly

PDF Benefits:
  - Universal compatibility
  - Prevents editing
  - Smaller file size (typically 30-50% of .pptx)
  - Better for printing

Limitations:
  - Animations not preserved
  - Embedded videos become static
  - Speaker notes not included
  - Transitions removed
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to export'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output PDF file path'
    )
    
    parser.add_argument(
        '--timeout',
        type=int,
        default=300,
        help='Export timeout in seconds (default: 300)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        output_path = args.output
        if not output_path.suffix.lower() == '.pdf':
            output_path = output_path.with_suffix('.pdf')
        
        result = export_pdf(
            filepath=args.file.resolve(),
            output=output_path.resolve(),
            timeout=args.timeout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            print(f"✅ Exported to PDF: {result['output_file']}")
            print(f"   Slides: {result['slide_count']}")
            print(f"   Input: {result['input_size_mb']} MB")
            print(f"   Output: {result['output_size_mb']} MB")
            print(f"   Compression: {result['compression_percent']}%")
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Ensure input file is a .pptx PowerPoint file",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except RuntimeError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "RuntimeError",
            "suggestion": "Install LibreOffice: sudo apt install libreoffice-impress",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Check LibreOffice installation and try increasing --timeout",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_extract_notes.py
```py
#!/usr/bin/env python3
"""
PowerPoint Extract Notes Tool v3.1.0
Extract speaker notes from all slides in a presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_extract_notes.py --file presentation.pptx --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent

__version__ = "3.1.0"


def extract_notes(filepath: Path) -> Dict[str, Any]:
    """
    Extract speaker notes from all slides in a presentation.
    
    This is a read-only operation that does not modify the file.
    
    Args:
        filepath: Path to PowerPoint file (.pptx only)
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute path to file
            - presentation_version: Version hash of presentation
            - total_slides: Total number of slides
            - notes_found: Count of slides that have notes content
            - notes: Dict mapping slide index (as string) to notes text
            - tool_version: Tool version string
            
    Raises:
        FileNotFoundError: If PowerPoint file doesn't exist
        ValueError: If file format is not .pptx
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError(
            f"Invalid file format '{filepath.suffix}'. Only .pptx files are supported."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        presentation_version = agent.get_presentation_version()
        notes = agent.extract_notes()
        total_slides = agent.get_slide_count()
    
    notes_with_content = sum(1 for text in notes.values() if text and text.strip())
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "presentation_version": presentation_version,
        "total_slides": total_slides,
        "notes_found": notes_with_content,
        "notes": notes,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Extract speaker notes from all slides in a PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Extract notes from presentation
    uv run tools/ppt_extract_notes.py --file presentation.pptx --json
    
    # Save notes to file
    uv run tools/ppt_extract_notes.py --file presentation.pptx --json > notes.json

Output Format:
    {
      "status": "success",
      "file": "/path/to/presentation.pptx",
      "presentation_version": "a1b2c3d4e5f6g7h8",
      "total_slides": 10,
      "notes_found": 5,
      "notes": {
        "0": "Speaker notes for slide 1...",
        "1": "",
        "2": "Important talking points...",
        ...
      },
      "tool_version": "3.1.0"
    }

Use Cases:
    - Export notes for speaker preparation
    - Backup presentation scripts
    - Convert notes to other formats
    - Accessibility: extract text alternatives
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='Path to PowerPoint file (.pptx)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output as JSON (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = extract_notes(filepath=args.file)
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Ensure file has .pptx extension."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_format_chart.py
```py
#!/usr/bin/env python3
"""
PowerPoint Format Chart Tool v3.1.0
Format existing chart (title, legend position)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_format_chart.py --file presentation.pptx --slide 1 --chart 0 --title "Revenue Growth" --legend bottom --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Limitations:
    python-pptx has limited chart formatting support. This tool handles:
    - Chart title text
    - Legend position
    
    Not supported (requires PowerPoint):
    - Individual series colors
    - Axis formatting
    - Data labels
    - Chart styles/templates
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Legend positions
LEGEND_POSITIONS = ['bottom', 'left', 'right', 'top', 'none']


def format_chart(
    filepath: Path,
    slide_index: int,
    chart_index: int = 0,
    title: Optional[str] = None,
    legend_position: Optional[str] = None
) -> Dict[str, Any]:
    """
    Format an existing chart on a slide.
    
    Updates chart title and/or legend position. At least one
    formatting option must be specified.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the chart (0-based)
        chart_index: Index of the chart on the slide (0-based, default: 0)
        title: New chart title text (optional)
        legend_position: Legend position - 'bottom', 'left', 'right', 'top', 'none'
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - chart_index: Index of the chart
            - formatting_applied: Dict with applied formatting
            - presentation_version_before: State hash before formatting
            - presentation_version_after: State hash after formatting
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ValueError: If no formatting options specified or chart not found
        
    Example:
        >>> result = format_chart(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=1,
        ...     chart_index=0,
        ...     title="Revenue Growth Trend",
        ...     legend_position="bottom"
        ... )
        >>> print(result["formatting_applied"]["title"])
        'Revenue Growth Trend'
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate at least one option is specified
    if title is None and legend_position is None:
        raise ValueError(
            "At least one formatting option (--title or --legend) must be specified"
        )
    
    # Validate legend position if provided
    if legend_position is not None and legend_position not in LEGEND_POSITIONS:
        raise ValueError(
            f"Invalid legend position: {legend_position}. "
            f"Valid options: {', '.join(LEGEND_POSITIONS)}"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE formatting
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Get slide info to check for charts
        slide_info = agent.get_slide_info(slide_index)
        
        # Count charts on slide (check shapes for chart type)
        chart_count = 0
        for shape in slide_info.get("shapes", []):
            if shape.get("type") == "CHART" or shape.get("has_chart", False):
                chart_count += 1
        
        # If we couldn't detect charts from slide_info, try the operation anyway
        # The core method will raise if chart doesn't exist
        if chart_count == 0:
            # Could be that slide_info doesn't expose chart detection
            # Let the core method handle validation
            pass
        elif not 0 <= chart_index < chart_count:
            raise ValueError(
                f"Chart index {chart_index} out of range. "
                f"Slide has {chart_count} chart(s) (indices 0-{chart_count - 1})."
            )
        
        # Format chart
        agent.format_chart(
            slide_index=slide_index,
            chart_index=chart_index,
            title=title,
            legend_position=legend_position
        )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER formatting
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Build formatting applied dict
    formatting_applied: Dict[str, Any] = {}
    if title is not None:
        formatting_applied["title"] = title
    if legend_position is not None:
        formatting_applied["legend_position"] = legend_position
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "chart_index": chart_index,
        "formatting_applied": formatting_applied,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Format PowerPoint chart (title, legend)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set chart title
  uv run tools/ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart 0 \\
    --title "Revenue Growth Trend" \\
    --json
  
  # Position legend at bottom
  uv run tools/ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --chart 0 \\
    --legend bottom \\
    --json
  
  # Set both title and legend
  uv run tools/ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --chart 0 \\
    --title "Q4 Performance" \\
    --legend right \\
    --json
  
  # Hide legend
  uv run tools/ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart 0 \\
    --legend none \\
    --json

Legend Positions:
  bottom  Below chart (common for wide charts)
  right   Right side of chart (default)
  top     Above chart
  left    Left side of chart
  none    Hide legend entirely

Finding Charts:
  Charts are indexed in order they appear on the slide (0, 1, 2...).
  Use ppt_get_slide_info.py to find charts:
  uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 1 --json

Best Practices:
  - Keep titles concise and descriptive
  - Use 'bottom' legend for wide charts
  - Use 'right' legend for tall charts
  - Hide legend if only one series
  - Match title to chart type (e.g., "Trend" for line charts)

⚠️ Formatting Limitations:
  python-pptx has limited chart formatting support.
  
  Supported by this tool:
  ✓ Chart title text
  ✓ Legend position
  
  Not supported (use PowerPoint directly):
  ✗ Individual series colors
  ✗ Axis formatting (labels, scale)
  ✗ Data labels
  ✗ Chart styles/templates
  ✗ Gridlines

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 1,
    "chart_index": 0,
    "formatting_applied": {
      "title": "Revenue Growth Trend",
      "legend_position": "bottom"
    },
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--chart',
        type=int,
        default=0,
        help='Chart index on slide (default: 0)'
    )
    
    parser.add_argument(
        '--title',
        help='Chart title text'
    )
    
    parser.add_argument(
        '--legend',
        choices=LEGEND_POSITIONS,
        help='Legend position (bottom, left, right, top, none)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = format_chart(
            filepath=args.file,
            slide_index=args.slide,
            chart_index=args.chart,
            title=args.title,
            legend_position=args.legend
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Specify --title and/or --legend, and verify chart index"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_format_shape.py
```py
#!/usr/bin/env python3
"""
PowerPoint Format Shape Tool v3.1.0
Update styling of existing shapes including fill, line, opacity, and text formatting.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_format_shape.py --file presentation.pptx --slide 0 --shape 1 \\
        --fill-color "#FF0000" --fill-opacity 0.8 --json

Exit Codes:
    0: Success
    1: Error occurred

Note: The --transparency parameter is DEPRECATED. Use --fill-opacity instead.
      Opacity: 0.0 = invisible, 1.0 = opaque (opposite of transparency)
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
    ColorHelper,
)

__version__ = "3.1.0"

COLOR_PRESETS = {
    "primary": "#0070C0",
    "secondary": "#595959",
    "accent": "#ED7D31",
    "success": "#70AD47",
    "warning": "#FFC000",
    "danger": "#C00000",
    "white": "#FFFFFF",
    "black": "#000000",
    "light_gray": "#D9D9D9",
    "dark_gray": "#404040",
    "transparent": None,
}

OPACITY_PRESETS = {
    "opaque": 1.0,
    "subtle": 0.85,
    "light": 0.7,
    "medium": 0.5,
    "heavy": 0.3,
    "very_light": 0.15,
    "nearly_invisible": 0.05,
}


def resolve_color(color: Optional[str]) -> Optional[str]:
    """Resolve color value, handling presets and hex formats."""
    if color is None:
        return None
    
    color_lower = color.lower().strip()
    
    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]
    
    if color_lower in ("none", "transparent", "clear"):
        return None
    
    if not color.startswith('#') and len(color) == 6:
        try:
            int(color, 16)
            return f"#{color}"
        except ValueError:
            pass
    
    return color


def resolve_opacity(value: Optional[str], is_transparency: bool = False) -> Optional[float]:
    """
    Resolve opacity value, handling presets and numeric values.
    
    Args:
        value: Opacity specification (float, preset name, or percentage string)
        is_transparency: If True, value is transparency (inverted)
        
    Returns:
        Opacity as float (0.0 = invisible, 1.0 = opaque) or None
    """
    if value is None:
        return None
    
    result: float
    
    if isinstance(value, str):
        value_lower = value.lower().strip()
        
        if value_lower in OPACITY_PRESETS:
            result = OPACITY_PRESETS[value_lower]
        elif value_lower.endswith('%'):
            try:
                pct = float(value_lower[:-1]) / 100.0
                result = pct if not is_transparency else (1.0 - pct)
            except ValueError:
                raise ValueError(f"Invalid opacity value: {value}")
        else:
            try:
                result = float(value_lower)
                if is_transparency:
                    result = 1.0 - result
            except ValueError:
                raise ValueError(f"Invalid opacity value: {value}")
    else:
        result = float(value)
        if is_transparency:
            result = 1.0 - result
    
    return max(0.0, min(1.0, result))


def validate_formatting_params(
    fill_color: Optional[str],
    line_color: Optional[str],
    fill_opacity: Optional[float]
) -> Dict[str, Any]:
    """Validate formatting parameters and generate warnings."""
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    if fill_opacity is not None:
        if fill_opacity < 0.05:
            warnings.append(
                f"Fill opacity {fill_opacity} is very low. Shape may be nearly invisible."
            )
        validation_results["fill_opacity"] = fill_opacity
    
    if fill_color:
        try:
            ColorHelper.from_hex(fill_color)
            validation_results["fill_color_valid"] = True
        except Exception as e:
            validation_results["fill_color_valid"] = False
            validation_results["fill_color_error"] = str(e)
            warnings.append(f"Invalid fill color format: {fill_color}")
    
    if line_color:
        try:
            ColorHelper.from_hex(line_color)
            validation_results["line_color_valid"] = True
        except Exception as e:
            validation_results["line_color_valid"] = False
            warnings.append(f"Invalid line color format: {line_color}")
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def format_shape(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    line_width: Optional[float] = None,
    fill_opacity: Optional[float] = None,
    line_opacity: Optional[float] = None,
    text_color: Optional[str] = None,
    text_size: Optional[int] = None,
    text_bold: Optional[bool] = None
) -> Dict[str, Any]:
    """
    Format existing shape with comprehensive styling options.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        shape_index: Target shape index (0-based)
        fill_color: Fill color (hex or preset name)
        line_color: Line/border color (hex or preset name)
        line_width: Line width in points
        fill_opacity: Fill opacity (0.0=invisible to 1.0=opaque)
        line_opacity: Line opacity (0.0=invisible to 1.0=opaque)
        text_color: Text color within shape
        text_size: Text size in points
        text_bold: Text bold setting
        
    Returns:
        Result dict with formatting details
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If no formatting options or file format invalid
        SlideNotFoundError: If slide index invalid
        ShapeNotFoundError: If shape index invalid
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    formatting_options = [
        fill_color, line_color, line_width, fill_opacity, line_opacity,
        text_color, text_size, text_bold
    ]
    if all(v is None for v in formatting_options):
        raise ValueError(
            "At least one formatting option required. "
            "Use --fill-color, --line-color, --fill-opacity, etc."
        )
    
    resolved_fill = resolve_color(fill_color)
    resolved_line = resolve_color(line_color)
    resolved_text_color = resolve_color(text_color)
    
    validation = validate_formatting_params(
        fill_color=resolved_fill,
        line_color=resolved_line,
        fill_opacity=fill_opacity
    )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        slide_info = agent.get_slide_info(slide_index)
        shape_count = slide_info.get("shape_count", 0)
        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{shape_count - 1})",
                details={"requested": shape_index, "available": shape_count}
            )
        
        version_before = agent.get_presentation_version()
        
        format_result = agent.format_shape(
            slide_index=slide_index,
            shape_index=shape_index,
            fill_color=resolved_fill,
            line_color=resolved_line,
            line_width=line_width,
            fill_opacity=fill_opacity,
            line_opacity=line_opacity
        )
        
        text_formatted = False
        if any(v is not None for v in [text_color, text_size, text_bold]):
            try:
                agent.format_text(
                    slide_index=slide_index,
                    shape_index=shape_index,
                    color=resolved_text_color,
                    font_size=text_size,
                    bold=text_bold
                )
                text_formatted = True
            except Exception as e:
                validation["warnings"].append(f"Could not format text: {e}")
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    result: Dict[str, Any] = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "formatting_applied": {
            "fill_color": resolved_fill,
            "fill_opacity": fill_opacity,
            "line_color": resolved_line,
            "line_opacity": line_opacity,
            "line_width": line_width,
            "text_color": resolved_text_color if text_formatted else None,
            "text_size": text_size if text_formatted else None,
            "text_bold": text_bold if text_formatted else None
        },
        "changes_from_core": format_result.get("changes_applied", []) if isinstance(format_result, dict) else [],
        "text_formatted": text_formatted,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if validation["validation_results"]:
        result["validation"] = validation["validation_results"]
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Format existing PowerPoint shape",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
FORMATTING OPTIONS:
  --fill-color     Shape fill color (hex or preset)
  --fill-opacity   Fill opacity: 0.0 (invisible) to 1.0 (opaque)
  --line-color     Border/line color (hex or preset)
  --line-opacity   Line opacity: 0.0 (invisible) to 1.0 (opaque)
  --line-width     Border width in points
  --text-color     Text color within shape
  --text-size      Text size in points
  --text-bold      Make text bold

COLOR PRESETS:
  primary (#0070C0)    secondary (#595959)    accent (#ED7D31)
  success (#70AD47)    warning (#FFC000)      danger (#C00000)
  white (#FFFFFF)      black (#000000)        transparent (none)

OPACITY PRESETS:
  opaque (1.0)         subtle (0.85)          light (0.7)
  medium (0.5)         heavy (0.3)            very_light (0.15)

EXAMPLES:

  # Change fill color
  uv run tools/ppt_format_shape.py --file deck.pptx --slide 0 --shape 1 \\
    --fill-color "#FF0000" --json

  # Semi-transparent overlay
  uv run tools/ppt_format_shape.py --file deck.pptx --slide 1 --shape 0 \\
    --fill-color black --fill-opacity 0.5 --json

  # Format text within shape
  uv run tools/ppt_format_shape.py --file deck.pptx --slide 0 --shape 3 \\
    --fill-color primary --text-color white --text-size 24 --text-bold --json

FINDING SHAPE INDEX:
  Use ppt_get_slide_info.py to find shape indices:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

DEPRECATED:
  --transparency is deprecated. Use --fill-opacity instead.
  transparency = 1.0 - fill_opacity (values are inverted)
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path (.pptx)')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--fill-color', help='Fill color: hex or preset')
    parser.add_argument('--fill-opacity', help='Fill opacity: 0.0 (invisible) to 1.0 (opaque)')
    parser.add_argument('--line-color', help='Line/border color')
    parser.add_argument('--line-opacity', help='Line opacity: 0.0 to 1.0')
    parser.add_argument('--line-width', type=float, help='Line width in points')
    parser.add_argument('--transparency', help='DEPRECATED: Use --fill-opacity. Transparency: 0.0 (opaque) to 1.0 (invisible)')
    parser.add_argument('--text-color', help='Text color within shape')
    parser.add_argument('--text-size', type=int, help='Text size in points')
    parser.add_argument('--text-bold', action='store_true', help='Make text bold')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON (default: true)')
    
    args = parser.parse_args()
    
    try:
        fill_opacity: Optional[float] = None
        deprecation_warning: Optional[str] = None
        
        if args.fill_opacity is not None:
            fill_opacity = resolve_opacity(args.fill_opacity, is_transparency=False)
        elif args.transparency is not None:
            fill_opacity = resolve_opacity(args.transparency, is_transparency=True)
            deprecation_warning = (
                "--transparency is deprecated. Use --fill-opacity instead. "
                f"Converted transparency {args.transparency} to fill_opacity {fill_opacity}"
            )
        
        line_opacity: Optional[float] = None
        if args.line_opacity is not None:
            line_opacity = resolve_opacity(args.line_opacity, is_transparency=False)
        
        result = format_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width,
            fill_opacity=fill_opacity,
            line_opacity=line_opacity,
            text_color=args.text_color,
            text_size=args.text_size,
            text_bold=args.text_bold if args.text_bold else None
        )
        
        if deprecation_warning:
            if "warnings" not in result:
                result["warnings"] = []
            result["warnings"].insert(0, deprecation_warning)
            result["status"] = "warning"
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Provide at least one formatting option and check opacity values."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check the presentation file is valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_format_table.py
```py
#!/usr/bin/env python3
"""
PowerPoint Format Table Tool v3.1.1
Style and format existing tables in presentations.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_format_table.py --file presentation.pptx --slide 0 --shape 2 --header-fill "#0070C0" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

This tool formats existing tables by applying styling options including:
- Header row colors and formatting
- Data row colors with optional banding
- Font styling (name, size, color)
- Border styling (color, width)
- First column highlighting
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def parse_color(color_str: Optional[str]) -> Optional[str]:
    """
    Parse and validate a color string.
    
    Args:
        color_str: Color in #RRGGBB or RRGGBB format
        
    Returns:
        Normalized color string with # prefix, or None
    """
    if not color_str:
        return None
    
    color = color_str.strip()
    if not color.startswith('#'):
        color = '#' + color
    
    if len(color) != 7:
        raise ValueError(f"Invalid color format: {color_str}. Expected #RRGGBB")
    
    try:
        int(color[1:], 16)
    except ValueError:
        raise ValueError(f"Invalid color format: {color_str}. Expected hexadecimal")
    
    return color.upper()


def is_table_shape(shape) -> bool:
    """
    Check if a shape is a table.
    
    Args:
        shape: Shape object from python-pptx
        
    Returns:
        True if shape is a table, False otherwise
    """
    return hasattr(shape, 'table') and shape.has_table


# ============================================================================
# MAIN LOGIC
# ============================================================================

def format_table(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    header_fill: Optional[str] = None,
    header_text: Optional[str] = None,
    row_fill: Optional[str] = None,
    alt_row_fill: Optional[str] = None,
    text_color: Optional[str] = None,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    border_color: Optional[str] = None,
    border_width: Optional[float] = None,
    first_col_highlight: bool = False,
    banding: bool = False
) -> Dict[str, Any]:
    """
    Format an existing table in a PowerPoint presentation.
    
    Args:
        filepath: Path to the PowerPoint file
        slide_index: Index of the slide containing the table (0-based)
        shape_index: Index of the table shape on the slide
        header_fill: Fill color for header row (#RRGGBB)
        header_text: Text color for header row
        row_fill: Fill color for data rows
        alt_row_fill: Alternating row fill color for banding
        text_color: Default text color for all cells
        font_name: Font family name
        font_size: Font size in points
        border_color: Border color
        border_width: Border width in points
        first_col_highlight: Highlight first column
        banding: Enable row banding (requires alt_row_fill)
        
    Returns:
        Dict with formatting results
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is invalid
        ShapeNotFoundError: If shape index is invalid
        ValueError: If shape is not a table
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    header_fill_parsed = parse_color(header_fill)
    header_text_parsed = parse_color(header_text)
    row_fill_parsed = parse_color(row_fill)
    alt_row_fill_parsed = parse_color(alt_row_fill)
    text_color_parsed = parse_color(text_color)
    border_color_parsed = parse_color(border_color)
    
    changes_applied: List[str] = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        slide = agent.prs.slides[slide_index]
        
        if not 0 <= shape_index < len(slide.shapes):
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{len(slide.shapes) - 1})",
                details={
                    "requested_index": shape_index,
                    "available_shapes": len(slide.shapes)
                }
            )
        
        shape = slide.shapes[shape_index]
        
        if not is_table_shape(shape):
            raise ValueError(
                f"Shape at index {shape_index} is not a table. "
                f"Shape type: {shape.shape_type}"
            )
        
        table = shape.table
        row_count = len(table.rows)
        col_count = len(table.columns)
        
        from pptx.util import Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        
        def hex_to_rgb(hex_color: str) -> RGBColor:
            hex_color = hex_color.lstrip('#')
            return RGBColor(
                int(hex_color[0:2], 16),
                int(hex_color[2:4], 16),
                int(hex_color[4:6], 16)
            )
        
        if header_fill_parsed and row_count > 0:
            for cell in table.rows[0].cells:
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_to_rgb(header_fill_parsed)
            changes_applied.append(f"header_fill={header_fill_parsed}")
        
        if header_text_parsed and row_count > 0:
            for cell in table.rows[0].cells:
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = hex_to_rgb(header_text_parsed)
            changes_applied.append(f"header_text={header_text_parsed}")
        
        if row_fill_parsed or (banding and alt_row_fill_parsed):
            for row_idx in range(1, row_count):
                if banding and alt_row_fill_parsed and row_idx % 2 == 0:
                    fill_color = alt_row_fill_parsed
                elif row_fill_parsed:
                    fill_color = row_fill_parsed
                else:
                    continue
                    
                for cell in table.rows[row_idx].cells:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = hex_to_rgb(fill_color)
            
            if row_fill_parsed:
                changes_applied.append(f"row_fill={row_fill_parsed}")
            if banding and alt_row_fill_parsed:
                changes_applied.append(f"banding_enabled=True")
                changes_applied.append(f"alt_row_fill={alt_row_fill_parsed}")
        
        if text_color_parsed:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = hex_to_rgb(text_color_parsed)
            changes_applied.append(f"text_color={text_color_parsed}")
        
        if font_name:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = font_name
            changes_applied.append(f"font_name={font_name}")
        
        if font_size:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(font_size)
            changes_applied.append(f"font_size={font_size}pt")
        
        if first_col_highlight and header_fill_parsed and col_count > 0:
            for row in table.rows:
                cell = row.cells[0]
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_to_rgb(header_fill_parsed)
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
            changes_applied.append("first_col_highlight=True")
        
        agent.save()
        
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "table_info": {
            "rows": row_count,
            "columns": col_count,
            "has_header": True
        },
        "changes_applied": changes_applied,
        "changes_count": len(changes_applied),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Format existing tables in PowerPoint presentations",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Format header row with blue fill
  uv run tools/ppt_format_table.py \\
    --file presentation.pptx --slide 0 --shape 2 \\
    --header-fill "#0070C0" --header-text "#FFFFFF" --json

  # Enable row banding
  uv run tools/ppt_format_table.py \\
    --file presentation.pptx --slide 1 --shape 3 \\
    --row-fill "#FFFFFF" --alt-row-fill "#F0F0F0" --banding --json

  # Complete formatting
  uv run tools/ppt_format_table.py \\
    --file presentation.pptx --slide 0 --shape 2 \\
    --header-fill "#0070C0" --header-text "#FFFFFF" \\
    --row-fill "#FFFFFF" --text-color "#333333" \\
    --font-name "Calibri" --font-size 11 \\
    --first-col --json

Color Format:
  Colors must be in #RRGGBB hexadecimal format.
  Examples: #0070C0 (blue), #FFFFFF (white), #333333 (dark gray)

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "shape_index": 2,
    "table_info": {"rows": 5, "columns": 4, "has_header": true},
    "changes_applied": ["header_fill=#0070C0", "header_text=#FFFFFF"],
    "presentation_version_before": "a1b2c3...",
    "presentation_version_after": "d4e5f6...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--shape',
        required=True,
        type=int,
        help='Shape index of the table'
    )
    
    parser.add_argument(
        '--header-fill',
        type=str,
        help='Header row fill color (#RRGGBB)'
    )
    
    parser.add_argument(
        '--header-text',
        type=str,
        help='Header row text color (#RRGGBB)'
    )
    
    parser.add_argument(
        '--row-fill',
        type=str,
        help='Data row fill color (#RRGGBB)'
    )
    
    parser.add_argument(
        '--alt-row-fill',
        type=str,
        help='Alternating row fill color for banding (#RRGGBB)'
    )
    
    parser.add_argument(
        '--text-color',
        type=str,
        help='Default text color (#RRGGBB)'
    )
    
    parser.add_argument(
        '--font-name',
        type=str,
        help='Font family name (e.g., "Calibri")'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        help='Font size in points'
    )
    
    parser.add_argument(
        '--border-color',
        type=str,
        help='Border color (#RRGGBB)'
    )
    
    parser.add_argument(
        '--border-width',
        type=float,
        help='Border width in points'
    )
    
    parser.add_argument(
        '--first-col',
        action='store_true',
        help='Highlight first column like header'
    )
    
    parser.add_argument(
        '--banding',
        action='store_true',
        help='Enable alternating row colors (requires --alt-row-fill)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = format_table(
            filepath=args.file.resolve(),
            slide_index=args.slide,
            shape_index=args.shape,
            header_fill=args.header_fill,
            header_text=args.header_text,
            row_fill=args.row_fill,
            alt_row_fill=args.alt_row_fill,
            text_color=args.text_color,
            font_name=args.font_name,
            font_size=args.font_size,
            border_color=args.border_color,
            border_width=args.border_width,
            first_col_highlight=args.first_col,
            banding=args.banding
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Ensure the shape is a table and colors are in #RRGGBB format",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check file integrity and table structure",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_format_text.py
```py
#!/usr/bin/env python3
"""
PowerPoint Format Text Tool v3.1.0
Format existing text with accessibility validation and contrast checking

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Features:
    - Font name, size, color, bold, italic formatting
    - WCAG 2.1 AA/AAA color contrast validation
    - Font size accessibility warnings (<12pt)
    - Before/after formatting comparison
    - Detailed validation results and recommendations

Usage:
    uv run tools/ppt_format_text.py --file deck.pptx --slide 0 --shape 0 --font-name "Arial" --font-size 24 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Accessibility:
    - Minimum font size: 12pt (14pt recommended for presentations)
    - Color contrast: 4.5:1 for normal text, 3:1 for large text (≥18pt)
    - Tool validates and warns about accessibility issues
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import math
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Define fallback exception if not available in core
try:
    from core.powerpoint_agent_core import ShapeNotFoundError
except ImportError:
    class ShapeNotFoundError(PowerPointAgentError):
        """Exception raised when shape is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)

# Color helper functions (fallback if not in core)
try:
    from core.powerpoint_agent_core import ColorHelper, RGBColor
except ImportError:
    # Define minimal color helpers locally
    class RGBColor:
        """Simple RGB color class."""
        def __init__(self, r: int, g: int, b: int):
            self.r = r
            self.g = g
            self.b = b
    
    class ColorHelper:
        """Color utilities for accessibility checking."""
        
        @staticmethod
        def from_hex(hex_color: str) -> RGBColor:
            """Convert hex color to RGBColor."""
            hex_color = hex_color.lstrip('#')
            if len(hex_color) != 6:
                raise ValueError(f"Invalid hex color format: #{hex_color}")
            try:
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                return RGBColor(r, g, b)
            except ValueError:
                raise ValueError(f"Invalid hex color: #{hex_color}")
        
        @staticmethod
        def _relative_luminance(color: RGBColor) -> float:
            """Calculate relative luminance per WCAG 2.1."""
            def channel_luminance(c: int) -> float:
                c_srgb = c / 255.0
                if c_srgb <= 0.03928:
                    return c_srgb / 12.92
                else:
                    return ((c_srgb + 0.055) / 1.055) ** 2.4
            
            return (0.2126 * channel_luminance(color.r) + 
                    0.7152 * channel_luminance(color.g) + 
                    0.0722 * channel_luminance(color.b))
        
        @staticmethod
        def contrast_ratio(color1: RGBColor, color2: RGBColor) -> float:
            """Calculate contrast ratio between two colors per WCAG 2.1."""
            l1 = ColorHelper._relative_luminance(color1)
            l2 = ColorHelper._relative_luminance(color2)
            
            lighter = max(l1, l2)
            darker = min(l1, l2)
            
            return (lighter + 0.05) / (darker + 0.05)
        
        @staticmethod
        def meets_wcag(text_color: RGBColor, bg_color: RGBColor, is_large_text: bool = False) -> bool:
            """Check if colors meet WCAG AA contrast requirements."""
            ratio = ColorHelper.contrast_ratio(text_color, bg_color)
            required = 3.0 if is_large_text else 4.5
            return ratio >= required


def validate_formatting(
    font_size: Optional[int] = None,
    color: Optional[str] = None,
    current_font_size: Optional[int] = None
) -> Dict[str, Any]:
    """
    Validate formatting parameters against accessibility guidelines.
    
    Args:
        font_size: New font size to validate
        color: New color to validate (hex format)
        current_font_size: Current font size for comparison
        
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    # Font size validation
    if font_size is not None:
        validation_results["font_size"] = font_size
        validation_results["font_size_ok"] = font_size >= 12
        
        if font_size < 10:
            warnings.append(
                f"Font size {font_size}pt is extremely small. "
                "Minimum recommended: 12pt for handouts, 14pt for presentations."
            )
        elif font_size < 12:
            warnings.append(
                f"Font size {font_size}pt is below minimum recommended 12pt. "
                "Audience may struggle to read."
            )
            recommendations.append("Use 12pt minimum, 14pt+ for projected content")
        elif font_size < 14:
            recommendations.append(
                f"Font size {font_size}pt is acceptable for handouts but consider 14pt+ for projected presentations"
            )
        
        # Check if decreasing size
        if current_font_size and font_size < current_font_size:
            diff = current_font_size - font_size
            recommendations.append(
                f"Decreasing font size by {diff}pt (from {current_font_size}pt to {font_size}pt). "
                "Verify readability on target display."
            )
    
    # Color contrast validation
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            bg_color = RGBColor(255, 255, 255)  # Assume white background
            
            # Determine if large text
            effective_font_size = font_size if font_size else (current_font_size if current_font_size else 18)
            is_large_text = effective_font_size >= 18
            
            contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
            wcag_aa = ColorHelper.meets_wcag(text_color, bg_color, is_large_text)
            
            validation_results["color_contrast"] = {
                "color": color,
                "ratio": round(contrast_ratio, 2),
                "wcag_aa": wcag_aa,
                "is_large_text": is_large_text,
                "required_ratio": 3.0 if is_large_text else 4.5
            }
            
            if not wcag_aa:
                required = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color {color} has contrast ratio {contrast_ratio:.2f}:1 "
                    f"(WCAG AA requires {required}:1 for {'large' if is_large_text else 'normal'} text). "
                    "May not meet accessibility standards."
                )
                recommendations.append(
                    "Use high-contrast colors: #000000 (black), #333333 (dark gray), #0070C0 (dark blue)"
                )
            elif contrast_ratio < 7.0:
                recommendations.append(
                    f"Color contrast {contrast_ratio:.2f}:1 meets WCAG AA but not AAA (7:1). "
                    "Consider darker color for maximum accessibility."
                )
        except ValueError as e:
            validation_results["color_error"] = str(e)
            warnings.append(f"Invalid color format: {color}. Use hex format like #FF0000")
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results
    }


def format_text(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    color: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None
) -> Dict[str, Any]:
    """
    Format text in a shape with validation and accessibility checking.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        shape_index: Shape index (0-based)
        font_name: Optional font name to apply
        font_size: Optional font size in points
        color: Optional text color (hex format, e.g., "#0070C0")
        bold: Optional bold setting (True/False/None for no change)
        italic: Optional italic setting (True/False/None for no change)
        
    Returns:
        Dict containing:
            - status: "success" or "warning" (if accessibility issues)
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the shape
            - before: Original formatting state
            - after: New formatting applied
            - changes_applied: List of changed properties
            - validation: Validation results
            - warnings: Accessibility/formatting warnings (if any)
            - recommendations: Suggested improvements (if any)
            - presentation_version_before: State hash before modification
            - presentation_version_after: State hash after modification
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ShapeNotFoundError: If shape index is out of range
        ValueError: If no formatting options provided or shape has no text
        
    Example:
        >>> result = format_text(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     shape_index=2,
        ...     font_size=18,
        ...     color="#0070C0",
        ...     bold=True
        ... )
        >>> print(result["changes_applied"])
        ['font_size', 'color', 'bold']
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Check that at least one formatting option is provided
    if all(v is None for v in [font_name, font_size, color, bold, italic]):
        raise ValueError(
            "At least one formatting option must be specified. "
            "Use --font-name, --font-size, --color, --bold, or --italic"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE formatting
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Get slide info to validate shape index
        slide_info = agent.get_slide_info(slide_index)
        shape_count = slide_info.get("shape_count", len(slide_info.get("shapes", [])))
        
        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{shape_count - 1})",
                details={
                    "requested_index": shape_index,
                    "available_shapes": shape_count
                }
            )
        
        # Check if shape has text
        shapes = slide_info.get("shapes", [])
        shape_info = shapes[shape_index] if shape_index < len(shapes) else {}
        
        if not shape_info.get("has_text", False):
            raise ValueError(
                f"Shape {shape_index} ({shape_info.get('type', 'unknown')}) does not contain text. "
                "Cannot format non-text shape. Use ppt_get_slide_info.py to find text-containing shapes."
            )
        
        # Extract current formatting info
        before_formatting = {
            "shape_type": shape_info.get("type"),
            "shape_name": shape_info.get("name"),
            "has_text": shape_info.get("has_text", False)
        }
        
        # Try to get current font size for validation
        current_font_size = None
        try:
            # Access slide directly for font size extraction
            slide = agent.prs.slides[slide_index]
            shape = slide.shapes[shape_index]
            if hasattr(shape, 'text_frame') and shape.text_frame.paragraphs:
                first_para = shape.text_frame.paragraphs[0]
                if first_para.runs and first_para.runs[0].font.size:
                    current_font_size = int(first_para.runs[0].font.size.pt)
                    before_formatting["font_size"] = current_font_size
                elif first_para.font.size:
                    current_font_size = int(first_para.font.size.pt)
                    before_formatting["font_size"] = current_font_size
        except Exception:
            pass  # Continue without current font size
        
        # Validate formatting parameters
        validation = validate_formatting(font_size, color, current_font_size)
        
        # Apply formatting via core
        agent.format_text(
            slide_index=slide_index,
            shape_index=shape_index,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color
        )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER formatting
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Determine status based on warnings
    status = "success" if len(validation["warnings"]) == 0 else "warning"
    
    # Build after formatting dict
    after_formatting: Dict[str, Any] = {}
    if font_name is not None:
        after_formatting["font_name"] = font_name
    if font_size is not None:
        after_formatting["font_size"] = font_size
    if color is not None:
        after_formatting["color"] = color
    if bold is not None:
        after_formatting["bold"] = bold
    if italic is not None:
        after_formatting["italic"] = italic
    
    # Build result
    result: Dict[str, Any] = {
        "status": status,
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "before": before_formatting,
        "after": after_formatting,
        "changes_applied": list(after_formatting.keys()),
        "validation": validation["validation_results"],
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Format text in PowerPoint shape with accessibility validation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Change font and size
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 0 \\
    --font-name "Arial" \\
    --font-size 24 \\
    --json
  
  # Make text bold and colored (with validation)
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --shape 2 \\
    --bold \\
    --color "#0070C0" \\
    --json
  
  # Comprehensive formatting
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --font-name "Calibri" \\
    --font-size 18 \\
    --bold \\
    --color "#000000" \\
    --json
  
  # Fix accessibility issue
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --shape 3 \\
    --font-size 16 \\
    --color "#333333" \\
    --json
  
  # Remove bold/italic
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --shape 1 \\
    --no-bold \\
    --no-italic \\
    --json

Finding Shape Index:
  Use ppt_get_slide_info.py to list shapes and their indices:
  uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 0 --json
  
  Look for shapes with "has_text": true

Common Cross-Platform Fonts:
  - Calibri (default Microsoft Office)
  - Arial (universal)
  - Times New Roman (classic serif)
  - Verdana (screen-optimized)

Accessible Color Palette:
  High Contrast (WCAG AAA - 7:1):
  - #000000 (Black)
  - #333333 (Dark Charcoal)
  - #003366 (Navy Blue)
  
  Good Contrast (WCAG AA - 4.5:1):
  - #595959 (Dark Gray)
  - #0070C0 (Corporate Blue)
  - #006400 (Forest Green)

Accessibility Guidelines:
  - Minimum font size: 12pt (14pt for presentations)
  - Color contrast: 4.5:1 for normal text, 3:1 for large text (≥18pt)
  - Tool automatically validates and warns about issues

Output Format:
  {
    "status": "warning",
    "slide_index": 0,
    "shape_index": 2,
    "before": {"font_size": 24},
    "after": {"font_size": 11, "color": "#CCCCCC"},
    "changes_applied": ["font_size", "color"],
    "validation": {
      "font_size": 11,
      "font_size_ok": false,
      "color_contrast": {"ratio": 2.1, "wcag_aa": false}
    },
    "warnings": ["Font size 11pt is below minimum..."],
    "recommendations": ["Use 12pt minimum..."],
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--shape',
        required=True,
        type=int,
        help='Shape index (0-based, use ppt_get_slide_info.py to find)'
    )
    
    parser.add_argument(
        '--font-name',
        help='Font name (e.g., Arial, Calibri)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        help='Font size in points (minimum recommended: 12pt)'
    )
    
    parser.add_argument(
        '--color',
        help='Text color hex (e.g., #0070C0). Contrast will be validated.'
    )
    
    parser.add_argument(
        '--bold',
        action='store_true',
        dest='bold',
        help='Make text bold'
    )
    
    parser.add_argument(
        '--no-bold',
        action='store_false',
        dest='bold',
        help='Remove bold formatting'
    )
    
    parser.add_argument(
        '--italic',
        action='store_true',
        dest='italic',
        help='Make text italic'
    )
    
    parser.add_argument(
        '--no-italic',
        action='store_false',
        dest='italic',
        help='Remove italic formatting'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.set_defaults(bold=None, italic=None)
    
    args = parser.parse_args()
    
    try:
        result = format_text(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            font_name=args.font_name,
            font_size=args.font_size,
            color=args.color,
            bold=args.bold,
            italic=args.italic
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Provide at least one formatting option and ensure shape has text"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None,
            "slide_index": args.slide if hasattr(args, 'slide') else None,
            "shape_index": args.shape if hasattr(args, 'shape') else None,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_get_info.py
```py
#!/usr/bin/env python3
"""
PowerPoint Get Info Tool v3.1.0
Get presentation metadata (slide count, dimensions, file size, version)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_get_info.py --file presentation.pptx --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError
)

__version__ = "3.1.0"


def get_info(filepath: Path) -> Dict[str, Any]:
    """
    Get comprehensive information about a PowerPoint presentation.
    
    This is a read-only operation that does not modify the file.
    It acquires no lock, allowing concurrent reads.
    
    Args:
        filepath: Path to the PowerPoint file to inspect
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to the file
            - slide_count: Number of slides
            - file_size_bytes: File size in bytes
            - file_size_mb: File size in megabytes (rounded to 2 decimals)
            - slide_dimensions: Width, height (inches), and aspect ratio
            - layouts: List of available layout names
            - layout_count: Number of available layouts
            - modified: Last modification timestamp (if available)
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        PowerPointAgentError: If the file cannot be read
        
    Example:
        >>> result = get_info(Path("presentation.pptx"))
        >>> print(result["slide_count"])
        15
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        # Open without acquiring lock (read-only operation)
        agent.open(filepath, acquire_lock=False)
        
        # Get comprehensive presentation info
        info = agent.get_presentation_info()
    
    # Build response with all available information
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_count": info.get("slide_count", 0),
        "file_size_bytes": info.get("file_size_bytes", 0),
        "file_size_mb": round(info.get("file_size_bytes", 0) / (1024 * 1024), 2),
        "slide_dimensions": {
            "width_inches": info.get("slide_width_inches", 13.333),
            "height_inches": info.get("slide_height_inches", 7.5),
            "aspect_ratio": info.get("aspect_ratio", "16:9")
        },
        "layouts": info.get("layouts", []),
        "layout_count": len(info.get("layouts", [])),
        "modified": info.get("modified"),
        "presentation_version": info.get("presentation_version"),
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Get PowerPoint presentation information",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Get presentation info
  uv run tools/ppt_get_info.py --file presentation.pptx --json
  
  # Check before making modifications
  uv run tools/ppt_get_info.py --file deck.pptx --json | jq '.slide_count'

Output Information:
  - file: Absolute path to the file
  - slide_count: Total number of slides
  - file_size_bytes/mb: File size
  - slide_dimensions: Width, height (inches), and aspect ratio
  - layouts: List of available layout names
  - layout_count: Number of available layouts
  - modified: Last modification timestamp
  - presentation_version: State hash for change tracking

Use Cases:
  - Verify presentation structure before editing
  - Check aspect ratio for compatibility
  - List available layouts for slide creation
  - Track presentation state via version hash
  - Validate file size limits

Aspect Ratios:
  - 16:9 (Widescreen): Most common, modern standard
  - 4:3 (Standard): Traditional, older format
  - 16:10: Some displays, between 16:9 and 4:3

Layout Information:
  The layouts list shows all slide layouts available in the presentation.
  Use these exact names with other tools:
  - ppt_create_new.py --layout "Title Slide"
  - ppt_add_slide.py --layout "Title and Content"
  - ppt_set_slide_layout.py --layout "Section Header"

Version Tracking:
  The presentation_version field is a hash of the presentation state
  including slide count, layouts, shape geometry, and text content.
  Use this to detect changes between operations.

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_count": 15,
    "file_size_bytes": 2568192,
    "file_size_mb": 2.45,
    "slide_dimensions": {
      "width_inches": 13.333,
      "height_inches": 7.5,
      "aspect_ratio": "16:9"
    },
    "layouts": [
      "Title Slide",
      "Title and Content",
      "Section Header",
      "Two Content",
      "Comparison",
      "Title Only",
      "Blank"
    ],
    "layout_count": 7,
    "modified": "2024-01-15T10:30:00",
    "presentation_version": "a1b2c3d4e5f6g7h8",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = get_info(filepath=args.file)
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_get_slide_info.py
```py
#!/usr/bin/env python3
"""
PowerPoint Get Slide Info Tool v3.1.0
Get detailed information about slide content (shapes, images, text, positions)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Features:
    - Full text content (no truncation)
    - Position information (inches and percentages)
    - Size information (inches and percentages)
    - Human-readable placeholder type names
    - Notes detection

Usage:
    uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 0 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Use Cases:
    - Finding shape indices for ppt_format_text.py
    - Locating images for ppt_replace_image.py
    - Debugging positioning issues
    - Auditing slide content
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"


def get_slide_info(
    filepath: Path,
    slide_index: int
) -> Dict[str, Any]:
    """
    Get detailed slide information including full text and positioning.
    
    This is a read-only operation that does not modify the file.
    It acquires no lock, allowing concurrent reads.
    
    Args:
        filepath: Path to the PowerPoint file
        slide_index: Slide index (0-based)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to file
            - slide_index: Index of the slide
            - layout: Layout name
            - shape_count: Total number of shapes
            - shapes: List of shape information dicts
            - has_notes: Whether slide has speaker notes
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Each shape dict contains:
        - index: Shape index (for targeting with other tools)
        - type: Shape type (with human-readable placeholder names)
        - name: Shape name
        - has_text: Boolean
        - text: Full text content (no truncation)
        - text_length: Character count
        - text_preview: First 100 chars (if text > 100 chars)
        - position: Dict with inches and percentages
        - size: Dict with inches and percentages
        - is_placeholder: Boolean
        - placeholder_type: Human-readable type (if placeholder)
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        
    Example:
        >>> result = get_slide_info(Path("presentation.pptx"), 0)
        >>> print(result["shape_count"])
        5
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        # Open without acquiring lock (read-only operation)
        agent.open(filepath, acquire_lock=False)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Get enhanced slide info from core
        slide_info = agent.get_slide_info(slide_index)
        
        # Get presentation version
        prs_info = agent.get_presentation_info()
        presentation_version = prs_info.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_info.get("slide_index", slide_index),
        "layout": slide_info.get("layout", "Unknown"),
        "shape_count": slide_info.get("shape_count", 0),
        "shapes": slide_info.get("shapes", []),
        "has_notes": slide_info.get("has_notes", False),
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Get PowerPoint slide information with full text and positioning",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Get info for first slide
  uv run tools/ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --json
  
  # Get info for specific slide
  uv run tools/ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --json
  
  # Find text shapes
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json | \\
    jq '.shapes[] | select(.has_text == true)'
  
  # Find footer elements (shapes at bottom)
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json | \\
    jq '.shapes[] | select(.type | contains("FOOTER"))'

Output Information:
  - Slide layout name
  - Total shape count
  - Presentation version (for change tracking)
  - List of all shapes with:
    - Shape index (for targeting with other tools)
    - Shape type with human-readable placeholder names
    - Shape name
    - Whether it contains text
    - FULL text content (no truncation)
    - Position in inches and percentages
    - Size in inches and percentages

Use Cases:
  - Find shape indices for ppt_format_text.py
  - Locate images for ppt_replace_image.py
  - Inspect slide layout and structure
  - Audit slide content
  - Debug positioning issues
  - Verify footer/header presence

Finding Shape Indices:
  Use this tool before:
  - ppt_format_text.py (needs shape index)
  - ppt_replace_image.py (needs image name)
  - ppt_format_shape.py (needs shape index)
  - ppt_set_image_properties.py (needs shape index)
  - ppt_crop_image.py (needs shape index)

Example Output:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "layout": "Title Slide",
    "shape_count": 5,
    "shapes": [
      {
        "index": 0,
        "type": "PLACEHOLDER (TITLE)",
        "name": "Title 1",
        "has_text": true,
        "text": "My Presentation Title",
        "text_length": 21,
        "position": {
          "left_inches": 0.5,
          "top_inches": 1.0,
          "left_percent": "5.0%",
          "top_percent": "13.3%"
        },
        "size": {
          "width_inches": 9.0,
          "height_inches": 1.5,
          "width_percent": "90.0%",
          "height_percent": "20.0%"
        }
      }
    ],
    "has_notes": false,
    "presentation_version": "a1b2c3d4e5f6g7h8",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = get_slide_info(
            filepath=args.file,
            slide_index=args.slide
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None,
            "slide_index": args.slide if hasattr(args, 'slide') else None,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_insert_image.py
```py
#!/usr/bin/env python3
"""
PowerPoint Insert Image Tool v3.1.0
Insert image into slide with automatic aspect ratio handling

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_insert_image.py --file presentation.pptx --slide 0 --image logo.png --position '{"left":"10%","top":"10%"}' --size '{"width":"20%","height":"auto"}' --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Accessibility:
    Always use --alt-text to provide alternative text for screen readers.
    This is required for WCAG 2.1 compliance.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Define fallback exceptions if not available in core
try:
    from core.powerpoint_agent_core import ImageNotFoundError
except ImportError:
    class ImageNotFoundError(PowerPointAgentError):
        """Exception raised when image file is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)

try:
    from core.powerpoint_agent_core import InvalidPositionError
except ImportError:
    class InvalidPositionError(PowerPointAgentError):
        """Exception raised when position specification is invalid."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def insert_image(
    filepath: Path,
    slide_index: int,
    image_path: Path,
    position: Dict[str, Any],
    size: Optional[Dict[str, Any]] = None,
    compress: bool = False,
    alt_text: Optional[str] = None
) -> Dict[str, Any]:
    """
    Insert an image into a PowerPoint slide.
    
    Supports automatic aspect ratio preservation when using "auto" for
    width or height. Optionally compresses large images and sets
    accessibility alt text.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the target slide (0-based)
        image_path: Path to the image file to insert
        position: Position specification dict (percentage, anchor, or grid-based)
        size: Size specification dict (optional, defaults to 50% width with auto height)
        compress: Whether to compress the image before insertion (default: False)
        alt_text: Alternative text for accessibility (highly recommended)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the inserted image shape
            - image_file: Path to the source image
            - image_size_bytes: Original image file size
            - image_size_mb: Original image size in MB
            - position: Applied position
            - size: Applied size
            - compressed: Whether compression was applied
            - alt_text: Applied alt text (or None)
            - presentation_version_before: State hash before insertion
            - presentation_version_after: State hash after insertion
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint or image file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        ValueError: If image format is unsupported
        InvalidPositionError: If position specification is invalid
        
    Example:
        >>> result = insert_image(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     image_path=Path("logo.png"),
        ...     position={"left": "10%", "top": "10%"},
        ...     size={"width": "20%", "height": "auto"},
        ...     alt_text="Company Logo"
        ... )
        >>> print(result["shape_index"])
        5
    """
    # Validate presentation file exists
    if not filepath.exists():
        raise FileNotFoundError(f"Presentation file not found: {filepath}")
    
    # Validate image file exists
    if not image_path.exists():
        raise ImageNotFoundError(
            f"Image file not found: {image_path}",
            details={"image_path": str(image_path)}
        )
    
    # Validate image format
    valid_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif'}
    if image_path.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Unsupported image format: {image_path.suffix}. "
            f"Supported formats: {', '.join(sorted(valid_extensions))}"
        )
    
    # Default size if not provided
    if size is None:
        size = {"width": "50%", "height": "auto"}
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE insertion
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Insert image
        result = agent.insert_image(
            slide_index=slide_index,
            image_path=image_path,
            position=position,
            size=size,
            compress=compress
        )
        
        # Extract shape index from result (handle both v3.0.x and v3.1.x)
        if isinstance(result, dict):
            shape_index = result.get("shape_index")
        else:
            # Fallback: get last shape index from slide info
            slide_info = agent.get_slide_info(slide_index)
            shape_index = slide_info["shape_count"] - 1
        
        # Set alt text if provided
        if alt_text and shape_index is not None:
            agent.set_image_properties(
                slide_index=slide_index,
                shape_index=shape_index,
                alt_text=alt_text
            )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER insertion
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
        
        # Get final slide info
        final_slide_info = agent.get_slide_info(slide_index)
    
    # Get image file info
    image_size = image_path.stat().st_size
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "image_file": str(image_path.resolve()),
        "image_size_bytes": image_size,
        "image_size_mb": round(image_size / (1024 * 1024), 2),
        "position": position,
        "size": size,
        "compressed": compress,
        "alt_text": alt_text,
        "slide_shape_count": final_slide_info.get("shape_count", 0),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Insert image into PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Insert logo with alt text (accessibility)
  uv run tools/ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --image company_logo.png \\
    --position '{"left":"5%","top":"5%"}' \\
    --size '{"width":"15%","height":"auto"}' \\
    --alt-text "Company Logo" \\
    --json
  
  # Insert centered hero image with compression
  uv run tools/ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --image product_photo.jpg \\
    --position '{"anchor":"center","offset_x":0,"offset_y":0}' \\
    --size '{"width":"80%","height":"auto"}' \\
    --compress \\
    --alt-text "Product photograph showing new design" \\
    --json
  
  # Insert chart with grid positioning
  uv run tools/ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --image revenue_chart.png \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"auto"}' \\
    --alt-text "Revenue growth chart: Q1 $100K, Q2 $150K, Q3 $200K, Q4 $250K" \\
    --json

Size Options:
  {"width": "50%", "height": "auto"}  - Auto-calculate height (recommended)
  {"width": "auto", "height": "40%"}  - Auto-calculate width
  {"width": "30%", "height": "20%"}   - Fixed dimensions
  {"width": 3.0, "height": 2.0}       - Absolute inches

Position Options:
  {"left": "10%", "top": "20%"}       - Percentage of slide
  {"anchor": "center"}                - Anchor-based
  {"left": 1.5, "top": 2.0}           - Absolute inches

Supported Formats:
  - PNG (recommended for logos, diagrams, transparency)
  - JPG/JPEG (recommended for photos)
  - GIF (first frame only, animation not supported)
  - BMP, TIFF (not recommended, large file size)

Compression (--compress):
  - Resizes to max 1920px width
  - Converts RGBA to RGB
  - JPEG quality 85%
  - Typically reduces size 50-70%

Accessibility (--alt-text):
  - REQUIRED for WCAG 2.1 compliance
  - Describe the image content and purpose
  - For charts/data: include key data points
  - For decorative images: use empty string ""

Best Practices:
  - Always use --alt-text for accessibility
  - Use "auto" for height OR width to maintain aspect ratio
  - Use --compress for images > 1MB
  - Recommended max resolution: 1920x1080
  - Use PNG for transparency, JPG for photos

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "shape_index": 5,
    "image_file": "/path/to/logo.png",
    "image_size_mb": 0.25,
    "alt_text": "Company Logo",
    "presentation_version_after": "a1b2c3d4...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--image',
        required=True,
        type=Path,
        help='Image file path'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=str,
        help='Position dict as JSON string'
    )
    
    parser.add_argument(
        '--size',
        type=str,
        default=None,
        help='Size dict as JSON string (default: 50%% width with auto height)'
    )
    
    parser.add_argument(
        '--compress',
        action='store_true',
        help='Compress image before inserting (recommended for large images)'
    )
    
    parser.add_argument(
        '--alt-text',
        help='Alternative text for accessibility (highly recommended)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse JSON arguments
        try:
            position = json.loads(args.position)
        except json.JSONDecodeError as e:
            raise ValueError(
                f"Invalid JSON in --position: {e}. "
                "Use single quotes around JSON: '{\"left\":\"10%\",\"top\":\"20%\"}'"
            )
        
        size = None
        if args.size:
            try:
                size = json.loads(args.size)
            except json.JSONDecodeError as e:
                raise ValueError(
                    f"Invalid JSON in --size: {e}. "
                    "Use single quotes around JSON: '{\"width\":\"50%\",\"height\":\"auto\"}'"
                )
        
        result = insert_image(
            filepath=args.file,
            slide_index=args.slide,
            image_path=args.image,
            position=position,
            size=size,
            compress=args.compress,
            alt_text=args.alt_text
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except (FileNotFoundError, ImageNotFoundError) as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Verify file paths exist and are accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check JSON format and image file format"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_json_adapter.py
```py
#!/usr/bin/env python3
"""
PowerPoint JSON Adapter Tool v3.1.1
Validates and normalizes JSON outputs from presentation CLI tools.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_json_adapter.py --schema ppt_get_info.schema.json --input raw.json

Behavior:
    - Validates input JSON against provided schema
    - Maps common alias keys to canonical keys
    - Emits normalized JSON to stdout
    - On validation failure, emits structured error JSON and exits non-zero

Exit Codes:
    0: Success (valid and normalized)
    2: Validation Error (schema validation failed)
    3: Input Load Error (could not read input file)
    5: Schema Load Error (could not read schema file)

Changelog v3.1.1:
    - Added hygiene block for JSON pipeline safety
    - Fixed ERROR_TEMPLATE bug causing duplicate keys
    - Added status wrapper to success output
    - Added tool_version to all outputs
    - Improved error response format
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that JSON parsers only see valid JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import argparse
import json
import hashlib
from typing import Dict, Any, Optional, List, Union
from pathlib import Path

try:
    from jsonschema import validate, ValidationError
except ImportError:
    validate = None
    ValidationError = Exception

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"

# Alias mapping table for common drifted/variant keys across tool versions
ALIAS_MAP = {
    # Slide count variants
    "slides_count": "slide_count",
    "slidesTotal": "slide_count",
    "num_slides": "slide_count",
    "total_slides": "slide_count",
    
    # Slides list variants
    "slides_list": "slides",
    "slidesList": "slides",
    
    # Probe variants
    "probe_time": "probe_timestamp",
    "probeTime": "probe_timestamp",
    "probed_at": "probe_timestamp",
    
    # Permission variants
    "canWrite": "can_write",
    "writeable": "can_write",
    "canRead": "can_read",
    "readable": "can_read",
    
    # Size variants
    "maxImageSizeMB": "max_image_size_mb",
    "max_image_size": "max_image_size_mb",
    
    # Version variants
    "version": "presentation_version",
    "pres_version": "presentation_version",
    "file_version": "presentation_version"
}


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def emit_error(
    error_code: str,
    message: str,
    details: Optional[Any] = None,
    retryable: bool = False
) -> None:
    """
    Emit a standardized error response to stdout.
    
    Args:
        error_code: Machine-readable error code
        message: Human-readable error message
        details: Additional error details
        retryable: Whether the operation can be retried
    """
    error_response = {
        "status": "error",
        "tool_version": __version__,
        "error": {
            "error_code": error_code,
            "message": message,
            "details": details,
            "retryable": retryable
        }
    }
    sys.stdout.write(json.dumps(error_response, indent=2) + "\n")
    sys.stdout.flush()


def load_json(path: Path) -> Dict[str, Any]:
    """
    Load JSON from a file path.
    
    Args:
        path: Path to JSON file
        
    Returns:
        Parsed JSON as dictionary
        
    Raises:
        FileNotFoundError: If file doesn't exist
        json.JSONDecodeError: If file contains invalid JSON
    """
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def map_aliases(obj: Any) -> Any:
    """
    Recursively map aliased keys to their canonical forms.
    
    Args:
        obj: Object to process (dict, list, or primitive)
        
    Returns:
        Object with aliased keys replaced by canonical keys
    """
    if isinstance(obj, dict):
        new_dict = {}
        for key, value in obj.items():
            canonical_key = ALIAS_MAP.get(key, key)
            if isinstance(value, dict):
                new_dict[canonical_key] = map_aliases(value)
            elif isinstance(value, list):
                new_dict[canonical_key] = [map_aliases(item) for item in value]
            else:
                new_dict[canonical_key] = value
        return new_dict
    elif isinstance(obj, list):
        return [map_aliases(item) for item in obj]
    else:
        return obj


def compute_presentation_version(info_obj: Dict[str, Any]) -> Optional[str]:
    """
    Compute a best-effort presentation_version if missing.
    
    This is a fallback approximation when the actual version from
    PowerPointAgent is unavailable. It uses available metadata to
    produce a deterministic hash.
    
    NOTE: This does NOT include shape geometry (left:top:width:height)
    as specified in the Core Handbook. It is only used when actual
    version tracking data is missing from the input.
    
    Args:
        info_obj: Presentation info dictionary
        
    Returns:
        SHA-256 hash string (first 16 chars) or None if computation fails
    """
    try:
        slides = info_obj.get("slides", [])
        
        slide_identifiers = []
        for slide in slides:
            slide_id = slide.get("id", slide.get("index", slide.get("slide_index", "")))
            slide_identifiers.append(str(slide_id))
        
        slide_ids_str = ",".join(slide_identifiers)
        
        file_path = info_obj.get("file", info_obj.get("filepath", ""))
        slide_count = info_obj.get("slide_count", len(slides))
        
        hash_input = f"{file_path}-{slide_count}-{slide_ids_str}"
        
        full_hash = hashlib.sha256(hash_input.encode("utf-8")).hexdigest()
        return full_hash[:16]
        
    except Exception:
        return None


def should_compute_version(schema: Dict[str, Any]) -> bool:
    """
    Determine if this schema type should have a computed version.
    
    Args:
        schema: JSON Schema dictionary
        
    Returns:
        True if presentation_version should be computed when missing
    """
    schema_id = schema.get("$id", "")
    schema_title = schema.get("title", "").lower()
    
    version_relevant_patterns = [
        "ppt_get_info",
        "get_info",
        "presentation_info",
        "ppt_capability_probe",
        "capability_probe"
    ]
    
    for pattern in version_relevant_patterns:
        if pattern in schema_id.lower() or pattern in schema_title:
            return True
    
    required_fields = schema.get("required", [])
    if "presentation_version" in required_fields:
        return True
    
    return False


# ============================================================================
# MAIN LOGIC
# ============================================================================

def adapt_json(
    schema_path: Path,
    input_path: Path
) -> Dict[str, Any]:
    """
    Validate and normalize JSON input against schema.
    
    Args:
        schema_path: Path to JSON Schema file
        input_path: Path to input JSON file
        
    Returns:
        Normalized and validated JSON wrapped in success response
    """
    if validate is None:
        emit_error(
            "DEPENDENCY_ERROR",
            "jsonschema library not installed",
            details={"required_package": "jsonschema"},
            retryable=False
        )
        sys.exit(5)
    
    try:
        schema = load_json(schema_path)
    except FileNotFoundError:
        emit_error(
            "SCHEMA_NOT_FOUND",
            f"Schema file not found: {schema_path}",
            details={"path": str(schema_path)},
            retryable=False
        )
        sys.exit(5)
    except json.JSONDecodeError as e:
        emit_error(
            "SCHEMA_PARSE_ERROR",
            f"Invalid JSON in schema file: {e.msg}",
            details={"path": str(schema_path), "line": e.lineno, "column": e.colno},
            retryable=False
        )
        sys.exit(5)
    except Exception as e:
        emit_error(
            "SCHEMA_LOAD_ERROR",
            str(e),
            details={"path": str(schema_path)},
            retryable=False
        )
        sys.exit(5)
    
    try:
        raw_input = load_json(input_path)
    except FileNotFoundError:
        emit_error(
            "INPUT_NOT_FOUND",
            f"Input file not found: {input_path}",
            details={"path": str(input_path)},
            retryable=True
        )
        sys.exit(3)
    except json.JSONDecodeError as e:
        emit_error(
            "INPUT_PARSE_ERROR",
            f"Invalid JSON in input file: {e.msg}",
            details={"path": str(input_path), "line": e.lineno, "column": e.colno},
            retryable=True
        )
        sys.exit(3)
    except Exception as e:
        emit_error(
            "INPUT_LOAD_ERROR",
            str(e),
            details={"path": str(input_path)},
            retryable=True
        )
        sys.exit(3)
    
    normalized = map_aliases(raw_input)
    
    if "presentation_version" not in normalized:
        if should_compute_version(schema):
            computed_version = compute_presentation_version(normalized)
            if computed_version:
                normalized["presentation_version"] = computed_version
                normalized["_version_computed"] = True
    
    try:
        validate(instance=normalized, schema=schema)
    except ValidationError as ve:
        schema_path_str = list(ve.schema_path) if ve.schema_path else None
        emit_error(
            "SCHEMA_VALIDATION_ERROR",
            ve.message,
            details={
                "schema_path": schema_path_str,
                "validator": ve.validator,
                "validator_value": str(ve.validator_value) if ve.validator_value else None,
                "instance_path": list(ve.absolute_path) if ve.absolute_path else None
            },
            retryable=False
        )
        sys.exit(2)
    
    return {
        "status": "success",
        "tool_version": __version__,
        "schema_used": str(schema_path),
        "input_file": str(input_path),
        "aliases_mapped": _count_mapped_aliases(raw_input, normalized),
        "data": normalized
    }


def _count_mapped_aliases(original: Any, normalized: Any) -> int:
    """
    Count how many aliases were mapped during normalization.
    
    Args:
        original: Original input object
        normalized: Normalized object
        
    Returns:
        Number of keys that were remapped
    """
    count = 0
    
    if isinstance(original, dict) and isinstance(normalized, dict):
        for key in original:
            if key in ALIAS_MAP:
                count += 1
            if key in original and isinstance(original[key], (dict, list)):
                canonical = ALIAS_MAP.get(key, key)
                if canonical in normalized:
                    count += _count_mapped_aliases(original[key], normalized[canonical])
    elif isinstance(original, list) and isinstance(normalized, list):
        for orig_item, norm_item in zip(original, normalized):
            count += _count_mapped_aliases(orig_item, norm_item)
    
    return count


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Validate and normalize JSON outputs from presentation CLI tools",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Validate and normalize tool output
  uv run tools/ppt_json_adapter.py \\
    --schema schemas/ppt_get_info.schema.json \\
    --input raw_output.json

  # Pipeline usage
  uv run tools/ppt_get_info.py --file deck.pptx --json > raw.json
  uv run tools/ppt_json_adapter.py --schema schemas/ppt_get_info.schema.json --input raw.json

Exit Codes:
  0: Success - valid JSON emitted
  2: Validation Error - input doesn't match schema
  3: Input Load Error - couldn't read input file
  5: Schema Load Error - couldn't read schema file

Alias Mapping:
  The adapter normalizes common key variations:
  - slides_count -> slide_count
  - slidesTotal -> slide_count
  - probe_time -> probe_timestamp
  - canWrite -> can_write
  etc.
        """
    )
    
    parser.add_argument(
        "--schema",
        required=True,
        type=Path,
        help="Path to JSON Schema file"
    )
    
    parser.add_argument(
        "--input",
        required=True,
        type=Path,
        help="Path to raw JSON input file"
    )
    
    args = parser.parse_args()
    
    result = adapt_json(
        schema_path=args.schema,
        input_path=args.input
    )
    
    sys.stdout.write(json.dumps(result, indent=2) + "\n")
    sys.stdout.flush()
    sys.exit(0)


if __name__ == "__main__":
    main()

```

# tools/ppt_merge_presentations.py
```py
#!/usr/bin/env python3
"""
PowerPoint Merge Presentations Tool v3.1.1
Combine slides from multiple presentations into one.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_merge_presentations.py --sources '[{"file":"a.pptx","slides":"all"},{"file":"b.pptx","slides":[0,2,4]}]' --output merged.pptx --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

This tool merges slides from multiple source presentations into a single output
presentation. You can specify which slides to include from each source.

Source Specification Format:
    [
        {"file": "path/to/first.pptx", "slides": "all"},
        {"file": "path/to/second.pptx", "slides": [0, 1, 2]},
        {"file": "path/to/third.pptx", "slides": [5, 6]}
    ]
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import shutil
from pathlib import Path
from typing import Dict, Any, List, Optional, Union

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# TYPE DEFINITIONS
# ============================================================================

SourceSpec = Dict[str, Union[str, List[int]]]


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def parse_sources(sources_json: str) -> List[SourceSpec]:
    """
    Parse and validate sources JSON specification.
    
    Args:
        sources_json: JSON string with source specifications
        
    Returns:
        List of validated source specifications
        
    Raises:
        ValueError: If JSON is invalid or missing required fields
    """
    try:
        sources = json.loads(sources_json)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON in sources: {e}")
    
    if not isinstance(sources, list):
        raise ValueError("Sources must be a JSON array")
    
    if len(sources) == 0:
        raise ValueError("At least one source is required")
    
    validated = []
    for idx, source in enumerate(sources):
        if not isinstance(source, dict):
            raise ValueError(f"Source {idx} must be an object")
        
        if "file" not in source:
            raise ValueError(f"Source {idx} missing required 'file' field")
        
        if "slides" not in source:
            source["slides"] = "all"
        
        slides = source["slides"]
        if slides != "all" and not isinstance(slides, list):
            raise ValueError(f"Source {idx} 'slides' must be 'all' or array of indices")
        
        if isinstance(slides, list):
            for slide_idx in slides:
                if not isinstance(slide_idx, int) or slide_idx < 0:
                    raise ValueError(f"Source {idx} has invalid slide index: {slide_idx}")
        
        validated.append(source)
    
    return validated


def validate_source_files(sources: List[SourceSpec]) -> None:
    """
    Validate all source files exist.
    
    Args:
        sources: List of source specifications
        
    Raises:
        FileNotFoundError: If any source file doesn't exist
    """
    for source in sources:
        filepath = Path(source["file"])
        if not filepath.exists():
            raise FileNotFoundError(f"Source file not found: {filepath}")
        if not filepath.suffix.lower() == '.pptx':
            raise ValueError(f"Source file must be .pptx: {filepath}")


# ============================================================================
# MAIN LOGIC
# ============================================================================

def merge_presentations(
    sources: List[SourceSpec],
    output: Path,
    base_template: Optional[Path] = None,
    preserve_formatting: bool = True
) -> Dict[str, Any]:
    """
    Merge slides from multiple presentations into one.
    
    Args:
        sources: List of source specifications with file paths and slide indices
        output: Path for the output merged presentation
        base_template: Optional template to use for theme/masters
        preserve_formatting: Whether to preserve original slide formatting
        
    Returns:
        Dict with merge results
        
    Raises:
        FileNotFoundError: If source files don't exist
        SlideNotFoundError: If specified slide indices are invalid
        ValueError: If sources specification is invalid
    """
    validate_source_files(sources)
    
    if base_template:
        if not base_template.exists():
            raise FileNotFoundError(f"Base template not found: {base_template}")
        shutil.copy2(base_template, output)
        initial_file = base_template
    else:
        first_source = Path(sources[0]["file"])
        shutil.copy2(first_source, output)
        initial_file = first_source
    
    warnings: List[str] = []
    sources_used: List[Dict[str, Any]] = []
    merge_details: Dict[str, int] = {}
    total_slides_copied = 0
    
    with PowerPointAgent(output) as agent:
        agent.open(output)
        
        if not base_template:
            first_source_info = {
                "file": str(Path(sources[0]["file"]).resolve()),
                "slides_spec": sources[0]["slides"],
                "slides_copied": agent.get_slide_count(),
                "is_base": True
            }
            sources_used.append(first_source_info)
            merge_details[str(Path(sources[0]["file"]).name)] = agent.get_slide_count()
            total_slides_copied += agent.get_slide_count()
            sources_to_process = sources[1:]
        else:
            sources_to_process = sources
        
        from pptx import Presentation
        
        for source_idx, source in enumerate(sources_to_process):
            source_path = Path(source["file"])
            slides_spec = source["slides"]
            
            try:
                source_prs = Presentation(str(source_path))
                source_slide_count = len(source_prs.slides)
                
                if slides_spec == "all":
                    slide_indices = list(range(source_slide_count))
                else:
                    slide_indices = slides_spec
                    for idx in slide_indices:
                        if idx >= source_slide_count:
                            raise SlideNotFoundError(
                                f"Slide {idx} not found in {source_path.name} (has {source_slide_count} slides)",
                                details={
                                    "source_file": str(source_path),
                                    "requested_index": idx,
                                    "available_slides": source_slide_count
                                }
                            )
                
                slides_copied = 0
                for slide_idx in slide_indices:
                    try:
                        source_slide = source_prs.slides[slide_idx]
                        
                        blank_layout = None
                        for layout in agent.prs.slide_layouts:
                            if "blank" in layout.name.lower():
                                blank_layout = layout
                                break
                        if blank_layout is None:
                            blank_layout = agent.prs.slide_layouts[0]
                        
                        new_slide = agent.prs.slides.add_slide(blank_layout)
                        
                        for shape in source_slide.shapes:
                            if shape.shape_type == 13:
                                continue
                            
                            try:
                                el = shape.element
                                new_slide.shapes._spTree.insert_element_before(
                                    el, 'p:extLst'
                                )
                            except Exception:
                                pass
                        
                        slides_copied += 1
                        total_slides_copied += 1
                        
                    except Exception as e:
                        warnings.append(f"Could not copy slide {slide_idx} from {source_path.name}: {str(e)}")
                
                sources_used.append({
                    "file": str(source_path.resolve()),
                    "slides_spec": slides_spec,
                    "slides_copied": slides_copied,
                    "is_base": False
                })
                merge_details[source_path.name] = slides_copied
                
            except SlideNotFoundError:
                raise
            except Exception as e:
                warnings.append(f"Error processing {source_path.name}: {str(e)}")
        
        agent.save()
        
        info = agent.get_presentation_info()
        presentation_version = info.get("presentation_version")
        final_slide_count = info.get("slide_count")
    
    return {
        "status": "success",
        "file": str(output.resolve()),
        "sources_used": sources_used,
        "total_slides": final_slide_count,
        "merge_details": merge_details,
        "base_template": str(base_template.resolve()) if base_template else None,
        "preserve_formatting": preserve_formatting,
        "warnings": warnings,
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Merge slides from multiple PowerPoint presentations",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Merge all slides from two presentations
  uv run tools/ppt_merge_presentations.py \\
    --sources '[{"file":"part1.pptx","slides":"all"},{"file":"part2.pptx","slides":"all"}]' \\
    --output merged.pptx --json

  # Select specific slides from each source
  uv run tools/ppt_merge_presentations.py \\
    --sources '[{"file":"intro.pptx","slides":[0,1]},{"file":"content.pptx","slides":[2,3,4]},{"file":"outro.pptx","slides":[0]}]' \\
    --output presentation.pptx --json

  # Use a template for consistent theming
  uv run tools/ppt_merge_presentations.py \\
    --sources '[{"file":"content1.pptx","slides":"all"},{"file":"content2.pptx","slides":"all"}]' \\
    --output merged.pptx --base-template corporate_template.pptx --json

Source Specification Format:
  The --sources argument must be a JSON array with objects containing:
  - "file": Path to the source .pptx file (required)
  - "slides": Either "all" or an array of slide indices [0, 1, 2] (optional, default: "all")

Behavior:
  - First source becomes the base (its theme/masters are used)
  - Subsequent sources have their slides copied into the base
  - Use --base-template to override with a specific template
  - Slide indices are 0-based

Output Format:
  {
    "status": "success",
    "file": "/path/to/merged.pptx",
    "sources_used": [
      {"file": "part1.pptx", "slides_copied": 5, "is_base": true},
      {"file": "part2.pptx", "slides_copied": 3, "is_base": false}
    ],
    "total_slides": 8,
    "merge_details": {"part1.pptx": 5, "part2.pptx": 3},
    "presentation_version": "a1b2c3...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--sources',
        required=True,
        type=str,
        help='JSON array of source specifications'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output merged presentation path'
    )
    
    parser.add_argument(
        '--base-template',
        type=Path,
        default=None,
        help='Optional template to use for theme/masters'
    )
    
    parser.add_argument(
        '--preserve-formatting',
        action='store_true',
        default=True,
        help='Preserve original slide formatting (default: true)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        sources = parse_sources(args.sources)
        
        output_path = args.output
        if not output_path.suffix.lower() == '.pptx':
            output_path = output_path.with_suffix('.pptx')
        
        result = merge_presentations(
            sources=sources,
            output=output_path.resolve(),
            base_template=args.base_template.resolve() if args.base_template else None,
            preserve_formatting=args.preserve_formatting
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify all source file paths exist and are accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check sources JSON format: [{\"file\":\"path.pptx\",\"slides\":\"all\"}]",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Check slide indices are valid for each source file",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check source file integrity and compatibility",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_remove_shape.py
```py
#!/usr/bin/env python3
"""
PowerPoint Remove Shape Tool v3.1.0
Safely remove shapes from slides with comprehensive safety controls.

⚠️  DESTRUCTIVE OPERATION WARNING ⚠️
- Shape removal CANNOT be undone
- Shape indices WILL shift after removal
- Always CLONE the presentation first
- Always use --dry-run to preview

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    # Preview removal (RECOMMENDED FIRST)
    uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 2 --dry-run --json
    
    # Execute removal
    uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 2 --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
)

__version__ = "3.1.0"


def get_shape_details(agent: PowerPointAgent, slide_index: int, shape_index: int) -> Dict[str, Any]:
    """Get detailed information about a shape before removal."""
    try:
        slide_info = agent.get_slide_info(slide_index)
        shapes = slide_info.get("shapes", [])
        
        if 0 <= shape_index < len(shapes):
            shape = shapes[shape_index]
            return {
                "index": shape_index,
                "type": shape.get("type", "unknown"),
                "name": shape.get("name", ""),
                "has_text": shape.get("has_text", False),
                "text_preview": (shape.get("text", "")[:100] + "...") if len(shape.get("text", "")) > 100 else shape.get("text", ""),
                "position": shape.get("position", {}),
                "size": shape.get("size", {})
            }
    except Exception as e:
        return {"index": shape_index, "error": str(e)}
    
    return {"index": shape_index, "type": "unknown"}


def find_shape_by_name(agent: PowerPointAgent, slide_index: int, name: str) -> Optional[int]:
    """Find shape index by name (partial match)."""
    try:
        slide_info = agent.get_slide_info(slide_index)
        shapes = slide_info.get("shapes", [])
        
        for idx, shape in enumerate(shapes):
            if shape.get("name", "") == name:
                return idx
        
        name_lower = name.lower()
        for idx, shape in enumerate(shapes):
            shape_name = shape.get("name", "").lower()
            if name_lower in shape_name or shape_name in name_lower:
                return idx
        
        return None
    except Exception:
        return None


def remove_shape(
    filepath: Path,
    slide_index: int,
    shape_index: Optional[int] = None,
    shape_name: Optional[str] = None,
    dry_run: bool = False
) -> Dict[str, Any]:
    """
    Remove shape from slide with safety controls.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        shape_index: Shape index to remove (0-based)
        shape_name: Shape name to remove (alternative to index)
        dry_run: If True, preview only without actual removal
        
    Returns:
        Result dict with removal details
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If invalid parameters
        SlideNotFoundError: If slide index invalid
        ShapeNotFoundError: If shape not found
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    if shape_index is None and shape_name is None:
        raise ValueError("Must specify either --shape (index) or --name (shape name)")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        slide_info_before = agent.get_slide_info(slide_index)
        shape_count_before = slide_info_before.get("shape_count", 0)
        
        resolved_index = shape_index
        if shape_name is not None:
            resolved_index = find_shape_by_name(agent, slide_index, shape_name)
            if resolved_index is None:
                raise ShapeNotFoundError(
                    f"Shape with name '{shape_name}' not found on slide {slide_index}",
                    details={
                        "slide_index": slide_index,
                        "shape_name": shape_name,
                        "available_shapes": [s.get("name") for s in slide_info_before.get("shapes", [])]
                    }
                )
        
        if not 0 <= resolved_index < shape_count_before:
            raise ShapeNotFoundError(
                f"Shape index {resolved_index} out of range (0-{shape_count_before - 1})",
                details={"requested": resolved_index, "available": shape_count_before}
            )
        
        shape_details = get_shape_details(agent, slide_index, resolved_index)
        version_before = agent.get_presentation_version()
        
        result: Dict[str, Any] = {
            "file": str(filepath.resolve()),
            "slide_index": slide_index,
            "shape_index": resolved_index,
            "shape_details": shape_details,
            "shape_count_before": shape_count_before,
            "dry_run": dry_run,
            "presentation_version_before": version_before,
            "tool_version": __version__
        }
        
        if dry_run:
            result["status"] = "preview"
            result["message"] = "DRY RUN: Shape would be removed. Run without --dry-run to execute."
            result["shape_count_after"] = shape_count_before - 1
            shapes_affected = shape_count_before - resolved_index - 1
            result["index_shift_info"] = {
                "shapes_affected": shapes_affected,
                "message": f"Shapes at indices {resolved_index + 1} to {shape_count_before - 1} would shift down by 1" if shapes_affected > 0 else "No other shapes would be affected"
            }
        else:
            agent.remove_shape(slide_index=slide_index, shape_index=resolved_index)
            agent.save()
            
            version_after = agent.get_presentation_version()
            slide_info_after = agent.get_slide_info(slide_index)
            shape_count_after = slide_info_after.get("shape_count", 0)
            
            result["status"] = "success"
            result["message"] = "Shape removed successfully"
            result["shape_count_after"] = shape_count_after
            result["presentation_version_after"] = version_after
            
            shapes_shifted = shape_count_before - resolved_index - 1
            if shapes_shifted > 0:
                result["index_shift_info"] = {
                    "shapes_shifted": shapes_shifted,
                    "warning": f"⚠️ {shapes_shifted} shape(s) have new indices. Re-query before further operations.",
                    "refresh_command": f"uv run tools/ppt_get_slide_info.py --file {filepath} --slide {slide_index} --json"
                }
            
            result["rollback_guidance"] = "This operation cannot be undone. Restore from backup clone."
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Remove shape from PowerPoint slide ⚠️ DESTRUCTIVE",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️  DESTRUCTIVE OPERATION - READ CAREFULLY ⚠️

This tool PERMANENTLY REMOVES shapes from presentations.
- Shape removal CANNOT be undone
- Shape indices WILL shift after removal
- Always CLONE the presentation first
- Always use --dry-run to preview

SAFE REMOVAL PROTOCOL:

  1. CLONE: ppt_clone_presentation.py --source original.pptx --output work.pptx
  2. INSPECT: ppt_get_slide_info.py --file work.pptx --slide 0 --json
  3. PREVIEW: ppt_remove_shape.py --file work.pptx --slide 0 --shape 2 --dry-run --json
  4. EXECUTE: ppt_remove_shape.py --file work.pptx --slide 0 --shape 2 --json
  5. REFRESH: ppt_get_slide_info.py --file work.pptx --slide 0 --json

EXAMPLES:

  # Preview removal (ALWAYS DO FIRST)
  uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 3 --dry-run --json

  # Remove by index
  uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 3 --json

  # Remove by name
  uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --name "Rectangle 1" --json
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path (.pptx)')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    
    shape_group = parser.add_mutually_exclusive_group(required=True)
    shape_group.add_argument('--shape', type=int, help='Shape index to remove (0-based)')
    shape_group.add_argument('--name', help='Shape name to remove')
    
    parser.add_argument('--dry-run', action='store_true', help='Preview without executing')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON (default: true)')
    
    args = parser.parse_args()
    
    try:
        result = remove_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            shape_name=args.name,
            dry_run=args.dry_run
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shapes."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Specify --shape INDEX or --name NAME, and ensure .pptx format."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_reorder_slides.py
```py
#!/usr/bin/env python3
"""
PowerPoint Reorder Slides Tool v3.1.0
Move a slide from one position to another

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_reorder_slides.py --file presentation.pptx --from-index 3 --to-index 1 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Notes:
    - Indices are 0-based
    - Moving a slide shifts other slides accordingly
    - Original content is preserved during move
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"


def reorder_slides(
    filepath: Path, 
    from_index: int, 
    to_index: int
) -> Dict[str, Any]:
    """
    Move a slide from one position to another.
    
    The slide at from_index is moved to to_index. Other slides
    shift accordingly to accommodate the move.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        from_index: Current position of the slide (0-based)
        to_index: Target position for the slide (0-based)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - moved_from: Original slide position
            - moved_to: New slide position
            - total_slides: Total slide count
            - presentation_version_before: State hash before reorder
            - presentation_version_after: State hash after reorder
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If from_index or to_index is out of range
        
    Example:
        >>> result = reorder_slides(
        ...     filepath=Path("presentation.pptx"),
        ...     from_index=5,
        ...     to_index=1
        ... )
        >>> print(result["moved_to"])
        1
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate indices are different
    if from_index == to_index:
        # Not an error, but no operation needed
        with PowerPointAgent(filepath) as agent:
            agent.open(filepath, acquire_lock=False)
            total = agent.get_slide_count()
            prs_info = agent.get_presentation_info()
        
        return {
            "status": "success",
            "file": str(filepath.resolve()),
            "moved_from": from_index,
            "moved_to": to_index,
            "total_slides": total,
            "note": "Source and target indices are the same. No change made.",
            "presentation_version_before": prs_info.get("presentation_version"),
            "presentation_version_after": prs_info.get("presentation_version"),
            "tool_version": __version__
        }
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE reorder
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate indices
        total = agent.get_slide_count()
        
        if not 0 <= from_index < total:
            raise SlideNotFoundError(
                f"Source index {from_index} out of range (0-{total - 1})",
                details={
                    "requested_index": from_index,
                    "available_slides": total,
                    "parameter": "from_index"
                }
            )
        
        if not 0 <= to_index < total:
            raise SlideNotFoundError(
                f"Target index {to_index} out of range (0-{total - 1})",
                details={
                    "requested_index": to_index,
                    "available_slides": total,
                    "parameter": "to_index"
                }
            )
        
        # Perform reorder
        agent.reorder_slides(from_index, to_index)
        
        # Save changes
        agent.save()
        
        # Capture version AFTER reorder
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "moved_from": from_index,
        "moved_to": to_index,
        "total_slides": total,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Reorder PowerPoint slides",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Move slide from position 3 to position 1
  uv run tools/ppt_reorder_slides.py \\
    --file presentation.pptx \\
    --from-index 3 \\
    --to-index 1 \\
    --json
  
  # Move last slide to beginning
  uv run tools/ppt_reorder_slides.py \\
    --file deck.pptx \\
    --from-index 9 \\
    --to-index 0 \\
    --json
  
  # Move first slide to end
  uv run tools/ppt_reorder_slides.py \\
    --file deck.pptx \\
    --from-index 0 \\
    --to-index 9 \\
    --json

Behavior:
  - Slide at from_index is moved to to_index
  - Other slides shift to accommodate the move
  - All slide content is preserved
  - Indices are 0-based

Finding Slide Count:
  Use ppt_get_info.py to check slide count:
  uv run tools/ppt_get_info.py --file presentation.pptx --json | jq '.slide_count'

Use Cases:
  - Reorganizing presentation flow
  - Moving section headers
  - Reordering topic sequences
  - Placing summary slides

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "moved_from": 3,
    "moved_to": 1,
    "total_slides": 10,
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    parser.add_argument(
        '--from-index', 
        required=True, 
        type=int, 
        help='Current slide index (0-based)'
    )
    parser.add_argument(
        '--to-index', 
        required=True, 
        type=int, 
        help='Target slide index (0-based)'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = reorder_slides(
            filepath=args.file, 
            from_index=args.from_index, 
            to_index=args.to_index
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check slide count"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_replace_image.py
```py
#!/usr/bin/env python3
"""
PowerPoint Replace Image Tool v3.1.0
Replace an existing image with a new one (preserves position and size)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_replace_image.py --file presentation.pptx --slide 0 --old-image "logo" --new-image new_logo.png --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Use Cases:
    - Logo updates during rebranding
    - Product photo updates
    - Chart/diagram refreshes
    - Team photo updates
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Define fallback exception if not available in core
try:
    from core.powerpoint_agent_core import ImageNotFoundError
except ImportError:
    class ImageNotFoundError(PowerPointAgentError):
        """Exception raised when image is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def replace_image(
    filepath: Path,
    slide_index: int,
    old_image: str,
    new_image: Path,
    compress: bool = False
) -> Dict[str, Any]:
    """
    Replace an existing image with a new one.
    
    Searches for an image by name (exact or partial match) and replaces
    it with the new image while preserving the original position and size.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the image (0-based)
        old_image: Name or partial name of the image to replace
        new_image: Path to the new image file
        compress: Whether to compress the new image (default: False)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - old_image: Name/pattern that was searched
            - new_image: Path to the new image
            - new_image_size_bytes: Size of new image file
            - new_image_size_mb: Size in MB
            - compressed: Whether compression was applied
            - replaced: True if replacement succeeded
            - presentation_version_before: State hash before replacement
            - presentation_version_after: State hash after replacement
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If PowerPoint or new image file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ImageNotFoundError: If old image is not found on the slide
        
    Example:
        >>> result = replace_image(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     old_image="company_logo",
        ...     new_image=Path("new_logo.png")
        ... )
        >>> print(result["replaced"])
        True
    """
    # Validate presentation file exists
    if not filepath.exists():
        raise FileNotFoundError(f"Presentation file not found: {filepath}")
    
    # Validate new image file exists
    if not new_image.exists():
        raise ImageNotFoundError(
            f"New image file not found: {new_image}",
            details={"new_image_path": str(new_image)}
        )
    
    # Validate image format
    valid_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif'}
    if new_image.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Unsupported image format: {new_image.suffix}. "
            f"Supported formats: {', '.join(sorted(valid_extensions))}"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE replacement
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Attempt replacement
        replaced = agent.replace_image(
            slide_index=slide_index,
            old_image_name=old_image,
            new_image_path=new_image,
            compress=compress
        )
        
        if not replaced:
            raise ImageNotFoundError(
                f"Image matching '{old_image}' not found on slide {slide_index}. "
                "Use ppt_get_slide_info.py to list available images.",
                details={
                    "search_pattern": old_image,
                    "slide_index": slide_index
                }
            )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER replacement
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Get new image size
    new_size = new_image.stat().st_size
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "old_image": old_image,
        "new_image": str(new_image.resolve()),
        "new_image_size_bytes": new_size,
        "new_image_size_mb": round(new_size / (1024 * 1024), 2),
        "compressed": compress,
        "replaced": True,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Replace image in PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Replace logo by name
  uv run tools/ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --old-image "company_logo" \\
    --new-image new_logo.png \\
    --json
  
  # Replace with compression
  uv run tools/ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --old-image "product_photo" \\
    --new-image updated_photo.jpg \\
    --compress \\
    --json
  
  # Partial name match
  uv run tools/ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --old-image "logo" \\
    --new-image rebrand_logo.png \\
    --json

Finding Images:
  Use ppt_get_slide_info.py to list images on a slide:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Search Strategy:
  The tool searches for images by:
  1. Exact name match
  2. Partial name match (contains)
  3. First match is replaced

Compression (--compress):
  - Resizes to max 1920px width
  - Converts to JPEG at 85% quality
  - Typically reduces size 50-70%
  - Recommended for images > 1MB

Best Practices:
  - Use descriptive image names in PowerPoint
  - Keep new image dimensions similar to original
  - Use --compress for large replacement images
  - Test on a cloned copy first
  - Verify aspect ratios match

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "old_image": "company_logo",
    "new_image": "/path/to/new_logo.png",
    "new_image_size_mb": 0.15,
    "compressed": false,
    "replaced": true,
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--old-image',
        required=True,
        help='Name or partial name of image to replace'
    )
    
    parser.add_argument(
        '--new-image',
        required=True,
        type=Path,
        help='Path to new image file'
    )
    
    parser.add_argument(
        '--compress',
        action='store_true',
        help='Compress new image before inserting'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = replace_image(
            filepath=args.file,
            slide_index=args.slide,
            old_image=args.old_image,
            new_image=args.new_image,
            compress=args.compress
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except (FileNotFoundError, ImageNotFoundError) as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to list available images on the slide"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check image file format (PNG, JPG, GIF, BMP supported)"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_replace_text.py
```py
#!/usr/bin/env python3
"""
PowerPoint Replace Text Tool v3.1.0
Find and replace text across presentation or in specific targets

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Features:
    - Global replacement (entire presentation)
    - Targeted replacement (specific slide)
    - Surgical replacement (specific shape)
    - Dry-run mode (preview without changes)
    - Case-sensitive matching option
    - Formatting-preserving replacement (run-level)
    - Location reporting

Usage:
    uv run tools/ppt_replace_text.py --file deck.pptx --find "Old" --replace "New" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Safety:
    Always use --dry-run first to preview changes before applying.
    For mass replacements, consider cloning the presentation first.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import re
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"


def perform_replacement_on_shape(
    shape, 
    find: str, 
    replace: str, 
    match_case: bool
) -> int:
    """
    Perform text replacement in a single shape.
    
    Uses a two-strategy approach:
    1. Run-level replacement (preserves formatting)
    2. Shape-level fallback (for text split across runs)
    
    Args:
        shape: PowerPoint shape object with text_frame
        find: Text to find
        replace: Replacement text
        match_case: Whether to match case
        
    Returns:
        Number of replacements made
    """
    if not hasattr(shape, 'text_frame'):
        return 0
    
    count = 0
    text_frame = shape.text_frame
    
    # Strategy 1: Replace in runs (preserves formatting)
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if match_case:
                if find in run.text:
                    run.text = run.text.replace(find, replace)
                    count += 1
            else:
                if find.lower() in run.text.lower():
                    pattern = re.compile(re.escape(find), re.IGNORECASE)
                    if pattern.search(run.text):
                        run.text = pattern.sub(replace, run.text)
                        count += 1
    
    if count > 0:
        return count
    
    # Strategy 2: Shape-level replacement (if runs didn't catch it due to splitting)
    try:
        full_text = shape.text
        should_replace = False
        
        if match_case:
            if find in full_text:
                should_replace = True
        else:
            if find.lower() in full_text.lower():
                should_replace = True
        
        if should_replace:
            if match_case:
                new_text = full_text.replace(find, replace)
            else:
                pattern = re.compile(re.escape(find), re.IGNORECASE)
                new_text = pattern.sub(replace, full_text)
            
            # Only apply if text actually changed
            if new_text != full_text:
                shape.text = new_text
                count += 1
    except Exception:
        pass  # Continue without shape-level replacement
    
    return count


def replace_text(
    filepath: Path,
    find: str,
    replace: str,
    slide_index: Optional[int] = None,
    shape_index: Optional[int] = None,
    match_case: bool = False,
    dry_run: bool = False
) -> Dict[str, Any]:
    """
    Find and replace text with optional targeting.
    
    Supports three scopes:
    1. Global: All slides, all shapes (default)
    2. Slide-specific: Single slide, all shapes (--slide N)
    3. Shape-specific: Single shape (--slide N --shape M)
    
    Args:
        filepath: Path to PowerPoint file
        find: Text to find
        replace: Replacement text
        slide_index: Optional specific slide index (0-based)
        shape_index: Optional specific shape index (requires slide_index)
        match_case: Whether to match case (default: False)
        dry_run: Preview without making changes (default: False)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Path to file
            - action: "dry_run" or "replace"
            - find/replace: Search parameters
            - scope: Target scope information
            - total_matches/replacements_made: Count
            - locations: List of affected locations
            - presentation_version_before: State hash before (if not dry_run)
            - presentation_version_after: State hash after (if not dry_run)
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If find is empty or invalid parameters
        SlideNotFoundError: If slide index is out of range
        
    Example:
        >>> result = replace_text(
        ...     filepath=Path("presentation.pptx"),
        ...     find="Old Company",
        ...     replace="New Company",
        ...     dry_run=True
        ... )
        >>> print(result["total_matches"])
        15
    """
    # Validate file extension
    valid_extensions = {'.pptx', '.pptm', '.potx'}
    if filepath.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Invalid PowerPoint file format: {filepath.suffix}. "
            f"Supported formats: {', '.join(valid_extensions)}"
        )
    
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate find text
    if not find:
        raise ValueError("Find text cannot be empty")
    
    # Validate parameters
    if shape_index is not None and slide_index is None:
        raise ValueError(
            "If --shape is specified, --slide must also be specified. "
            "Shape indices are slide-specific."
        )
    
    action = "dry_run" if dry_run else "replace"
    total_count = 0
    locations: List[Dict[str, Any]] = []
    version_before = None
    version_after = None
    
    with PowerPointAgent(filepath) as agent:
        # Open with appropriate locking
        agent.open(filepath, acquire_lock=not dry_run)
        
        # Capture version BEFORE (only for actual replacements)
        if not dry_run:
            info_before = agent.get_presentation_info()
            version_before = info_before.get("presentation_version")
        
        slide_count = agent.get_slide_count()
        
        # Include performance note in response for large presentations
        large_presentation = slide_count > 50
        
        # Determine target slides
        target_slides: List[tuple] = []
        
        if slide_index is not None:
            # Single slide scope
            if not 0 <= slide_index < slide_count:
                raise SlideNotFoundError(
                    f"Slide index {slide_index} out of range (0-{slide_count - 1})",
                    details={
                        "requested_index": slide_index,
                        "available_slides": slide_count
                    }
                )
            # NOTE: Direct prs access required for shape-level text manipulation
            target_slides = [(slide_index, agent.prs.slides[slide_index])]
        else:
            # Global scope
            target_slides = [(i, slide) for i, slide in enumerate(agent.prs.slides)]
        
        # Process each target slide
        for s_idx, slide in target_slides:
            # Determine target shapes
            target_shapes: List[tuple] = []
            
            if shape_index is not None:
                # Single shape scope
                if not 0 <= shape_index < len(slide.shapes):
                    raise ValueError(
                        f"Shape index {shape_index} out of range (0-{len(slide.shapes) - 1}) on slide {s_idx}"
                    )
                target_shapes = [(shape_index, slide.shapes[shape_index])]
            else:
                # All shapes on slide
                target_shapes = [(i, shape) for i, shape in enumerate(slide.shapes)]
            
            # Process each target shape
            for sh_idx, shape in target_shapes:
                if not hasattr(shape, 'text_frame'):
                    continue
                
                if dry_run:
                    # Count occurrences without modifying
                    text = shape.text_frame.text
                    occurrences = 0
                    
                    if match_case:
                        occurrences = text.count(find)
                    else:
                        occurrences = text.lower().count(find.lower())
                    
                    if occurrences > 0:
                        total_count += occurrences
                        preview = text[:100] + "..." if len(text) > 100 else text
                        locations.append({
                            "slide": s_idx,
                            "shape": sh_idx,
                            "occurrences": occurrences,
                            "preview": preview
                        })
                else:
                    # Perform actual replacement
                    replacements = perform_replacement_on_shape(
                        shape, find, replace, match_case
                    )
                    
                    if replacements > 0:
                        total_count += replacements
                        locations.append({
                            "slide": s_idx,
                            "shape": sh_idx,
                            "replacements": replacements
                        })
        
        # Save changes (only for actual replacements)
        if not dry_run:
            agent.save()
            
            # Capture version AFTER
            info_after = agent.get_presentation_info()
            version_after = info_after.get("presentation_version")
    
    # Build result
    result: Dict[str, Any] = {
        "status": "success",
        "file": str(filepath.resolve()),
        "action": action,
        "find": find,
        "replace": replace,
        "match_case": match_case,
        "scope": {
            "slide": slide_index if slide_index is not None else "all",
            "shape": shape_index if shape_index is not None else "all"
        },
        "locations": locations,
        "tool_version": __version__
    }
    
    # Add appropriate count field
    if dry_run:
        result["total_matches"] = total_count
    else:
        result["replacements_made"] = total_count
        result["presentation_version_before"] = version_before
        result["presentation_version_after"] = version_after
    
    # Add performance note for large presentations
    if large_presentation:
        result["note"] = f"Large presentation ({slide_count} slides) processed"
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Find and replace text in PowerPoint",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Preview global replacement (ALWAYS do this first!)
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "Old Company" \\
    --replace "New Company" \\
    --dry-run \\
    --json
  
  # Execute global replacement
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "Old Company" \\
    --replace "New Company" \\
    --json
  
  # Targeted replacement (specific slide)
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --find "Draft" \\
    --replace "Final" \\
    --json
  
  # Surgical replacement (specific shape)
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --find "2024" \\
    --replace "2025" \\
    --json
  
  # Case-sensitive replacement
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "API" \\
    --replace "REST API" \\
    --match-case \\
    --json

Scope Options:
  Global (default):     All slides, all shapes
  Slide-specific:       --slide N (single slide, all shapes)
  Shape-specific:       --slide N --shape M (single shape)

Safety Recommendations:
  1. ALWAYS use --dry-run first to preview changes
  2. Clone the presentation before mass replacements:
     uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx
  3. Check dry-run output for unexpected matches
  4. Use --slide/--shape to limit scope when appropriate

Replacement Strategy:
  The tool uses a two-tier approach:
  1. Run-level replacement (preserves formatting)
  2. Shape-level fallback (for text split across runs)
  
  This ensures text is replaced even when PowerPoint splits it
  across multiple text runs, while preserving formatting when possible.

Finding Shape Indices:
  Use ppt_get_slide_info.py to identify shapes:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Output Format (dry-run):
  {
    "status": "success",
    "action": "dry_run",
    "find": "Old Company",
    "replace": "New Company",
    "scope": {"slide": "all", "shape": "all"},
    "total_matches": 15,
    "locations": [
      {"slide": 0, "shape": 1, "occurrences": 2, "preview": "Welcome to Old Company..."},
      {"slide": 3, "shape": 4, "occurrences": 1, "preview": "Old Company was founded..."}
    ],
    "tool_version": "3.1.0"
  }

Output Format (replace):
  {
    "status": "success",
    "action": "replace",
    "replacements_made": 15,
    "locations": [...],
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    parser.add_argument(
        '--find', 
        required=True, 
        help='Text to find'
    )
    parser.add_argument(
        '--replace', 
        required=True, 
        help='Replacement text'
    )
    parser.add_argument(
        '--slide', 
        type=int, 
        help='Target specific slide index (0-based)'
    )
    parser.add_argument(
        '--shape', 
        type=int, 
        help='Target specific shape index (requires --slide)'
    )
    parser.add_argument(
        '--match-case', 
        action='store_true', 
        help='Case-sensitive matching'
    )
    parser.add_argument(
        '--dry-run', 
        action='store_true', 
        help='Preview changes without modifying (RECOMMENDED first step)'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = replace_text(
            filepath=args.file,
            find=args.find,
            replace=args.replace,
            slide_index=args.slide,
            shape_index=args.shape,
            match_case=args.match_case,
            dry_run=args.dry_run
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check file format and parameter values"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_search_content.py
```py
#!/usr/bin/env python3
"""
PowerPoint Search Content Tool v3.1.1
Search for text content across all slides in a presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_search_content.py --file presentation.pptx --query "Revenue" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

This tool searches for text content across slides, including:
- Text in shapes and text boxes
- Slide titles and subtitles
- Speaker notes
- Table cell contents

Use this tool to locate content before using ppt_replace_text.py or to
navigate large presentations efficiently.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import re
from pathlib import Path
from typing import Dict, Any, List, Optional, Pattern

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# TYPE DEFINITIONS
# ============================================================================

Match = Dict[str, Any]


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def compile_pattern(
    query: str,
    is_regex: bool = False,
    case_sensitive: bool = False
) -> Pattern:
    """
    Compile search pattern from query string.
    
    Args:
        query: Search query (plain text or regex)
        is_regex: If True, treat query as regular expression
        case_sensitive: If True, perform case-sensitive search
        
    Returns:
        Compiled regex pattern
        
    Raises:
        ValueError: If regex is invalid
    """
    flags = 0 if case_sensitive else re.IGNORECASE
    
    if is_regex:
        try:
            return re.compile(query, flags)
        except re.error as e:
            raise ValueError(f"Invalid regex pattern: {e}")
    else:
        escaped = re.escape(query)
        return re.compile(escaped, flags)


def extract_context(text: str, match_start: int, match_end: int, context_chars: int = 50) -> str:
    """
    Extract text context around a match.
    
    Args:
        text: Full text
        match_start: Start position of match
        match_end: End position of match
        context_chars: Characters to include before/after
        
    Returns:
        Context string with match highlighted
    """
    start = max(0, match_start - context_chars)
    end = min(len(text), match_end + context_chars)
    
    prefix = "..." if start > 0 else ""
    suffix = "..." if end < len(text) else ""
    
    context = text[start:end]
    
    return f"{prefix}{context}{suffix}"


def search_text_frame(
    text_frame,
    pattern: Pattern,
    slide_index: int,
    shape_index: int,
    shape_name: str,
    shape_type: str,
    location: str = "text"
) -> List[Match]:
    """
    Search within a text frame.
    
    Args:
        text_frame: TextFrame object
        pattern: Compiled search pattern
        slide_index: Parent slide index
        shape_index: Parent shape index
        shape_name: Shape name
        shape_type: Shape type string
        location: Location identifier ("text" or "notes")
        
    Returns:
        List of match dictionaries
    """
    matches = []
    
    try:
        full_text = ""
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                full_text += run.text
            full_text += "\n"
        
        full_text = full_text.strip()
        
        if not full_text:
            return matches
        
        for match in pattern.finditer(full_text):
            matches.append({
                "slide_index": slide_index,
                "shape_index": shape_index,
                "shape_name": shape_name,
                "shape_type": shape_type,
                "location": location,
                "match_text": match.group(),
                "match_start": match.start(),
                "match_end": match.end(),
                "context": extract_context(full_text, match.start(), match.end())
            })
    except Exception:
        pass
    
    return matches


def search_table(
    table,
    pattern: Pattern,
    slide_index: int,
    shape_index: int,
    shape_name: str
) -> List[Match]:
    """
    Search within a table.
    
    Args:
        table: Table object
        pattern: Compiled search pattern
        slide_index: Parent slide index
        shape_index: Parent shape index
        shape_name: Shape name
        
    Returns:
        List of match dictionaries
    """
    matches = []
    
    try:
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text_frame.text if cell.text_frame else ""
                
                if not cell_text:
                    continue
                
                for match in pattern.finditer(cell_text):
                    matches.append({
                        "slide_index": slide_index,
                        "shape_index": shape_index,
                        "shape_name": shape_name,
                        "shape_type": "TABLE_CELL",
                        "location": "table",
                        "cell_row": row_idx,
                        "cell_col": col_idx,
                        "match_text": match.group(),
                        "match_start": match.start(),
                        "match_end": match.end(),
                        "context": extract_context(cell_text, match.start(), match.end())
                    })
    except Exception:
        pass
    
    return matches


# ============================================================================
# MAIN LOGIC
# ============================================================================

def search_content(
    filepath: Path,
    query: str,
    is_regex: bool = False,
    case_sensitive: bool = False,
    scope: str = "all",
    slide_index: Optional[int] = None
) -> Dict[str, Any]:
    """
    Search for content across a PowerPoint presentation.
    
    Args:
        filepath: Path to the PowerPoint file
        query: Search query (text or regex pattern)
        is_regex: If True, treat query as regular expression
        case_sensitive: If True, perform case-sensitive search
        scope: Search scope - "text", "notes", "tables", or "all"
        slide_index: Optional specific slide to search (None = all slides)
        
    Returns:
        Dict with search results
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If specified slide index is invalid
        ValueError: If regex pattern is invalid
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    pattern = compile_pattern(query, is_regex, case_sensitive)
    
    all_matches: List[Match] = []
    slides_searched: List[int] = []
    slides_with_matches: List[int] = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        presentation_version = agent.get_presentation_version()
        total_slides = agent.get_slide_count()
        
        if slide_index is not None:
            if not 0 <= slide_index < total_slides:
                raise SlideNotFoundError(
                    f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                    details={
                        "requested_index": slide_index,
                        "available_slides": total_slides
                    }
                )
            slides_to_search = [slide_index]
        else:
            slides_to_search = list(range(total_slides))
        
        for slide_idx in slides_to_search:
            slides_searched.append(slide_idx)
            slide = agent.prs.slides[slide_idx]
            slide_matches: List[Match] = []
            
            if scope in ["text", "all"]:
                for shape_idx, shape in enumerate(slide.shapes):
                    shape_name = getattr(shape, 'name', f'Shape_{shape_idx}')
                    shape_type = str(shape.shape_type).replace('MSO_SHAPE_TYPE.', '')
                    
                    if hasattr(shape, 'text_frame') and shape.has_text_frame:
                        matches = search_text_frame(
                            shape.text_frame,
                            pattern,
                            slide_idx,
                            shape_idx,
                            shape_name,
                            shape_type,
                            "text"
                        )
                        slide_matches.extend(matches)
                    
                    if scope in ["tables", "all"] and hasattr(shape, 'table') and shape.has_table:
                        matches = search_table(
                            shape.table,
                            pattern,
                            slide_idx,
                            shape_idx,
                            shape_name
                        )
                        slide_matches.extend(matches)
            
            if scope in ["notes", "all"]:
                try:
                    notes_slide = slide.notes_slide
                    if notes_slide and notes_slide.notes_text_frame:
                        notes_text = notes_slide.notes_text_frame.text
                        if notes_text:
                            for match in pattern.finditer(notes_text):
                                slide_matches.append({
                                    "slide_index": slide_idx,
                                    "shape_index": None,
                                    "shape_name": "Speaker Notes",
                                    "shape_type": "NOTES",
                                    "location": "notes",
                                    "match_text": match.group(),
                                    "match_start": match.start(),
                                    "match_end": match.end(),
                                    "context": extract_context(notes_text, match.start(), match.end())
                                })
                except Exception:
                    pass
            
            if slide_matches:
                slides_with_matches.append(slide_idx)
                all_matches.extend(slide_matches)
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "query": query,
        "options": {
            "regex": is_regex,
            "case_sensitive": case_sensitive,
            "scope": scope
        },
        "total_matches": len(all_matches),
        "slides_searched": len(slides_searched),
        "slides_with_matches": slides_with_matches,
        "matches": all_matches,
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Search for content across PowerPoint slides",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Simple text search
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "Revenue" --json

  # Case-sensitive search
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "Q4" --case-sensitive --json

  # Regex search for dates
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "\\d{4}-\\d{2}-\\d{2}" --regex --json

  # Search only in speaker notes
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "TODO" --scope notes --json

  # Search specific slide
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "Summary" --slide 5 --json

Scope Options:
  all    - Search everywhere (default)
  text   - Search only in text shapes
  notes  - Search only in speaker notes
  tables - Search only in table cells

Use Cases:
  1. Find slides before using ppt_replace_text.py
  2. Locate placeholder text to update
  3. Audit presentations for sensitive content
  4. Navigate large presentations efficiently
  5. Verify content updates were applied

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "query": "Revenue",
    "total_matches": 5,
    "slides_with_matches": [0, 2, 7],
    "matches": [
      {
        "slide_index": 0,
        "shape_index": 3,
        "shape_name": "Title 1",
        "shape_type": "PLACEHOLDER",
        "location": "text",
        "match_text": "Revenue",
        "context": "...Q4 Revenue Growth..."
      }
    ],
    "presentation_version": "a1b2c3...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to search'
    )
    
    parser.add_argument(
        '--query',
        required=True,
        type=str,
        help='Search query (text or regex pattern)'
    )
    
    parser.add_argument(
        '--regex',
        action='store_true',
        help='Treat query as regular expression'
    )
    
    parser.add_argument(
        '--case-sensitive',
        action='store_true',
        help='Perform case-sensitive search'
    )
    
    parser.add_argument(
        '--scope',
        choices=['all', 'text', 'notes', 'tables'],
        default='all',
        help='Search scope (default: all)'
    )
    
    parser.add_argument(
        '--slide',
        type=int,
        default=None,
        help='Limit search to specific slide index'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = search_content(
            filepath=args.file.resolve(),
            query=args.query,
            is_regex=args.regex,
            case_sensitive=args.case_sensitive,
            scope=args.scope,
            slide_index=args.slide
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check regex syntax if using --regex flag",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check file integrity and format",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_set_background.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Background Tool v3.1.0
Set slide background to a solid color or image.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --color "#FFFFFF" --json
    uv run tools/ppt_set_background.py --file deck.pptx --all-slides --color "#F5F5F5" --json
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --image background.jpg --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ColorHelper,
)

__version__ = "3.1.0"


def set_background(
    filepath: Path,
    color: Optional[str] = None,
    image: Optional[Path] = None,
    slide_index: Optional[int] = None,
    all_slides: bool = False
) -> Dict[str, Any]:
    """
    Set slide background to a solid color or image.
    
    Args:
        filepath: Path to PowerPoint file (.pptx only)
        color: Hex color code (e.g., "#FFFFFF")
        image: Path to background image file
        slide_index: Specific slide index (0-based), or None
        all_slides: If True, apply to all slides
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute path to file
            - slides_affected: Number of slides modified
            - slide_indices: List of modified slide indices
            - type: 'color' or 'image'
            - value: The color code or image path used
            - presentation_version_before: Version hash before changes
            - presentation_version_after: Version hash after changes
            - tool_version: Tool version string
            - deprecated_default_used: True if defaulted to all slides (backward compat)
            
    Raises:
        FileNotFoundError: If file or image doesn't exist
        ValueError: If parameters are invalid or mutually exclusive
        SlideNotFoundError: If slide index is out of range
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError(
            f"Invalid file format '{filepath.suffix}'. Only .pptx files are supported."
        )
    
    if not color and not image:
        raise ValueError("Must specify either --color or --image")
    
    if color and image:
        raise ValueError("Cannot specify both --color and --image; choose one")
    
    if slide_index is not None and all_slides:
        raise ValueError("Cannot specify both --slide and --all-slides; choose one")
    
    if color:
        color_clean = color.strip()
        if not color_clean.startswith('#'):
            color_clean = '#' + color_clean
        if len(color_clean) != 7:
            raise ValueError(
                f"Invalid color format '{color}'. Use hex format: #RRGGBB (e.g., #FFFFFF)"
            )
        try:
            int(color_clean[1:], 16)
        except ValueError:
            raise ValueError(
                f"Invalid color format '{color}'. Must contain valid hex characters."
            )
    
    if image and not image.exists():
        raise FileNotFoundError(f"Image file not found: {image}")
    
    deprecated_default_used = False
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        slide_count = agent.get_slide_count()
        
        if slide_count == 0:
            raise PowerPointAgentError("Presentation has no slides")
        
        if slide_index is not None:
            if not 0 <= slide_index < slide_count:
                raise SlideNotFoundError(
                    f"Slide index {slide_index} out of range (0-{slide_count - 1})",
                    details={"requested": slide_index, "available": slide_count}
                )
            target_indices = [slide_index]
        elif all_slides:
            target_indices = list(range(slide_count))
        else:
            target_indices = list(range(slide_count))
            deprecated_default_used = True
        
        for idx in target_indices:
            slide = agent.prs.slides[idx]
            bg = slide.background
            fill = bg.fill
            
            if color:
                fill.solid()
                fill.fore_color.rgb = ColorHelper.from_hex(color)
            elif image:
                fill.user_picture(str(image.resolve()))
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    result = {
        "status": "success",
        "file": str(filepath.resolve()),
        "slides_affected": len(target_indices),
        "slide_indices": target_indices,
        "type": "color" if color else "image",
        "value": color if color else str(image.resolve()),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if deprecated_default_used:
        result["deprecated_default_used"] = True
        result["deprecation_warning"] = (
            "Defaulting to all slides is deprecated. "
            "Future versions will require explicit --slide N or --all-slides flag."
        )
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Set PowerPoint slide background to a solid color or image",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Set single slide background to white
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --color "#FFFFFF" --json
    
    # Set all slides to light gray (explicit)
    uv run tools/ppt_set_background.py --file deck.pptx --all-slides --color "#F5F5F5" --json
    
    # Set background image on single slide
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --image bg.jpg --json
    
    # Set background image on all slides
    uv run tools/ppt_set_background.py --file deck.pptx --all-slides --image pattern.png --json

Color Format:
    Use hex color codes: #RRGGBB
    Examples: #FFFFFF (white), #000000 (black), #0070C0 (blue)

Supported Image Formats:
    PNG, JPEG, GIF, BMP, TIFF

Notes:
    - Use --slide N for a single slide (0-based index)
    - Use --all-slides to apply to entire presentation
    - If neither is specified, defaults to all slides (deprecated behavior)
    - Cannot use both --color and --image simultaneously
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='Path to PowerPoint file (.pptx)'
    )
    
    parser.add_argument(
        '--slide',
        type=int,
        dest='slide_index',
        help='Slide index to modify (0-based)'
    )
    
    parser.add_argument(
        '--all-slides',
        action='store_true',
        dest='all_slides',
        help='Apply background to all slides'
    )
    
    parser.add_argument(
        '--color',
        type=str,
        help='Hex color code (e.g., #FFFFFF)'
    )
    
    parser.add_argument(
        '--image',
        type=Path,
        help='Path to background image file'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output as JSON (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_background(
            filepath=args.file,
            color=args.color,
            image=args.image,
            slide_index=args.slide_index,
            all_slides=args.all_slides
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file and image paths exist and are accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide count."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check color format (#RRGGBB), ensure only one of --color or --image is used."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Verify the file is not corrupted and has at least one slide."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_set_footer.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Footer Tool v3.1.0
Configure slide footer with Dual Strategy (Placeholder + Text Box Fallback).

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_footer.py --file deck.pptx --text "Company © 2024" --json
    uv run tools/ppt_set_footer.py --file deck.pptx --text "Confidential" --show-number --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Set

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent

try:
    from pptx.enum.shapes import PP_PLACEHOLDER
except ImportError:
    class PP_PLACEHOLDER:
        FOOTER = 15
        SLIDE_NUMBER = 13

__version__ = "3.1.0"


def set_footer(
    filepath: Path,
    text: str = None,
    show_number: bool = False
) -> Dict[str, Any]:
    """
    Set footer on slides using Dual Strategy.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        text: Footer text
        show_number: Whether to show slide numbers
        
    Returns:
        Dict with results
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format invalid
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    slide_indices_updated: Set[int] = set()
    method_used = None
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        # Strategy 1: Try placeholders on slide masters
        try:
            for master in agent.prs.slide_masters:
                for layout in master.slide_layouts:
                    for shape in layout.placeholders:
                        try:
                            if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                                if text:
                                    shape.text = text
                        except Exception:
                            pass
        except Exception:
            pass
        
        # Try placeholders on slides
        for slide_idx, slide in enumerate(agent.prs.slides):
            try:
                for shape in slide.placeholders:
                    try:
                        if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                            if text:
                                shape.text = text
                            slide_indices_updated.add(slide_idx)
                    except Exception:
                        pass
            except Exception:
                pass
        
        # Strategy 2: Fallback to text boxes if placeholders didn't work
        if len(slide_indices_updated) == 0:
            method_used = "text_box"
            for slide_idx in range(len(agent.prs.slides)):
                try:
                    if text:
                        agent.add_text_box(
                            slide_index=slide_idx,
                            text=text,
                            position={"left": "5%", "top": "92%"},
                            size={"width": "60%", "height": "5%"},
                            font_size=10,
                            color="#595959"
                        )
                        slide_indices_updated.add(slide_idx)
                    if show_number:
                        agent.add_text_box(
                            slide_index=slide_idx,
                            text=str(slide_idx + 1),
                            position={"left": "92%", "top": "92%"},
                            size={"width": "5%", "height": "5%"},
                            font_size=10,
                            color="#595959"
                        )
                        slide_indices_updated.add(slide_idx)
                except Exception:
                    pass
        else:
            method_used = "placeholder"
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    return {
        "status": "success" if len(slide_indices_updated) > 0 else "warning",
        "file": str(filepath.resolve()),
        "method_used": method_used,
        "slides_updated": len(slide_indices_updated),
        "footer_text": text,
        "show_number": show_number,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Set slide footer with text and/or page numbers",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set footer text
  uv run tools/ppt_set_footer.py --file deck.pptx --text "Company © 2024" --json

  # Add page numbers
  uv run tools/ppt_set_footer.py --file deck.pptx --show-number --json

  # Both footer text and page numbers
  uv run tools/ppt_set_footer.py --file deck.pptx --text "Confidential" --show-number --json

Strategy:
  1. Tries to use slide placeholders first (preserves template formatting)
  2. Falls back to text boxes if placeholders not available
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path (.pptx)')
    parser.add_argument('--text', help='Footer text')
    parser.add_argument('--show-number', action='store_true', help='Show slide numbers')
    parser.add_argument('--show-date', action='store_true', help='Show date (placeholder only)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON (default: true)')
    
    args = parser.parse_args()
    
    try:
        result = set_footer(
            filepath=args.file,
            text=args.text,
            show_number=args.show_number
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Ensure file has .pptx extension."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_set_image_properties.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Image Properties Tool v3.1.0
Set alt text and opacity for image shapes (accessibility support)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_image_properties.py --file deck.pptx --slide 0 --shape 1 --alt-text "Company Logo" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Accessibility:
    Alt text is required for WCAG 2.1 compliance. All images should have
    descriptive alternative text that conveys the image's content and purpose.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional
import warnings

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Define fallback exception if not available in core
try:
    from core.powerpoint_agent_core import ShapeNotFoundError
except ImportError:
    class ShapeNotFoundError(PowerPointAgentError):
        """Exception raised when shape is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def set_image_properties(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    alt_text: Optional[str] = None,
    opacity: Optional[float] = None,
    transparency: Optional[float] = None  # Deprecated, for backward compat
) -> Dict[str, Any]:
    """
    Set properties on an image shape.
    
    Supports setting alternative text for accessibility and opacity
    for visual effects. At least one property must be specified.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the shape (0-based)
        shape_index: Index of the image shape (0-based)
        alt_text: Alternative text for accessibility (recommended for all images)
        opacity: Image opacity from 0.0 (invisible) to 1.0 (opaque)
        transparency: DEPRECATED - use opacity instead. If provided, converted to opacity.
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the shape
            - properties_set: Dict of properties that were set
            - presentation_version_before: State hash before modification
            - presentation_version_after: State hash after modification
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        ShapeNotFoundError: If the shape index is out of range
        ValueError: If no properties specified or invalid values
        
    Example:
        >>> result = set_image_properties(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     shape_index=1,
        ...     alt_text="Company Logo - Blue and white design"
        ... )
        >>> print(result["properties_set"]["alt_text"])
        'Company Logo - Blue and white design'
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Handle deprecated transparency parameter
    effective_opacity = opacity
    transparency_converted = False
    
    if transparency is not None:
        if opacity is not None:
            raise ValueError(
                "Cannot specify both 'opacity' and 'transparency'. "
                "Use 'opacity' (transparency is deprecated)."
            )
        # Convert transparency to opacity (inverse relationship)
        effective_opacity = 1.0 - transparency
        transparency_converted = True
    
    # Validate at least one property is being set
    if alt_text is None and effective_opacity is None:
        raise ValueError(
            "At least one property must be set (--alt-text or --opacity)"
        )
    
    # Validate opacity range
    if effective_opacity is not None:
        if not (0.0 <= effective_opacity <= 1.0):
            raise ValueError(
                f"Opacity must be between 0.0 and 1.0, got: {effective_opacity}"
            )

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE modification
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Get slide info to validate shape index
        slide_info = agent.get_slide_info(slide_index)
        shape_count = slide_info.get("shape_count", 0)
        
        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{shape_count - 1})",
                details={
                    "requested_index": shape_index,
                    "available_shapes": shape_count
                }
            )
        
        # Set image properties
        # Note: Core method may use different parameter names
        try:
            agent.set_image_properties(
                slide_index=slide_index,
                shape_index=shape_index,
                alt_text=alt_text,
                # Pass opacity as fill_opacity if core supports it
                fill_opacity=effective_opacity
            )
        except TypeError:
            # Fallback if core uses different signature
            agent.set_image_properties(
                slide_index=slide_index,
                shape_index=shape_index,
                alt_text=alt_text,
                transparency=1.0 - effective_opacity if effective_opacity is not None else None
            )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER modification
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Build properties dict
    properties_set = {}
    if alt_text is not None:
        properties_set["alt_text"] = alt_text
    if effective_opacity is not None:
        properties_set["opacity"] = effective_opacity
        if transparency_converted:
            properties_set["transparency_converted"] = True
            properties_set["original_transparency"] = transparency
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "properties_set": properties_set,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Set image properties (alt text, opacity) in PowerPoint",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set alt text for accessibility
  uv run tools/ppt_set_image_properties.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --alt-text "Company Logo - Blue and white circular design" \\
    --json
  
  # Set opacity for watermark effect
  uv run tools/ppt_set_image_properties.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --shape 3 \\
    --opacity 0.3 \\
    --json
  
  # Set both properties
  uv run tools/ppt_set_image_properties.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 0 \\
    --alt-text "Background watermark" \\
    --opacity 0.15 \\
    --json

Finding Shape Indices:
  Use ppt_get_slide_info.py to identify shape indices:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Alt Text Guidelines (WCAG 2.1):
  - Describe image content and purpose
  - For logos: "Company Name Logo"
  - For charts: Include key data points
  - For photos: Describe what's shown
  - For decorative images: Use empty string ""
  - Keep under 125 characters when possible

Opacity Values:
  - 0.0 = Fully transparent (invisible)
  - 0.5 = 50% visible
  - 1.0 = Fully opaque (default)
  
  Use Cases:
  - Watermarks: 0.1-0.2
  - Background images: 0.3-0.5
  - Subtle overlays: 0.15-0.25

Deprecation Notice:
  --transparency is deprecated. Use --opacity instead.
  transparency = 1.0 - opacity

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "shape_index": 1,
    "properties_set": {
      "alt_text": "Company Logo",
      "opacity": 1.0
    },
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    parser.add_argument(
        '--slide', 
        required=True, 
        type=int, 
        help='Slide index (0-based)'
    )
    parser.add_argument(
        '--shape', 
        required=True, 
        type=int, 
        help='Shape index (0-based)'
    )
    parser.add_argument(
        '--alt-text', 
        help='Alternative text for accessibility'
    )
    parser.add_argument(
        '--opacity', 
        type=float, 
        help='Opacity from 0.0 (invisible) to 1.0 (opaque)'
    )
    parser.add_argument(
        '--transparency', 
        type=float, 
        help='DEPRECATED: Use --opacity instead. Transparency from 0.0 to 1.0'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_image_properties(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            alt_text=args.alt_text,
            opacity=args.opacity,
            transparency=args.transparency
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Specify at least --alt-text or --opacity (0.0-1.0)"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_set_slide_layout.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Slide Layout Tool v3.1.0
Change the layout of an existing slide with safety warnings

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

⚠️ IMPORTANT WARNING:
    Changing slide layouts can cause CONTENT LOSS!
    - Text in removed placeholders may disappear
    - Shapes may be repositioned
    - This is a python-pptx limitation
    
    ALWAYS backup your presentation before changing layouts!

Usage:
    uv run tools/ppt_set_slide_layout.py --file presentation.pptx --slide 2 --layout "Title Only" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Safety:
    The --force flag is required for layouts that may cause content loss
    (e.g., "Blank", "Title Only")
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional
from difflib import get_close_matches

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Define fallback exception
try:
    from core.powerpoint_agent_core import LayoutNotFoundError
except ImportError:
    class LayoutNotFoundError(PowerPointAgentError):
        """Exception raised when layout is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)

# Layouts known to potentially cause content loss
DESTRUCTIVE_LAYOUTS = ["Blank", "Title Only"]


def set_slide_layout(
    filepath: Path,
    slide_index: int,
    layout_name: str,
    force: bool = False
) -> Dict[str, Any]:
    """
    Change slide layout with safety warnings.
    
    ⚠️ WARNING: Changing layouts can cause content loss due to python-pptx
    limitations. Always backup presentations before layout changes.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Slide index (0-based)
        layout_name: Target layout name (fuzzy matching supported)
        force: Acknowledge content loss risk (required for destructive layouts)
        
    Returns:
        Dict containing:
            - status: "success" or "warning"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - old_layout: Previous layout name
            - new_layout: New layout name
            - layout_changed: Whether layout actually changed
            - placeholders: Before/after/change counts
            - available_layouts: All available layouts
            - warnings: Content loss warnings (if any)
            - recommendations: Suggested actions (if any)
            - presentation_version_before: State hash before change
            - presentation_version_after: State hash after change
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        LayoutNotFoundError: If layout is not found
        PowerPointAgentError: If force required but not provided
        
    Example:
        >>> result = set_slide_layout(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=2,
        ...     layout_name="Section Header"
        ... )
        >>> print(result["new_layout"])
        'Section Header'
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    warnings: List[str] = []
    recommendations: List[str] = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE change
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Get available layouts
        available_layouts = agent.get_available_layouts()
        
        # Get current slide info
        slide_info_before = agent.get_slide_info(slide_index)
        old_layout = slide_info_before.get("layout", "Unknown")
        placeholders_before = sum(
            1 for shape in slide_info_before.get("shapes", [])
            if "PLACEHOLDER" in shape.get("type", "")
        )
        
        # Layout name matching with fuzzy search
        matched_layout: Optional[str] = None
        
        # Exact match (case-insensitive)
        for layout in available_layouts:
            if layout.lower() == layout_name.lower():
                matched_layout = layout
                break
        
        # Substring match if no exact match
        if not matched_layout:
            for layout in available_layouts:
                if layout_name.lower() in layout.lower():
                    matched_layout = layout
                    warnings.append(
                        f"Matched '{layout_name}' to layout '{layout}' (substring match)"
                    )
                    break
        
        # Fuzzy match using difflib
        if not matched_layout:
            close_matches = get_close_matches(
                layout_name, available_layouts, n=3, cutoff=0.6
            )
            if close_matches:
                raise LayoutNotFoundError(
                    f"Layout '{layout_name}' not found. Did you mean one of these?\n" +
                    "\n".join(f"  - {match}" for match in close_matches) +
                    f"\n\nAll available layouts:\n" +
                    "\n".join(f"  - {layout}" for layout in available_layouts),
                    details={
                        "requested_layout": layout_name,
                        "suggestions": close_matches,
                        "available_layouts": available_layouts
                    }
                )
            else:
                raise LayoutNotFoundError(
                    f"Layout '{layout_name}' not found.\n\n" +
                    f"Available layouts:\n" +
                    "\n".join(f"  - {layout}" for layout in available_layouts),
                    details={
                        "requested_layout": layout_name,
                        "available_layouts": available_layouts
                    }
                )
        
        # Safety warnings for destructive layouts
        if matched_layout in DESTRUCTIVE_LAYOUTS and placeholders_before > 0:
            warnings.append(
                f"⚠️ CONTENT LOSS RISK: Changing from '{old_layout}' to '{matched_layout}' "
                f"may remove {placeholders_before} placeholder(s) and their content!"
            )
            
            if not force:
                raise PowerPointAgentError(
                    f"Layout change from '{old_layout}' to '{matched_layout}' requires --force flag.\n"
                    f"This change may cause content loss ({placeholders_before} placeholders affected).\n\n"
                    "To proceed, add --force flag:\n"
                    f"  --layout \"{matched_layout}\" --force\n\n"
                    "RECOMMENDATION: Backup your presentation first!"
                )
        
        # Warn about same layout
        if matched_layout == old_layout:
            recommendations.append(
                f"Slide already uses '{old_layout}' layout. No change needed."
            )
        
        # Apply layout change
        agent.set_slide_layout(slide_index, matched_layout)
        
        # Get slide info after change
        slide_info_after = agent.get_slide_info(slide_index)
        placeholders_after = sum(
            1 for shape in slide_info_after.get("shapes", [])
            if "PLACEHOLDER" in shape.get("type", "")
        )
        
        # Detect content loss
        if placeholders_after < placeholders_before:
            lost_count = placeholders_before - placeholders_after
            warnings.append(
                f"Content loss detected: {lost_count} placeholder(s) removed during layout change."
            )
            recommendations.append(
                "Review slide content and restore any lost text using ppt_add_text_box.py"
            )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER change
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Build response
    status = "success" if len(warnings) == 0 else "warning"
    
    result: Dict[str, Any] = {
        "status": status,
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "old_layout": old_layout,
        "new_layout": matched_layout,
        "layout_changed": (old_layout != matched_layout),
        "placeholders": {
            "before": placeholders_before,
            "after": placeholders_after,
            "change": placeholders_after - placeholders_before
        },
        "available_layouts": available_layouts,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if warnings:
        result["warnings"] = warnings
    
    if recommendations:
        result["recommendations"] = recommendations
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Change PowerPoint slide layout with safety warnings",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️ IMPORTANT WARNING ⚠️
    Changing slide layouts can cause CONTENT LOSS!
    - Text in removed placeholders may disappear
    - Shapes may be repositioned
    
    ALWAYS backup your presentation before changing layouts!

Examples:
  # List available layouts first
  uv run tools/ppt_get_info.py --file presentation.pptx --json | jq '.layouts'
  
  # Change to Title Only layout (low risk)
  uv run tools/ppt_set_slide_layout.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --layout "Title Only" \\
    --json
  
  # Change to Blank layout (HIGH RISK - requires --force)
  uv run tools/ppt_set_slide_layout.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --layout "Blank" \\
    --force \\
    --json
  
  # Fuzzy matching (will match "Title and Content")
  uv run tools/ppt_set_slide_layout.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --layout "title content" \\
    --json

Common Layouts:
  Low Risk (preserve most content):
  - "Title and Content" - Most versatile
  - "Two Content" - Side-by-side content
  - "Section Header" - Section dividers
  
  Medium Risk:
  - "Title Only" - Removes content placeholders
  - "Content with Caption" - Repositions content
  
  High Risk (requires --force):
  - "Blank" - Removes ALL placeholders!

Layout Matching:
  This tool supports flexible matching:
  - Exact: "Title and Content" matches "Title and Content"
  - Case-insensitive: "title slide" matches "Title Slide"
  - Substring: "content" matches "Title and Content"
  - Fuzzy: "tile slide" suggests "Title Slide"

Safety Features:
  - Warns about content loss risk
  - Requires --force for destructive layouts
  - Reports placeholder count changes
  - Suggests recovery actions

Output Format:
  {
    "status": "warning",
    "slide_index": 2,
    "old_layout": "Title and Content",
    "new_layout": "Title Only",
    "layout_changed": true,
    "placeholders": {
      "before": 2,
      "after": 1,
      "change": -1
    },
    "warnings": ["Content loss detected..."],
    "recommendations": ["Review slide content..."],
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }

Recovery from Content Loss:
  If content was lost during layout change:
  1. Restore from backup (you did backup, right?)
  2. Use ppt_get_slide_info.py to inspect current state
  3. Restore text with ppt_add_text_box.py

Related Tools:
  - ppt_get_info.py: List all available layouts
  - ppt_get_slide_info.py: Inspect current slide layout
  - ppt_add_text_box.py: Restore lost content
  - ppt_clone_presentation.py: Create backup before changes
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--layout',
        required=True,
        help='New layout name (fuzzy matching supported)'
    )
    
    parser.add_argument(
        '--force',
        action='store_true',
        help='Force destructive layout change (acknowledges content loss risk)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_slide_layout(
            filepath=args.file,
            slide_index=args.slide,
            layout_name=args.layout,
            force=args.force
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except LayoutNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "LayoutNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to list available layouts"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Add --force flag if you accept the content loss risk"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None,
            "slide_index": args.slide if hasattr(args, 'slide') else None,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_set_title.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Title Tool v3.1.0
Set slide title and optional subtitle with comprehensive validation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_title.py --file presentation.pptx --slide 0 --title "Q4 Results" --json
    uv run tools/ppt_set_title.py --file deck.pptx --slide 0 --title "2024 Strategy" \\
        --subtitle "Growth & Innovation" --json

Exit Codes:
    0: Success
    1: Error occurred

Best Practices:
- Keep titles under 60 characters for readability
- Keep subtitles under 100 characters
- Use "Title Slide" layout for first slide (index 0)
- Use title case: "This Is Title Case"
- Subtitles provide context, not repetition
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
)

__version__ = "3.1.0"


def set_title(
    filepath: Path,
    slide_index: int,
    title: str,
    subtitle: Optional[str] = None
) -> Dict[str, Any]:
    """
    Set slide title and subtitle with validation.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Slide index (0-based)
        title: Title text
        subtitle: Optional subtitle text
        
    Returns:
        Dict containing:
        - status: "success" or "warning"
        - file: Absolute file path
        - slide_index: Modified slide
        - title: Title set
        - subtitle: Subtitle set (if any)
        - layout: Current layout name
        - warnings: List of validation warnings
        - recommendations: Suggested improvements
        - presentation_version_before: Version hash before
        - presentation_version_after: Version hash after
        - tool_version: Tool version
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format invalid
        SlideNotFoundError: If slide index out of range
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    warnings: List[str] = []
    recommendations: List[str] = []
    
    if len(title) > 60:
        warnings.append(
            f"Title is {len(title)} characters (recommended: ≤60 for readability). "
            "Consider shortening for better visual impact."
        )
    
    if len(title) > 100:
        warnings.append(
            "Title exceeds 100 characters and may not fit on slide. "
            "Strong recommendation to shorten."
        )
    
    if subtitle and len(subtitle) > 100:
        warnings.append(
            f"Subtitle is {len(subtitle)} characters (recommended: ≤100). "
            "Long subtitles reduce readability."
        )
    
    if title == title.upper() and len(title) > 10:
        recommendations.append(
            "Title is all uppercase. Consider using title case for better readability: "
            "'This Is Title Case' instead of 'THIS IS TITLE CASE'"
        )
    
    if title == title.lower() and len(title) > 10:
        recommendations.append(
            "Title is all lowercase. Consider using title case for professionalism."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        slide_info_before = agent.get_slide_info(slide_index)
        layout_name = slide_info_before.get("layout", "Unknown")
        
        if slide_index == 0 and "Title Slide" not in layout_name:
            recommendations.append(
                f"First slide has layout '{layout_name}'. "
                "Consider using 'Title Slide' layout for cover slides."
            )
        
        has_title_placeholder = False
        has_subtitle_placeholder = False
        
        for shape in slide_info_before.get("shapes", []):
            shape_type = shape.get("type", "")
            if "TITLE" in shape_type or "CENTER_TITLE" in shape_type:
                has_title_placeholder = True
            if "SUBTITLE" in shape_type:
                has_subtitle_placeholder = True
        
        if not has_title_placeholder:
            warnings.append(
                f"Layout '{layout_name}' may not have a title placeholder. "
                "Title may not display as expected. Consider changing layout first."
            )
        
        if subtitle and not has_subtitle_placeholder:
            warnings.append(
                f"Layout '{layout_name}' does not have a subtitle placeholder. "
                "Subtitle will not be displayed. Consider using 'Title Slide' layout."
            )
        
        agent.set_title(slide_index, title, subtitle)
        
        slide_info_after = agent.get_slide_info(slide_index)
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    status = "success" if len(warnings) == 0 else "warning"
    
    result: Dict[str, Any] = {
        "status": status,
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "title": title,
        "subtitle": subtitle,
        "layout": layout_name,
        "shape_count": slide_info_after.get("shape_count", 0),
        "placeholders_found": {
            "title": has_title_placeholder,
            "subtitle": has_subtitle_placeholder
        },
        "validation": {
            "title_length": len(title),
            "title_length_ok": len(title) <= 60,
            "subtitle_length": len(subtitle) if subtitle else 0,
            "subtitle_length_ok": len(subtitle) <= 100 if subtitle else True
        },
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if warnings:
        result["warnings"] = warnings
    
    if recommendations:
        result["recommendations"] = recommendations
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Set PowerPoint slide title and subtitle with validation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set title only
  uv run tools/ppt_set_title.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --title "Q4 Financial Results" \\
    --json
  
  # Set title and subtitle (first slide)
  uv run tools/ppt_set_title.py \\
    --file deck.pptx \\
    --slide 0 \\
    --title "2024 Strategic Plan" \\
    --subtitle "Driving Growth and Innovation" \\
    --json
  
  # Update section title (middle slide)
  uv run tools/ppt_set_title.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --title "Market Analysis" \\
    --json

Best Practices:
  Title Guidelines:
  - Keep under 60 characters (optimal readability)
  - Use title case: "This Is Title Case"
  - Be specific and descriptive
  - Avoid jargon and abbreviations
  - One clear message per title
  
  Subtitle Guidelines:
  - Keep under 100 characters
  - Provide context, not repetition
  - Use for date, location, or clarification
  - Optional on content slides
  
  Layout Recommendations:
  - Slide 0 (first): Use "Title Slide" layout
  - Section headers: Use "Section Header" layout
  - Content slides: Use "Title and Content" layout
  - Blank slides: Use "Title Only" layout

Validation:
  This tool performs automatic validation:
  - Title length (warns if >60 chars, strong warning if >100)
  - Subtitle length (warns if >100 chars)
  - Title case recommendations
  - Placeholder availability checks
  - Layout compatibility warnings

Related Tools:
  - ppt_get_slide_info.py: Inspect slide layout and placeholders
  - ppt_set_slide_layout.py: Change slide layout
  - ppt_get_info.py: Get presentation info (total slides, layouts)
  - ppt_add_text_box.py: Add custom text if placeholders unavailable
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path (.pptx)'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based, e.g., 0 for first slide)'
    )
    
    parser.add_argument(
        '--title',
        required=True,
        help='Title text (recommended: ≤60 characters)'
    )
    
    parser.add_argument(
        '--subtitle',
        help='Optional subtitle text (recommended: ≤100 characters)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_title(
            filepath=args.file,
            slide_index=args.slide,
            title=args.title,
            subtitle=args.subtitle
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Ensure file is .pptx format."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check the presentation file is valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_set_z_order.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Z-Order Tool v3.1.0
Manage shape layering (Bring to Front, Send to Back, etc.).

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_z_order.py --file deck.pptx --slide 0 --shape 1 --action bring_to_front --json

Exit Codes:
    0: Success
    1: Error occurred

⚠️  IMPORTANT: Shape indices change after z-order operations!
    Always refresh indices with ppt_get_slide_info.py before targeting shapes.
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
)

__version__ = "3.1.0"


def _validate_xml_structure(sp_tree) -> bool:
    """Validate XML tree integrity after manipulation."""
    return all(child is not None for child in sp_tree)


def set_z_order(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    action: str
) -> Dict[str, Any]:
    """
    Change the Z-order (stacking order) of a shape.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        shape_index: Target shape index (0-based)
        action: One of 'bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'
        
    Returns:
        Result dict with z-order change details
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format invalid or invalid action
        SlideNotFoundError: If slide index invalid
        ShapeNotFoundError: If shape index invalid
        PowerPointAgentError: If XML manipulation fails
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    valid_actions = ['bring_to_front', 'send_to_back', 'bring_forward', 'send_backward']
    if action not in valid_actions:
        raise ValueError(f"Invalid action '{action}'. Must be one of: {valid_actions}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        slide_count = agent.get_slide_count()
        if not 0 <= slide_index < slide_count:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{slide_count - 1})",
                details={"requested": slide_index, "available": slide_count}
            )
        
        slide = agent.prs.slides[slide_index]
        shape_count = len(slide.shapes)
        
        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{shape_count - 1})",
                details={"requested": shape_index, "available": shape_count}
            )
        
        shape = slide.shapes[shape_index]
        
        # XML Manipulation for Z-Order
        sp_tree = slide.shapes._spTree
        element = shape.element
        
        # Find current position in XML tree
        current_index = -1
        for i, child in enumerate(sp_tree):
            if child == element:
                current_index = i
                break
        
        if current_index == -1:
            raise PowerPointAgentError("Could not locate shape in XML tree")
        
        new_index = current_index
        max_index = len(sp_tree) - 1
        
        # Execute Z-Order Action
        if action == 'bring_to_front':
            sp_tree.remove(element)
            sp_tree.append(element)
            new_index = max_index
            
        elif action == 'send_to_back':
            sp_tree.remove(element)
            sp_tree.insert(0, element)
            new_index = 0
            
        elif action == 'bring_forward':
            if current_index < max_index:
                sp_tree.remove(element)
                sp_tree.insert(current_index + 1, element)
                new_index = current_index + 1
                
        elif action == 'send_backward':
            if current_index > 0:
                sp_tree.remove(element)
                sp_tree.insert(current_index - 1, element)
                new_index = current_index - 1
        
        # Validate XML structure after manipulation
        if not _validate_xml_structure(sp_tree):
            raise PowerPointAgentError("XML structure corrupted during Z-order operation")
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index_target": shape_index,
        "action": action,
        "z_order_change": {
            "from": current_index,
            "to": new_index
        },
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__,
        "warning": "⚠️ Shape indices may have changed. Use ppt_get_slide_info.py to refresh before further operations.",
        "refresh_command": f"uv run tools/ppt_get_slide_info.py --file {filepath} --slide {slide_index} --json"
    }


def main():
    parser = argparse.ArgumentParser(
        description="Set shape Z-Order (layering)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Actions:
  bring_to_front  - Move shape to top layer (in front of all)
  send_to_back    - Move shape to bottom layer (behind all)
  bring_forward   - Move shape up one layer
  send_backward   - Move shape down one layer

Examples:
  # Send overlay to back (for readability overlays)
  uv run tools/ppt_set_z_order.py --file deck.pptx --slide 0 --shape 5 \\
    --action send_to_back --json

  # Bring logo to front
  uv run tools/ppt_set_z_order.py --file deck.pptx --slide 0 --shape 2 \\
    --action bring_to_front --json

⚠️  IMPORTANT: Shape indices change after z-order operations!
    Always run ppt_get_slide_info.py to refresh indices before targeting shapes.
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path (.pptx)')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--action', required=True,
                        choices=['bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'],
                        help='Layering action')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON (default: true)')
    
    args = parser.parse_args()
    
    try:
        result = set_z_order(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            action=args.action
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shapes."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check file format (.pptx) and action is valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "XML manipulation failed. File may be corrupted."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_update_chart_data.py
```py
#!/usr/bin/env python3
"""
PowerPoint Update Chart Data Tool v3.1.0
Update the data of an existing chart

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_update_chart_data.py --file deck.pptx --slide 0 --chart 0 --data new_data.json --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

⚠️ LIMITATION WARNING:
    python-pptx has LIMITED chart update support. The replace_data() method
    may fail if the new data schema doesn't match the original chart exactly.
    
    If update fails, consider the alternative approach:
    1. Delete the existing chart: ppt_remove_shape.py
    2. Add a new chart with new data: ppt_add_chart.py
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Import CategoryChartData safely
try:
    from pptx.chart.data import CategoryChartData
    CHART_DATA_AVAILABLE = True
except ImportError:
    CHART_DATA_AVAILABLE = False
    CategoryChartData = None


def update_chart_data(
    filepath: Path,
    slide_index: int,
    chart_index: int,
    data: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Update the data of an existing chart.
    
    Replaces the chart's data with new categories and series values.
    The new data must be compatible with the existing chart type.
    
    ⚠️ LIMITATION: python-pptx's replace_data() may fail if the new
    data structure doesn't match the original. If this fails, consider
    deleting the chart and creating a new one.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the chart (0-based)
        chart_index: Index of the chart on the slide (0-based)
        data: New chart data dict with 'categories' and 'series' keys
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - chart_index: Index of the chart
            - categories: Number of categories
            - series: Number of data series
            - data_points: Total data points updated
            - presentation_version_before: State hash before update
            - presentation_version_after: State hash after update
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ValueError: If data format is invalid or chart not found
        RuntimeError: If chart data update fails (python-pptx limitation)
        
    Example:
        >>> data = {
        ...     "categories": ["Q1", "Q2", "Q3"],
        ...     "series": [{"name": "Sales", "values": [100, 150, 200]}]
        ... }
        >>> result = update_chart_data(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=1,
        ...     chart_index=0,
        ...     data=data
        ... )
        >>> print(result["data_points"])
        3
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate data structure
    if "categories" not in data:
        raise ValueError(
            "Data must contain 'categories' key. "
            "Example: {\"categories\": [\"A\", \"B\"], \"series\": [...]}"
        )
    
    if "series" not in data or not data["series"]:
        raise ValueError(
            "Data must contain at least one series. "
            "Example: {\"series\": [{\"name\": \"Sales\", \"values\": [10, 20]}]}"
        )
    
    # Validate series data
    cat_len = len(data["categories"])
    for i, series in enumerate(data["series"]):
        if "name" not in series:
            raise ValueError(f"Series {i} missing 'name' key")
        if "values" not in series:
            raise ValueError(f"Series {i} missing 'values' key")
        if len(series["values"]) != cat_len:
            raise ValueError(
                f"Series '{series['name']}' has {len(series['values'])} values, "
                f"but there are {cat_len} categories. Counts must match."
            )
    
    # Check if CategoryChartData is available
    if not CHART_DATA_AVAILABLE:
        raise RuntimeError(
            "pptx.chart.data.CategoryChartData not available. "
            "Ensure python-pptx is properly installed."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE update
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # NOTE: Direct prs access required for chart data manipulation
        # python-pptx requires direct access to chart objects for replace_data()
        slide = agent.prs.slides[slide_index]
        
        # Find charts on slide
        charts = [shape for shape in slide.shapes if shape.has_chart]
        
        if not charts:
            raise ValueError(
                f"No charts found on slide {slide_index}. "
                "Use ppt_add_chart.py to create a chart first."
            )
        
        if not 0 <= chart_index < len(charts):
            raise ValueError(
                f"Chart index {chart_index} out of range. "
                f"Slide has {len(charts)} chart(s) (indices 0-{len(charts) - 1})."
            )
        
        chart_shape = charts[chart_index]
        chart = chart_shape.chart
        
        # Create new chart data
        chart_data = CategoryChartData()
        chart_data.categories = data["categories"]
        
        for series in data["series"]:
            chart_data.add_series(series["name"], series["values"])
        
        # Attempt to replace data
        try:
            chart.replace_data(chart_data)
        except Exception as e:
            raise RuntimeError(
                f"Failed to update chart data: {e}. "
                "This may be due to python-pptx limitations with complex charts. "
                "Consider deleting the chart (ppt_remove_shape.py) and "
                "creating a new one (ppt_add_chart.py) instead."
            )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER update
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "chart_index": chart_index,
        "categories": len(data["categories"]),
        "series": len(data["series"]),
        "data_points": sum(len(s["values"]) for s in data["series"]),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Update PowerPoint chart data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️ LIMITATION WARNING:
  python-pptx has LIMITED chart update support. The replace_data()
  method may fail if the new data doesn't match the original chart.
  
  If this tool fails, use the alternative approach:
  1. Get chart position: ppt_get_slide_info.py
  2. Delete chart: ppt_remove_shape.py (with approval token)
  3. Create new chart: ppt_add_chart.py

Examples:
  # Update from JSON file
  uv run tools/ppt_update_chart_data.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart 0 \\
    --data updated_data.json \\
    --json
  
  # Update with inline data
  uv run tools/ppt_update_chart_data.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --chart 0 \\
    --data-string '{"categories":["Q1","Q2","Q3"],"series":[{"name":"Sales","values":[100,150,200]}]}' \\
    --json

Data Format (JSON):
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "Revenue", "values": [100, 120, 140, 160]},
    {"name": "Costs", "values": [80, 90, 100, 110]}
  ]
}

Requirements:
  - Number of values in each series must match number of categories
  - Each series must have 'name' and 'values' keys
  - Data structure should match original chart type

Finding Charts:
  Use ppt_get_slide_info.py to identify charts:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Common Issues:
  - "Failed to update chart data": Schema mismatch
    Solution: Delete and recreate the chart
  
  - "No charts found": Slide has no charts
    Solution: Use ppt_add_chart.py to create one
  
  - Series count mismatch may cause issues
    Solution: Match the original number of series

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 1,
    "chart_index": 0,
    "categories": 4,
    "series": 2,
    "data_points": 8,
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    parser.add_argument(
        '--slide', 
        required=True, 
        type=int, 
        help='Slide index (0-based)'
    )
    parser.add_argument(
        '--chart', 
        required=True, 
        type=int, 
        help='Chart index on slide (0-based)'
    )
    parser.add_argument(
        '--data', 
        type=Path, 
        help='JSON file with chart data'
    )
    parser.add_argument(
        '--data-string',
        help='Inline JSON data string'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        # Load chart data
        if args.data:
            if not args.data.exists():
                raise FileNotFoundError(f"Data file not found: {args.data}")
            with open(args.data, 'r') as f:
                data_content = json.load(f)
        elif args.data_string:
            try:
                data_content = json.loads(args.data_string)
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSON in --data-string: {e}")
        else:
            raise ValueError("Either --data or --data-string is required")
        
        result = update_chart_data(
            filepath=args.file, 
            slide_index=args.slide, 
            chart_index=args.chart,
            data=data_content
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file paths exist and are accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check data format and chart index"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except RuntimeError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "RuntimeError",
            "suggestion": "Consider deleting chart and creating new one with ppt_add_chart.py"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_validate_presentation.py
```py
#!/usr/bin/env python3
"""
PowerPoint Validate Presentation Tool v3.1.1
Comprehensive validation for structure, accessibility, assets, and design quality.

Fully aligned with PowerPoint Agent Core v3.1.0+ and System Prompt v3.0 validation gates.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --json
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --policy strict --json

Exit Codes:
    0: Success (valid or only warnings within policy thresholds)
    1: Error occurred or critical issues exceed policy thresholds

Changelog v3.1.1:
    - Added presentation_version to output for audit trail
    - Populated fix_command for actionable remediation
    - Expanded _validate_design_rules with color and 6x6 rule checking
    - Added tool_version to output
    - Added acquire_lock documentation comments
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that JSON parsers only see valid JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional, Set
from dataclasses import dataclass, field, asdict
from datetime import datetime

# Configure logging to null handler to prevent any accidental output
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from core.powerpoint_agent_core import (
        PowerPointAgent,
        PowerPointAgentError,
        __version__ as CORE_VERSION
    )
except ImportError:
    CORE_VERSION = "0.0.0"
    PowerPointAgent = None
    PowerPointAgentError = Exception

# ============================================================================
# CONSTANTS & POLICIES
# ============================================================================

__version__ = "3.1.1"

VALIDATION_POLICIES = {
    "lenient": {
        "name": "Lenient",
        "description": "Minimal validation - suitable for drafts and work-in-progress",
        "thresholds": {
            "max_critical_issues": 10,
            "max_accessibility_issues": 20,
            "max_design_warnings": 50,
            "max_empty_slides": 5,
            "max_slides_without_titles": 10,
            "max_missing_alt_text": 20,
            "max_low_contrast": 10,
            "max_large_images": 10,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 10,
            "max_colors": 20,
            "min_font_size_pt": 8,
        }
    },
    "standard": {
        "name": "Standard",
        "description": "Balanced validation - suitable for internal presentations",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 5,
            "max_design_warnings": 10,
            "max_empty_slides": 0,
            "max_slides_without_titles": 3,
            "max_missing_alt_text": 5,
            "max_low_contrast": 3,
            "max_large_images": 5,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 5,
            "max_colors": 10,
            "min_font_size_pt": 10,
        }
    },
    "strict": {
        "name": "Strict",
        "description": "Maximum validation - suitable for external/production presentations",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 0,
            "max_design_warnings": 3,
            "max_empty_slides": 0,
            "max_slides_without_titles": 0,
            "max_missing_alt_text": 0,
            "max_low_contrast": 0,
            "max_large_images": 3,
            "require_all_alt_text": True,
            "enforce_6x6_rule": True,
            "max_fonts": 3,
            "max_colors": 5,
            "min_font_size_pt": 12,
        }
    }
}

# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class ValidationIssue:
    """Represents a single validation issue found in the presentation."""
    category: str
    severity: str
    message: str
    slide_index: Optional[int] = None
    shape_index: Optional[int] = None
    fix_command: Optional[str] = None
    details: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary, excluding None values."""
        result = {}
        for key, value in asdict(self).items():
            if value is not None and value != {}:
                result[key] = value
        return result


@dataclass
class ValidationSummary:
    """Summary statistics for validation results."""
    total_issues: int = 0
    critical_count: int = 0
    warning_count: int = 0
    info_count: int = 0
    empty_slides: int = 0
    slides_without_titles: int = 0
    missing_alt_text: int = 0
    low_contrast: int = 0
    large_images: int = 0
    fonts_used: int = 0
    colors_detected: int = 0
    bullet_violations: int = 0
    small_font_count: int = 0
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return asdict(self)


@dataclass
class ValidationPolicy:
    """Validation policy with thresholds."""
    name: str
    thresholds: Dict[str, Any]
    description: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return asdict(self)


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_policy(
    policy_name: str,
    custom_thresholds: Optional[Dict[str, Any]] = None
) -> ValidationPolicy:
    """
    Get validation policy by name with optional custom overrides.
    
    Args:
        policy_name: Name of policy ('lenient', 'standard', 'strict', 'custom')
        custom_thresholds: Optional custom threshold overrides
        
    Returns:
        ValidationPolicy instance
    """
    if policy_name == "custom" and custom_thresholds:
        base = VALIDATION_POLICIES["standard"]["thresholds"].copy()
        base.update(custom_thresholds)
        return ValidationPolicy(
            name="Custom",
            thresholds=base,
            description="Custom policy with user-defined thresholds"
        )
    
    config = VALIDATION_POLICIES.get(policy_name, VALIDATION_POLICIES["standard"])
    return ValidationPolicy(
        name=config["name"],
        thresholds=config["thresholds"],
        description=config.get("description", "")
    )


def generate_fix_command(
    filepath: Path,
    issue_type: str,
    slide_index: Optional[int] = None,
    shape_index: Optional[int] = None,
    extra_args: Optional[Dict[str, str]] = None
) -> Optional[str]:
    """
    Generate a CLI command to fix a specific issue.
    
    Args:
        filepath: Path to the presentation file
        issue_type: Type of issue to fix
        slide_index: Slide index if applicable
        shape_index: Shape index if applicable
        extra_args: Additional arguments for the fix command
        
    Returns:
        CLI command string or None if no fix available
    """
    base_path = str(filepath)
    
    fix_commands = {
        "missing_alt_text": (
            f"uv run tools/ppt_set_image_properties.py "
            f"--file \"{base_path}\" --slide {slide_index} --shape {shape_index} "
            f"--alt-text \"DESCRIBE_IMAGE_HERE\" --json"
        ),
        "empty_slide": (
            f"uv run tools/ppt_delete_slide.py "
            f"--file \"{base_path}\" --index {slide_index} --json"
        ),
        "missing_title": (
            f"uv run tools/ppt_set_title.py "
            f"--file \"{base_path}\" --slide {slide_index} "
            f"--title \"ADD_TITLE_HERE\" --json"
        ),
        "low_contrast": (
            f"uv run tools/ppt_format_text.py "
            f"--file \"{base_path}\" --slide {slide_index} --shape {shape_index} "
            f"--color \"#111111\" --json"
        ),
    }
    
    if issue_type in fix_commands:
        cmd = fix_commands[issue_type]
        if slide_index is None:
            return None
        return cmd
    
    return None


# ============================================================================
# VALIDATION PROCESSORS
# ============================================================================

def _process_core_validation(
    core_result: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """
    Process results from agent.validate_presentation().
    
    Args:
        core_result: Result from validate_presentation()
        issues: List to append issues to
        summary: Summary to update
        filepath: Path for fix commands
    """
    issue_data = core_result.get("issues", {})
    
    empty_slides = issue_data.get("empty_slides", [])
    summary.empty_slides = len(empty_slides)
    for idx in empty_slides:
        issues.append(ValidationIssue(
            category="structure",
            severity="critical",
            message=f"Empty slide with no content",
            slide_index=idx,
            fix_command=generate_fix_command(filepath, "empty_slide", slide_index=idx),
            details={"issue_type": "empty_slide"}
        ))
    
    untitled_slides = issue_data.get("slides_without_titles", [])
    summary.slides_without_titles = len(untitled_slides)
    for idx in untitled_slides:
        issues.append(ValidationIssue(
            category="structure",
            severity="warning",
            message=f"Slide missing title",
            slide_index=idx,
            fix_command=generate_fix_command(filepath, "missing_title", slide_index=idx),
            details={"issue_type": "missing_title"}
        ))
    
    fonts_used = issue_data.get("fonts_used", [])
    if isinstance(fonts_used, list):
        summary.fonts_used = len(fonts_used)


def _process_accessibility(
    acc_result: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """
    Process results from agent.check_accessibility().
    
    Args:
        acc_result: Result from check_accessibility()
        issues: List to append issues to
        summary: Summary to update
        filepath: Path for fix commands
    """
    issue_data = acc_result.get("issues", {})
    
    missing_alt = issue_data.get("missing_alt_text", [])
    summary.missing_alt_text = len(missing_alt)
    for item in missing_alt:
        slide_idx = item.get("slide", item.get("slide_index"))
        shape_idx = item.get("shape", item.get("shape_index"))
        issues.append(ValidationIssue(
            category="accessibility",
            severity="critical",
            message=f"Image missing alt text",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=generate_fix_command(
                filepath, "missing_alt_text",
                slide_index=slide_idx, shape_index=shape_idx
            ),
            details={
                "issue_type": "missing_alt_text",
                "shape_name": item.get("name", "Unknown")
            }
        ))
    
    low_contrast = issue_data.get("low_contrast", [])
    summary.low_contrast = len(low_contrast)
    for item in low_contrast:
        slide_idx = item.get("slide", item.get("slide_index"))
        shape_idx = item.get("shape", item.get("shape_index"))
        issues.append(ValidationIssue(
            category="accessibility",
            severity="warning",
            message=f"Low color contrast ratio ({item.get('ratio', 'N/A')})",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=generate_fix_command(
                filepath, "low_contrast",
                slide_index=slide_idx, shape_index=shape_idx
            ),
            details={
                "issue_type": "low_contrast",
                "contrast_ratio": item.get("ratio"),
                "wcag_minimum": 4.5
            }
        ))
    
    small_fonts = issue_data.get("small_fonts", [])
    summary.small_font_count = len(small_fonts)
    for item in small_fonts:
        issues.append(ValidationIssue(
            category="accessibility",
            severity="warning",
            message=f"Font size too small ({item.get('size', 'N/A')}pt)",
            slide_index=item.get("slide"),
            shape_index=item.get("shape"),
            details={
                "issue_type": "small_font",
                "font_size_pt": item.get("size"),
                "minimum_recommended": 12
            }
        ))


def _process_assets(
    asset_result: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """
    Process results from agent.validate_assets().
    
    Args:
        asset_result: Result from validate_assets()
        issues: List to append issues to
        summary: Summary to update
        filepath: Path for fix commands
    """
    issue_data = asset_result.get("issues", {})
    
    large_images = issue_data.get("large_images", [])
    summary.large_images = len(large_images)
    for item in large_images:
        issues.append(ValidationIssue(
            category="assets",
            severity="info",
            message=f"Large image may slow presentation ({item.get('size_mb', 'N/A')} MB)",
            slide_index=item.get("slide"),
            shape_index=item.get("shape"),
            details={
                "issue_type": "large_image",
                "size_mb": item.get("size_mb"),
                "recommended_max_mb": 2.0
            }
        ))
    
    missing_assets = issue_data.get("missing_assets", [])
    for item in missing_assets:
        issues.append(ValidationIssue(
            category="assets",
            severity="critical",
            message=f"Referenced asset not found: {item.get('name', 'Unknown')}",
            slide_index=item.get("slide"),
            details={
                "issue_type": "missing_asset",
                "asset_name": item.get("name")
            }
        ))


def _validate_design_rules(
    agent: PowerPointAgent,
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    policy: ValidationPolicy,
    filepath: Path
) -> None:
    """
    Validate design rules according to policy thresholds.
    
    Checks:
    - Font count limit
    - Color count limit  
    - 6x6 rule (bullets per slide, words per bullet)
    
    Args:
        agent: PowerPointAgent instance
        issues: List to append issues to
        summary: Summary to update
        policy: Validation policy with thresholds
        filepath: Path for fix commands
    """
    thresholds = policy.thresholds
    
    if summary.fonts_used > thresholds.get("max_fonts", 5):
        issues.append(ValidationIssue(
            category="design",
            severity="warning",
            message=f"Too many fonts used ({summary.fonts_used} > {thresholds.get('max_fonts', 5)})",
            details={
                "issue_type": "excessive_fonts",
                "font_count": summary.fonts_used,
                "threshold": thresholds.get("max_fonts", 5),
                "recommendation": "Limit to 2-3 font families for consistency"
            }
        ))
    
    try:
        presentation_info = agent.get_presentation_info()
        slide_count = presentation_info.get("slide_count", 0)
        
        colors_detected: Set[str] = set()
        bullet_violations = 0
        
        for slide_idx in range(slide_count):
            try:
                slide_info = agent.get_slide_info(slide_idx)
                shapes = slide_info.get("shapes", [])
                
                for shape in shapes:
                    if "fill_color" in shape and shape["fill_color"]:
                        colors_detected.add(shape["fill_color"])
                    if "line_color" in shape and shape["line_color"]:
                        colors_detected.add(shape["line_color"])
                    if "text_color" in shape and shape["text_color"]:
                        colors_detected.add(shape["text_color"])
                    
                    if thresholds.get("enforce_6x6_rule", False):
                        if shape.get("has_text_frame", False):
                            paragraphs = shape.get("paragraphs", [])
                            bullet_count = len([p for p in paragraphs if p.get("is_bullet", False)])
                            
                            if bullet_count > 6:
                                bullet_violations += 1
                                issues.append(ValidationIssue(
                                    category="design",
                                    severity="warning",
                                    message=f"Too many bullet points ({bullet_count} > 6)",
                                    slide_index=slide_idx,
                                    shape_index=shape.get("index"),
                                    details={
                                        "issue_type": "6x6_violation",
                                        "bullet_count": bullet_count,
                                        "max_allowed": 6
                                    }
                                ))
                            
            except Exception:
                continue
        
        summary.colors_detected = len(colors_detected)
        summary.bullet_violations = bullet_violations
        
        max_colors = thresholds.get("max_colors", 10)
        if summary.colors_detected > max_colors:
            issues.append(ValidationIssue(
                category="design",
                severity="warning",
                message=f"Too many colors used ({summary.colors_detected} > {max_colors})",
                details={
                    "issue_type": "excessive_colors",
                    "color_count": summary.colors_detected,
                    "threshold": max_colors,
                    "recommendation": "Limit to 3-5 primary colors for visual coherence"
                }
            ))
            
    except Exception:
        pass


def _check_policy_compliance(
    summary: ValidationSummary,
    policy: ValidationPolicy
) -> tuple:
    """
    Check if validation results comply with policy thresholds.
    
    Args:
        summary: Validation summary
        policy: Validation policy
        
    Returns:
        Tuple of (passed: bool, violations: List[str])
    """
    violations = []
    thresholds = policy.thresholds
    
    checks = [
        ("max_critical_issues", summary.critical_count, "Critical issues"),
        ("max_empty_slides", summary.empty_slides, "Empty slides"),
        ("max_slides_without_titles", summary.slides_without_titles, "Untitled slides"),
        ("max_missing_alt_text", summary.missing_alt_text, "Missing alt text"),
        ("max_low_contrast", summary.low_contrast, "Low contrast issues"),
        ("max_large_images", summary.large_images, "Large images"),
        ("max_fonts", summary.fonts_used, "Font families"),
        ("max_colors", summary.colors_detected, "Colors"),
    ]
    
    for threshold_key, actual_value, label in checks:
        threshold_value = thresholds.get(threshold_key)
        if threshold_value is not None and actual_value > threshold_value:
            violations.append(f"{label} ({actual_value}) exceeds limit ({threshold_value})")
    
    if thresholds.get("require_all_alt_text", False) and summary.missing_alt_text > 0:
        violations.append(f"All images must have alt text ({summary.missing_alt_text} missing)")
    
    return len(violations) == 0, violations


def _generate_recommendations(
    issues: List[ValidationIssue],
    policy: ValidationPolicy
) -> List[Dict[str, Any]]:
    """
    Generate prioritized recommendations based on issues found.
    
    Args:
        issues: List of validation issues
        policy: Validation policy
        
    Returns:
        List of recommendation dictionaries
    """
    recommendations = []
    
    critical_issues = [i for i in issues if i.severity == "critical"]
    accessibility_issues = [i for i in issues if i.category == "accessibility"]
    design_issues = [i for i in issues if i.category == "design"]
    
    if any(i.details.get("issue_type") == "empty_slide" for i in critical_issues):
        recommendations.append({
            "priority": "high",
            "category": "structure",
            "action": "Remove or populate empty slides",
            "impact": "Improves presentation flow and professionalism"
        })
    
    if any(i.details.get("issue_type") == "missing_alt_text" for i in accessibility_issues):
        recommendations.append({
            "priority": "high",
            "category": "accessibility",
            "action": "Add descriptive alt text to all images",
            "impact": "Required for WCAG 2.1 AA compliance and screen reader users"
        })
    
    if any(i.details.get("issue_type") == "low_contrast" for i in accessibility_issues):
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "action": "Improve text/background contrast ratios",
            "impact": "Ensures readability for users with visual impairments"
        })
    
    if any(i.details.get("issue_type") == "excessive_fonts" for i in design_issues):
        recommendations.append({
            "priority": "medium",
            "category": "design",
            "action": "Consolidate to 2-3 font families",
            "impact": "Creates visual consistency and professional appearance"
        })
    
    if any(i.details.get("issue_type") == "excessive_colors" for i in design_issues):
        recommendations.append({
            "priority": "low",
            "category": "design",
            "action": "Reduce color palette to 3-5 primary colors",
            "impact": "Improves visual coherence and brand consistency"
        })
    
    return recommendations


# ============================================================================
# MAIN VALIDATION FUNCTION
# ============================================================================

def validate_presentation(
    filepath: Path,
    policy: ValidationPolicy
) -> Dict[str, Any]:
    """
    Perform comprehensive presentation validation.
    
    Args:
        filepath: Path to PowerPoint file
        policy: Validation policy to apply
        
    Returns:
        Complete validation report dictionary
        
    Raises:
        FileNotFoundError: If file doesn't exist
        PowerPointAgentError: If validation fails
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    issues: List[ValidationIssue] = []
    summary = ValidationSummary()
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only validation, no lock needed
        
        presentation_info = agent.get_presentation_info()
        slide_count = presentation_info.get("slide_count", 0)
        presentation_version = agent.get_presentation_version()
        
        core_validation = agent.validate_presentation()
        accessibility_validation = agent.check_accessibility()
        asset_validation = agent.validate_assets()
        
        _process_core_validation(core_validation, issues, summary, filepath)
        _process_accessibility(accessibility_validation, issues, summary, filepath)
        _process_assets(asset_validation, issues, summary, filepath)
        _validate_design_rules(agent, issues, summary, policy, filepath)
    
    summary.total_issues = len(issues)
    summary.critical_count = sum(1 for i in issues if i.severity == "critical")
    summary.warning_count = sum(1 for i in issues if i.severity == "warning")
    summary.info_count = sum(1 for i in issues if i.severity == "info")
    
    passed, violations = _check_policy_compliance(summary, policy)
    
    if summary.critical_count > 0:
        status = "critical"
    elif not passed:
        status = "failed"
    elif summary.warning_count > 0:
        status = "warnings"
    else:
        status = "valid"
    
    return {
        "status": status,
        "passed": passed,
        "tool_version": __version__,
        "core_version": CORE_VERSION,
        "file": str(filepath.resolve()),
        "presentation_version": presentation_version,
        "validated_at": datetime.utcnow().isoformat() + "Z",
        "policy": policy.to_dict(),
        "summary": summary.to_dict(),
        "policy_violations": violations,
        "issues": [i.to_dict() for i in issues],
        "recommendations": _generate_recommendations(issues, policy),
        "presentation_info": {
            "slide_count": slide_count,
            "file_size_mb": presentation_info.get("file_size_mb"),
            "aspect_ratio": presentation_info.get("aspect_ratio"),
            "has_notes": presentation_info.get("has_notes", False)
        }
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Comprehensive PowerPoint presentation validation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Standard validation
  uv run tools/ppt_validate_presentation.py --file deck.pptx --json

  # Strict validation for production
  uv run tools/ppt_validate_presentation.py --file deck.pptx --policy strict --json

  # Custom thresholds
  uv run tools/ppt_validate_presentation.py --file deck.pptx \\
    --max-empty-slides 0 --max-missing-alt-text 0 --json

Policies:
  lenient  - Minimal validation for drafts
  standard - Balanced validation (default)
  strict   - Maximum validation for production

Validation Categories:
  structure     - Empty slides, missing titles
  accessibility - Alt text, contrast, font sizes
  assets        - Large images, missing files
  design        - Font/color limits, 6x6 rule
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to validate'
    )
    
    parser.add_argument(
        '--policy',
        choices=['lenient', 'standard', 'strict'],
        default='standard',
        help='Validation policy (default: standard)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.add_argument(
        '--max-missing-alt-text',
        type=int,
        metavar='N',
        help='Override maximum missing alt text allowed'
    )
    
    parser.add_argument(
        '--max-slides-without-titles',
        type=int,
        metavar='N',
        help='Override maximum untitled slides allowed'
    )
    
    parser.add_argument(
        '--max-empty-slides',
        type=int,
        metavar='N',
        help='Override maximum empty slides allowed'
    )
    
    parser.add_argument(
        '--require-all-alt-text',
        action='store_true',
        help='Require alt text on all images'
    )
    
    parser.add_argument(
        '--enforce-6x6',
        action='store_true',
        help='Enforce 6x6 content density rule'
    )
    
    parser.add_argument(
        '--summary-only',
        action='store_true',
        help='Output summary only, omit individual issues'
    )
    
    args = parser.parse_args()
    
    try:
        custom_thresholds = {}
        if args.max_missing_alt_text is not None:
            custom_thresholds["max_missing_alt_text"] = args.max_missing_alt_text
        if args.max_slides_without_titles is not None:
            custom_thresholds["max_slides_without_titles"] = args.max_slides_without_titles
        if args.max_empty_slides is not None:
            custom_thresholds["max_empty_slides"] = args.max_empty_slides
        if args.require_all_alt_text:
            custom_thresholds["require_all_alt_text"] = True
        if args.enforce_6x6:
            custom_thresholds["enforce_6x6_rule"] = True
        
        if custom_thresholds:
            policy = get_policy("custom", custom_thresholds)
        else:
            policy = get_policy(args.policy)
        
        result = validate_presentation(args.file.resolve(), policy)
        
        if args.summary_only:
            result.pop("issues", None)
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        
        exit_code = 1 if result["status"] in ("critical", "failed") else 0
        sys.exit(exit_code)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Check file integrity and PowerPoint format",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# schemas/capability_probe.v3.1.0.schema.json
```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "$id": "https://schemas.example.com/capability_probe/v1.1.1",
  "title": "PowerPoint capability probe output (v1.1.1)",
  "type": "object",
  "oneOf": [
    {
      "title": "Success payload",
      "type": "object",
      "required": ["status", "metadata", "slide_dimensions", "layouts", "theme", "capabilities", "warnings", "info"],
      "properties": {
        "status": { "const": "success" },
        "metadata": {
          "type": "object",
          "required": [
            "file",
            "probed_at",
            "tool_version",
            "schema_version",
            "operation_id",
            "deep_analysis",
            "atomic_verified",
            "duration_ms",
            "library_versions",
            "layout_count_total",
            "layout_count_analyzed"
          ],
          "properties": {
            "file": { "type": "string", "minLength": 1 },
            "probed_at": { "type": "string", "format": "date-time" },
            "tool_version": { "type": "string", "pattern": "^\\d+\\.\\d+\\.\\d+$" },
            "schema_version": { "type": "string", "pattern": "^capability_probe\\.v\\d+\\.\\d+\\.\\d+$" },
            "operation_id": { "type": "string", "pattern": "^[0-9a-fA-F-]{36}$" },
            "deep_analysis": { "type": "boolean" },
            "analysis_mode": { "type": "string", "enum": ["deep", "essential"] },
            "atomic_verified": { "type": "boolean" },
            "duration_ms": { "type": "integer", "minimum": 0 },
            "library_versions": {
              "type": "object",
              "required": ["python-pptx", "Pillow"],
              "properties": {
                "python-pptx": { "type": "string" },
                "Pillow": { "type": "string" }
              },
              "additionalProperties": true
            },
            "checksum": { "type": ["string", "null"], "pattern": "^[0-9a-f]{32}$" },
            "timeout_seconds": { "type": ["integer", "null"], "minimum": 0 },
            "layout_count_total": { "type": "integer", "minimum": 0 },
            "layout_count_analyzed": { "type": "integer", "minimum": 0 },
            "warnings_count": { "type": "integer", "minimum": 0 }
          },
          "additionalProperties": true
        },
        "slide_dimensions": {
          "type": "object",
          "required": [
            "width_inches",
            "height_inches",
            "width_emu",
            "height_emu",
            "width_pixels",
            "height_pixels",
            "aspect_ratio",
            "aspect_ratio_float",
            "dpi_estimate"
          ],
          "properties": {
            "width_inches": { "type": "number", "minimum": 0 },
            "height_inches": { "type": "number", "minimum": 0 },
            "width_emu": { "type": "integer", "minimum": 0 },
            "height_emu": { "type": "integer", "minimum": 0 },
            "width_pixels": { "type": "integer", "minimum": 0 },
            "height_pixels": { "type": "integer", "minimum": 0 },
            "aspect_ratio": { "type": "string", "minLength": 3 },
            "aspect_ratio_float": { "type": "number", "minimum": 0 },
            "dpi_estimate": { "type": "integer", "minimum": 1 }
          },
          "additionalProperties": false
        },
        "layouts": {
          "type": "array",
          "items": {
            "type": "object",
            "required": ["index", "original_index", "name", "placeholder_count", "master_index"],
            "properties": {
              "index": { "type": "integer", "minimum": 0 },
              "original_index": { "type": "integer", "minimum": 0 },
              "name": { "type": "string" },
              "placeholder_count": { "type": "integer", "minimum": 0 },
              "master_index": { "type": ["integer", "null"], "minimum": 0 },
              "placeholders": {
                "type": "array",
                "items": {
                  "type": "object",
                  "required": ["type", "type_code", "idx", "name", "position_source"],
                  "properties": {
                    "type": { "type": "string" },
                    "type_code": { "type": "integer" },
                    "idx": { "type": "integer", "minimum": 0 },
                    "name": { "type": "string" },
                    "position_source": { "enum": ["instantiated", "template", "error"] },
                    "position_inches": {
                      "type": "object",
                      "properties": {
                        "left": { "type": "number" },
                        "top": { "type": "number" }
                      },
                      "required": ["left", "top"],
                      "additionalProperties": false
                    },
                    "position_percent": {
                      "type": "object",
                      "properties": {
                        "left": { "type": "string", "pattern": "^\\d+(\\.\\d+)?%$" },
                        "top": { "type": "string", "pattern": "^\\d+(\\.\\d+)?%$" }
                      },
                      "required": ["left", "top"],
                      "additionalProperties": false
                    },
                    "position_emu": {
                      "type": "object",
                      "properties": {
                        "left": { "type": "integer", "minimum": 0 },
                        "top": { "type": "integer", "minimum": 0 }
                      },
                      "required": ["left", "top"],
                      "additionalProperties": false
                    },
                    "size_inches": {
                      "type": "object",
                      "properties": {
                        "width": { "type": "number", "minimum": 0 },
                        "height": { "type": "number", "minimum": 0 }
                      },
                      "required": ["width", "height"],
                      "additionalProperties": false
                    },
                    "size_percent": {
                      "type": "object",
                      "properties": {
                        "width": { "type": "string", "pattern": "^\\d+(\\.\\d+)?%$" },
                        "height": { "type": "string", "pattern": "^\\d+(\\.\\d+)?%$" }
                      },
                      "required": ["width", "height"],
                      "additionalProperties": false
                    },
                    "size_emu": {
                      "type": "object",
                      "properties": {
                        "width": { "type": "integer", "minimum": 0 },
                        "height": { "type": "integer", "minimum": 0 }
                      },
                      "required": ["width", "height"],
                      "additionalProperties": false
                    },
                    "error": { "type": "string" }
                  },
                  "additionalProperties": true
                }
              },
              "instantiation_complete": { "type": "boolean" },
              "placeholder_expected": { "type": "integer", "minimum": 0 },
              "placeholder_instantiated": { "type": "integer", "minimum": 0 },
              "placeholder_types": {
                "type": "array",
                "items": { "type": "string" }
              },
              "placeholder_map": {
                "type": "object",
                "additionalProperties": { "type": "integer", "minimum": 0 }
              }
            },
            "additionalProperties": true
          }
        },
        "theme": {
          "type": "object",
          "required": ["colors", "fonts"],
          "properties": {
            "colors": { "type": "object", "additionalProperties": { "type": "string" } },
            "fonts": { "type": "object", "additionalProperties": { "type": "string" } },
            "per_master": {
              "type": "array",
              "items": {
                "type": "object",
                "required": ["master_index", "colors", "fonts"],
                "properties": {
                  "master_index": { "type": "integer", "minimum": 0 },
                  "colors": { "type": "object", "additionalProperties": { "type": "string" } },
                  "fonts": { "type": "object", "additionalProperties": { "type": "string" } }
                },
                "additionalProperties": false
              }
            }
          },
          "additionalProperties": true
        },
        "capabilities": {
          "type": "object",
          "required": [
            "has_footer_placeholders",
            "has_slide_number_placeholders",
            "has_date_placeholders",
            "layouts_with_footer",
            "layouts_with_slide_number",
            "layouts_with_date",
            "total_layouts",
            "total_master_slides",
            "per_master",
            "footer_support_mode",
            "slide_number_strategy",
            "recommendations",
            "analysis_complete"
          ],
          "properties": {
            "has_footer_placeholders": { "type": "boolean" },
            "has_slide_number_placeholders": { "type": "boolean" },
            "has_date_placeholders": { "type": "boolean" },
            "layouts_with_footer": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "index": { "type": "integer", "minimum": 0 },
                  "original_index": { "type": "integer", "minimum": 0 },
                  "name": { "type": "string" },
                  "master_index": { "type": ["integer", "null"], "minimum": 0 }
                },
                "additionalProperties": true
              }
            },
            "layouts_with_slide_number": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "index": { "type": "integer", "minimum": 0 },
                  "original_index": { "type": "integer", "minimum": 0 },
                  "name": { "type": "string" },
                  "master_index": { "type": ["integer", "null"], "minimum": 0 }
                },
                "additionalProperties": true
              }
            },
            "layouts_with_date": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "index": { "type": "integer", "minimum": 0 },
                  "original_index": { "type": "integer", "minimum": 0 },
                  "name": { "type": "string" },
                  "master_index": { "type": ["integer", "null"], "minimum": 0 }
                },
                "additionalProperties": true
              }
            },
            "total_layouts": { "type": "integer", "minimum": 0 },
            "total_master_slides": { "type": "integer", "minimum": 0 },
            "per_master": {
              "type": "array",
              "items": {
                "type": "object",
                "required": [
                  "master_index",
                  "layout_count",
                  "has_footer_layouts",
                  "has_slide_number_layouts",
                  "has_date_layouts"
                ],
                "properties": {
                  "master_index": { "type": "integer", "minimum": 0 },
                  "layout_count": { "type": "integer", "minimum": 0 },
                  "has_footer_layouts": { "type": "integer", "minimum": 0 },
                  "has_slide_number_layouts": { "type": "integer", "minimum": 0 },
                  "has_date_layouts": { "type": "integer", "minimum": 0 }
                },
                "additionalProperties": false
              }
            },
            "footer_support_mode": { "enum": ["placeholder", "fallback_textbox"] },
            "slide_number_strategy": { "enum": ["placeholder", "textbox"] },
            "recommendations": { "type": "array", "items": { "type": "string" } },
            "analysis_complete": { "type": "boolean" }
          },
          "additionalProperties": true
        },
        "warnings": { "type": "array", "items": { "type": "string" } },
        "info": { "type": "array", "items": { "type": "string" } }
      },
      "additionalProperties": true
    },
    {
      "title": "Error payload",
      "type": "object",
      "required": ["status", "error", "error_type", "metadata", "warnings"],
      "properties": {
        "status": { "const": "error" },
        "error": { "type": "string" },
        "error_type": { "type": "string" },
        "metadata": {
          "type": "object",
          "required": ["file", "tool_version", "operation_id", "probed_at"],
          "properties": {
            "file": { "type": ["string", "null"] },
            "tool_version": { "type": "string", "pattern": "^\\d+\\.\\d+\\.\\d+$" },
            "operation_id": { "type": "string", "pattern": "^[0-9a-fA-F-]{36}$" },
            "probed_at": { "type": "string", "format": "date-time" }
          },
          "additionalProperties": true
        },
        "warnings": { "type": "array", "items": { "type": "string" } }
      },
      "additionalProperties": true
    }
  ]
}

```

# schemas/ppt_capability_probe.schema.json
```json
{
    "$schema": "http://json-schema.org/draft-07/schema#",
    "title": "ppt_capability_probe Output Schema",
    "type": "object",
    "required": [
        "tool_name",
        "tool_version",
        "schema_version",
        "file",
        "probe_timestamp",
        "capabilities"
    ],
    "properties": {
        "tool_name": {
            "type": "string"
        },
        "tool_version": {
            "type": "string"
        },
        "schema_version": {
            "type": "string"
        },
        "file": {
            "type": "string"
        },
        "probe_timestamp": {
            "type": "string",
            "format": "date-time"
        },
        "capabilities": {
            "type": "object",
            "required": [
                "can_read",
                "can_write",
                "layouts",
                "slide_dimensions"
            ],
            "properties": {
                "can_read": {
                    "type": "boolean"
                },
                "can_write": {
                    "type": "boolean"
                },
                "layouts": {
                    "type": "array",
                    "items": {
                        "type": "string"
                    }
                },
                "slide_dimensions": {
                    "type": "object",
                    "required": [
                        "width_pt",
                        "height_pt"
                    ],
                    "properties": {
                        "width_pt": {
                            "type": "number"
                        },
                        "height_pt": {
                            "type": "number"
                        }
                    }
                },
                "max_image_size_mb": {
                    "type": "number"
                },
                "supports_z_order": {
                    "type": "boolean"
                }
            },
            "additionalProperties": true
        },
        "warnings": {
            "type": "array",
            "items": {
                "type": "string"
            }
        },
        "metadata": {
            "type": "object",
            "additionalProperties": true
        }
    },
    "additionalProperties": false
}
```

# schemas/ppt_get_info.schema.json
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "ppt_get_info Output Schema",
  "type": "object",
  "required": ["tool_name", "tool_version", "schema_version", "file", "presentation_version", "slide_count", "slides"],
  "properties": {
    "tool_name": { "type": "string" },
    "tool_version": { "type": "string" },
    "schema_version": { "type": "string" },
    "file": { "type": "string" },
    "presentation_version": { "type": "string" },
    "slide_count": { "type": "integer", "minimum": 0 },
    "slides": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["index", "id", "layout", "shape_count"],
        "properties": {
          "index": { "type": "integer", "minimum": 0 },
          "id": { "type": "string" },
          "layout": { "type": "string" },
          "shape_count": { "type": "integer", "minimum": 0 },
          "notes": { "type": "string" }
        }
      }
    },
    "template_id": { "type": ["string", "null"] },
    "theme": {
      "type": ["object", "null"],
      "properties": {
        "primary_color": { "type": "string" },
        "accent_colors": {
          "type": "array",
          "items": { "type": "string" }
        },
        "font_family": { "type": "string" }
      },
      "additionalProperties": true
    },
    "metadata": { "type": "object", "additionalProperties": true }
  },
  "additionalProperties": false
}

```

# schemas/change_manifest.schema.json
```json
{
    "$schema": "http://json-schema.org/draft-07/schema#",
    "title": "Presentation Change Manifest",
    "type": "object",
    "required": [
        "manifest_id",
        "source_file",
        "work_copy",
        "operations",
        "created_by",
        "timestamp",
        "approval_token"
    ],
    "properties": {
        "manifest_id": {
            "type": "string"
        },
        "source_file": {
            "type": "string"
        },
        "work_copy": {
            "type": "string"
        },
        "created_by": {
            "type": "string"
        },
        "timestamp": {
            "type": "string",
            "format": "date-time"
        },
        "approval_token": {
            "type": "string"
        },
        "operations": {
            "type": "array",
            "items": {
                "type": "object",
                "required": [
                    "op_id",
                    "cmd",
                    "args",
                    "expected_effect"
                ],
                "properties": {
                    "op_id": {
                        "type": "string"
                    },
                    "cmd": {
                        "type": "string"
                    },
                    "args": {
                        "type": "object"
                    },
                    "expected_effect": {
                        "type": "string"
                    },
                    "rollback_cmd": {
                        "type": "string"
                    },
                    "critical": {
                        "type": "boolean"
                    }
                },
                "additionalProperties": false
            }
        },
        "diff_summary": {
            "type": "object",
            "properties": {
                "slides_added": {
                    "type": "integer"
                },
                "slides_removed": {
                    "type": "integer"
                },
                "shapes_added": {
                    "type": "integer"
                },
                "shapes_removed": {
                    "type": "integer"
                }
            },
            "additionalProperties": false
        },
        "notes": {
            "type": "string"
        }
    },
    "additionalProperties": false
}
```

