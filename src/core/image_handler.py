"""
Image Handler for Pic2Doc
Handles image file operations and validation
"""

import os
from pathlib import Path
from typing import List, Optional, Tuple
from PIL import Image
from ..utils.constants import SUPPORTED_IMAGE_EXTENSIONS, LANDSCAPE_RATIO, PORTRAIT_RATIO


class ImageOrientation:
    """Image orientation types"""
    LANDSCAPE = "landscape"
    PORTRAIT = "portrait"
    SQUARE = "square"


class ImageHandler:
    """Handles image file operations"""

    def __init__(self, image_folder: str):
        """
        Initialize image handler

        Args:
            image_folder: Path to folder containing images
        """
        self.image_folder = Path(image_folder)

        if not self.image_folder.exists():
            raise FileNotFoundError(f"Bilder-Ordner nicht gefunden: {self.image_folder}")

    def find_image(self, filename_without_ext: str) -> Optional[Path]:
        """
        Find image file by filename (without extension)

        Tries common image extensions (.jpg, .jpeg, .png, .bmp)

        Args:
            filename_without_ext: Filename without extension (e.g., "FIAS 21B 000001")

        Returns:
            Path to image file if found, None otherwise
        """
        # Try each supported extension
        for ext in SUPPORTED_IMAGE_EXTENSIONS:
            image_path = self.image_folder / f"{filename_without_ext}{ext}"
            if image_path.exists():
                return image_path

        return None

    def validate_images(self, filenames: List[str]) -> tuple[List[str], List[str]]:
        """
        Validate that image files exist for given filenames

        Args:
            filenames: List of filenames without extensions

        Returns:
            Tuple of (found_files, missing_files)
        """
        found = []
        missing = []

        for filename in filenames:
            image_path = self.find_image(filename)
            if image_path:
                found.append(str(image_path))
            else:
                missing.append(filename)

        return found, missing

    def get_image_path(self, filename_without_ext: str) -> str:
        """
        Get full path to image file

        Args:
            filename_without_ext: Filename without extension

        Returns:
            Full path to image file

        Raises:
            FileNotFoundError: If image file not found
        """
        image_path = self.find_image(filename_without_ext)

        if not image_path:
            raise FileNotFoundError(
                f"Bilddatei nicht gefunden: {filename_without_ext} "
                f"(versuchte Endungen: {', '.join(SUPPORTED_IMAGE_EXTENSIONS)})"
            )

        return str(image_path)

    def get_image_dimensions(self, image_path: str) -> Tuple[int, int]:
        """
        Get image dimensions

        Args:
            image_path: Path to image file

        Returns:
            Tuple of (width, height) in pixels

        Raises:
            ValueError: If image cannot be read or is corrupted
        """
        try:
            with Image.open(image_path) as img:
                img.verify()  # Verify it's a valid image
            # Reopen after verify (verify closes the file)
            with Image.open(image_path) as img:
                return img.size
        except Exception as e:
            raise ValueError(f"Bilddatei beschädigt oder ungültig: {e}")

    def get_image_orientation(self, image_path: str) -> str:
        """
        Detect image orientation based on aspect ratio

        Args:
            image_path: Path to image file

        Returns:
            Orientation: "landscape", "portrait", or "square"
        """
        width, height = self.get_image_dimensions(image_path)
        aspect_ratio = width / height

        if aspect_ratio > LANDSCAPE_RATIO:
            return ImageOrientation.LANDSCAPE
        elif aspect_ratio < PORTRAIT_RATIO:
            return ImageOrientation.PORTRAIT
        else:
            return ImageOrientation.SQUARE

    def get_image_info(self, filename_without_ext: str) -> dict:
        """
        Get comprehensive image information

        Args:
            filename_without_ext: Filename without extension

        Returns:
            Dictionary with image info (path, dimensions, orientation)
        """
        image_path = self.get_image_path(filename_without_ext)
        width, height = self.get_image_dimensions(image_path)
        orientation = self.get_image_orientation(image_path)

        return {
            'path': image_path,
            'width': width,
            'height': height,
            'orientation': orientation,
            'aspect_ratio': width / height
        }
