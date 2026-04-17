from pathlib import Path
from typing import List, Tuple

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"}

def get_image_files(folder_path: Path) -> List[Path]:
    """Return a sorted list of image file paths inside *folder_path* (non-recursive).
    """
    files = sorted(
        p for p in folder_path.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )
    return files

