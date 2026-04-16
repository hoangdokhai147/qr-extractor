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

def get_targets(root_path: Path) -> Tuple[List[Path], str]:
    """Auto-detect layout: nested vs flat.
    
    Returns:
        Tuple of (list of targets, string layout_type)
        Targets can be subfolders (nested) or the root directory itself (flat).
    """
    subfolders = sorted(p for p in root_path.iterdir() if p.is_dir() and p.name != 'wechat_qr_models')
    direct_images = get_image_files(root_path)

    if subfolders:
        return subfolders, "nested"
    elif direct_images:
        return [root_path], "flat"
    else:
        return [], "empty"
