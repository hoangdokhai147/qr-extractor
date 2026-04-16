import logging
import urllib.request
from pathlib import Path
from typing import Tuple

import cv2
import numpy as np
from PIL import Image
from pyzbar.pyzbar import decode as pyzbar_decode

logger = logging.getLogger(__name__)

MODELS_DIR = Path(__file__).parent.parent.parent / "wechat_qr_models"
BASE_URL = "https://raw.githubusercontent.com/WeChatCV/opencv_3rdparty/wechat_qrcode/"

FILES = ["detect.prototxt", "detect.caffemodel", "sr.prototxt", "sr.caffemodel"]


def ensure_models_downloaded() -> bool:
    """Download WeChat QRCode CNN models if they don't exist."""
    MODELS_DIR.mkdir(parents=True, exist_ok=True)

    for filename in FILES:
        filepath = MODELS_DIR / filename
        if not filepath.exists():
            url = BASE_URL + filename
            logger.info("Downloading WeChat QR model: %s", filename)
            try:
                urllib.request.urlretrieve(url, filepath)
            except Exception as e:
                logger.error("Failed to download model %s: %s", filename, e)
                return False
    return True


def get_wechat_detector():
    """Initializes and returns cv2.wechat_qrcode_WeChatQRCode."""
    if not ensure_models_downloaded():
        raise RuntimeError(
            "Failed to obtain WeChat QR models. Check your internet connection."
        )

    detector = cv2.wechat_qrcode_WeChatQRCode(
        str(MODELS_DIR / "detect.prototxt"),
        str(MODELS_DIR / "detect.caffemodel"),
        str(MODELS_DIR / "sr.prototxt"),
        str(MODELS_DIR / "sr.caffemodel"),
    )
    return detector


def _parse_qr_bytes(raw_bytes: bytes) -> str:
    """Safely decode pyzbar bytes to str (UTF-8 priority)."""
    try:
        return raw_bytes.decode("utf-8")
    except UnicodeDecodeError:
        return raw_bytes.decode("latin-1")


def _scan_with_pyzbar(img_arr: np.ndarray) -> str | None:
    """Scan with standard pyzbar as fallback."""
    try:
        codes = pyzbar_decode(Image.fromarray(img_arr))
        if codes:
            return "; ".join(_parse_qr_bytes(c.data) for c in codes)
    except Exception:
        pass
    return None


class QRScanner:
    def __init__(self):
        try:
            self.wechat = get_wechat_detector()
            logger.info("WeChatQRCode AI detector loaded successfully!")
        except Exception as e:
            logger.warning(
                "Could not initialize WeChatQRCode: %s. Using pyzbar fallback.", e
            )
            self.wechat = None

    def scan_image(self, image_path: Path) -> Tuple[str, str]:
        """Reads image and performs multi-strategy detection.
        Returns: (decoded_text, status) where status is "success" or "fail".
        """
        img = cv2.imread(str(image_path))
        if img is None:
            logger.warning("Cannot read image %s", image_path.name)
            return "", "fail"

        result = self._try_wechat_strategies(img)
        if result:
            return result, "success"

        # If WeChat completely fails, try standard PyZbar multi-strategy
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        result = self._try_pyzbar_strategies(gray)
        if result:
            return result, "success"

        return "", "fail"

    def _try_wechat_strategies(self, img: np.ndarray) -> str | None:
        if not self.wechat:
            return None

        def run_wechat(arr):
            res, _ = self.wechat.detectAndDecode(arr)
            if res:
                # We return standard join if there's multiple QR codes
                return "; ".join(res)
            return None

        # Strategy 1: Base image and rotations
        for k in [0, 1, 2, 3]:
            rot_img = np.rot90(img, k=k) if k > 0 else img
            ans = run_wechat(rot_img)
            if ans:
                return ans

        # Strategy 2: Upscaling
        for scale in [2, 3]:
            scaled = cv2.resize(
                img, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC
            )
            for k in [0, 1, 2]:  # just limited rotations for scaled
                rot_img = np.rot90(scaled, k=k) if k > 0 else scaled
                ans = run_wechat(rot_img)
                if ans:
                    return ans

        # Strategy 3: Sliding window (for tiny QR on large image without threshold bias)
        h, w = img.shape[:2]
        window_size = 800
        step = 400

        # Slicing the image into blocks
        for y in range(0, h, step):
            for x in range(0, w, step):
                crop = img[y : min(y + window_size, h), x : min(x + window_size, w)]
                if crop.shape[0] < 100 or crop.shape[1] < 100:
                    continue

                # Check normal crop
                ans = run_wechat(crop)
                if ans:
                    return ans

                # Check upscaled crop
                big = cv2.resize(crop, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                ans = run_wechat(big)
                if ans:
                    return ans

                # Check upscaled rotated crops
                for k in [1, 2]:
                    ans = run_wechat(np.rot90(big, k=k))
                    if ans:
                        return ans

        return None

    def _try_pyzbar_strategies(self, gray: np.ndarray) -> str | None:
        """Fallback to original PyZbar pipeline for weird edge cases."""
        # Baseline native try
        rgb_approx = cv2.cvtColor(gray, cv2.COLOR_GRAY2RGB)
        ans = _scan_with_pyzbar(rgb_approx)
        if ans:
            return ans

        # Upscales
        for scale in [2, 3]:
            scaled = cv2.resize(
                gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC
            )
            ans = _scan_with_pyzbar(scaled)
            if ans:
                return ans

            # Rotations
            for k in [1, 2, 3]:
                ans = _scan_with_pyzbar(np.rot90(scaled, k=k))
                if ans:
                    return ans

            # Adaptive threshold
            adap = cv2.adaptiveThreshold(
                scaled, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
            )
            ans = _scan_with_pyzbar(adap)
            if ans:
                return ans

        return None
