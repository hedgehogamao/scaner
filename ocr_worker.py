"""Offline OCR process. The caller supplies an already cropped local PNG."""
from pathlib import Path
import json
import os
import sys

ROOT = Path(__file__).resolve().parent
MODELS = ("PP-OCRv5_mobile_det", "PP-OCRv5_mobile_rec")


def infer(image, model_dir=ROOT / "models"):
    for name in MODELS:
        if not (model_dir / name / "inference.pdiparams").is_file():
            raise FileNotFoundError(f"缺少本地模型 {name}；请运行 Install.cmd。")
    os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "1"
    os.environ["HF_HUB_OFFLINE"] = "1"
    # This executable is a child process. Keep Paddle's hard-coded dataset cache
    # out of the user's profile without changing the app or system environment.
    cache = ROOT / ".cache"
    profile = cache / "profile"
    profile.mkdir(parents=True, exist_ok=True)
    os.environ["PADDLE_PDX_CACHE_HOME"] = str(cache / "paddlex")
    os.environ["USERPROFILE"] = str(profile)
    os.environ["HOME"] = str(profile)
    import numpy as np
    from PIL import Image
    from paddleocr import PaddleOCR

    engine = PaddleOCR(
        text_detection_model_name=MODELS[0],
        text_detection_model_dir=str(model_dir / MODELS[0]),
        text_recognition_model_name=MODELS[1],
        text_recognition_model_dir=str(model_dir / MODELS[1]),
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        enable_mkldnn=False,
        cpu_threads=4,
        device="cpu",
    )
    with Image.open(image) as photo:
        bgr = np.asarray(photo.convert("RGB"))[:, :, ::-1].copy()
    words = []
    for result in engine.predict(bgr, text_det_limit_type="max", text_det_limit_side_len=2200, text_rec_score_thresh=0.12):
        data = result.json["res"]
        for label, score, rect in zip(data["rec_texts"], data["rec_scores"], data["rec_boxes"], strict=True):
            if label.strip():
                words.append({"text": label, "score": float(score), "rect": [float(x) for x in rect]})
    return words


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: ocr_worker.py INPUT.png OUTPUT.json")
    result = infer(sys.argv[1])
    Path(sys.argv[2]).write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
