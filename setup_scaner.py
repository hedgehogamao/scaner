"""Install dependencies and verified PaddleOCR model files into this folder."""
from pathlib import Path
import hashlib
import os
import subprocess
import sys
import tarfile
import tempfile
import urllib.request

ROOT = Path(__file__).resolve().parent
SOURCE = "https://paddle-model-ecology.bj.bcebos.com/paddlex/official_inference_model/paddle3.0.0"
MODELS = {
    "PP-OCRv5_mobile_det": "50446e5d01ac2a73d5319c89513281f6578414c888c602f9af13f93feefffc58",
    "PP-OCRv5_mobile_rec": "566b9512b34e34a9f0db54d87b51fa5a0b9ed2cf1ab7e49728cc0b8b5a64f414",
}
FILES = ("inference.json", "inference.pdiparams", "inference.yml")


def install_models(root):
    models = root / "models"
    models.mkdir(exist_ok=True)
    for name, expected in MODELS.items():
        dest = models / name
        if all((dest / f).is_file() for f in FILES):
            continue
        if dest.exists():
            raise RuntimeError(f"模型目录不完整：{dest}。请换到新解压的程序目录再运行安装。")
        with tempfile.TemporaryDirectory(dir=models, prefix="setup-") as scratch:
            scratch = Path(scratch)
            package = scratch / "model.tar"
            url = f"{SOURCE}/{name}_infer.tar"
            print(f"下载模型：{name}", flush=True)
            with urllib.request.urlopen(url, timeout=90) as response, package.open("wb") as target:
                sha = hashlib.sha256()
                while data := response.read(524288):
                    sha.update(data)
                    target.write(data)
            if sha.hexdigest() != expected:
                raise ValueError(f"{name} SHA-256 不符，已停止。")
            stage = scratch / "checked"
            stage.mkdir()
            with tarfile.open(package) as tar:
                for filename in FILES:
                    member = tar.getmember(f"{name}_infer/{filename}")
                    if not member.isfile():
                        raise ValueError(f"意外的模型文件类型：{filename}")
                    with tar.extractfile(member) as src:
                        (stage / filename).write_bytes(src.read())
            stage.rename(dest)


def install_dependencies(root):
    deps = root / ".packages"
    if deps.exists():
        if not all((deps/d).is_dir() for d in ("paddle", "paddlex", "paddleocr", "PIL", "openpyxl", "xlrd")):
            raise RuntimeError(f"依赖目录不完整：{deps}。请换到新解压的程序目录再安装。")
        return
    print("从 PyPI 安装锁定的 Windows 本地依赖，不修改系统 Python…", flush=True)
    with tempfile.TemporaryDirectory(dir=root, prefix="setup-") as scratch:
        stage = Path(scratch) / "pkgs"
        subprocess.run([sys.executable, "-m", "pip", "--isolated", "install", "--no-cache-dir",
                        "--index-url", "https://pypi.org/simple", "--target", str(stage),
                        "-r", str(root / "requirements.txt")], check=True)
        stage.rename(deps)


if __name__ == "__main__":
    install_dependencies(ROOT)
    install_models(ROOT)
    print("安装完成。双击 Run.cmd 启动。")
