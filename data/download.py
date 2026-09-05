import hashlib
import json
import sys
import tarfile
import urllib.request
from pathlib import Path

URL = "https://huggingface.co/datasets/KAKA22/SpreadsheetBench/resolve/main/spreadsheetbench_verified_400.tar.gz?download=true"
SHA256 = "10ef893dd29cb13ab97143ea787e68cdc9574a13873ab9a54e50b31dc03fc949"
HERE = Path(__file__).resolve().parent
TARBALL = HERE / "spreadsheetbench_verified_400.tar.gz"
TARGET = HERE / "spreadsheetbench_verified_400"


def main():
    # Checks if dataset exists
    if (TARGET / "dataset.json").exists():
        print(f"already present: {TARGET}")
        return

    if not TARBALL.exists():
        print("downloading 15 MB from Hugging Face")
        urllib.request.urlretrieve(URL, TARBALL)

    digest = hashlib.sha256(TARBALL.read_bytes()).hexdigest()
    if digest != SHA256:
        sys.exit(f"checksum mismatch for {TARBALL}: {digest}")

    print(f"{TARBALL.name}: OK")

    with tarfile.open(TARBALL) as tar:
        tar.extractall(HERE, filter="data")

    n = len(json.loads((TARGET / "dataset.json").read_text()))
    print(f"{n} tasks in {TARGET}")


if __name__ == "__main__":
    main()
