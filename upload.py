"""Build the package and optionally upload the freshly built artifacts."""
from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path


def upload_app(build_only: bool = False) -> None:
    if build_only:
        output_dir = Path("dist")
        output_dir.mkdir(exist_ok=True)
        subprocess.run(
            [sys.executable, "-m", "build", "--outdir", str(output_dir)],
            check=True,
        )
        return

    # Isolating the upload build prevents stale distributions in ./dist from
    # being uploaded alongside the current version.
    with tempfile.TemporaryDirectory(prefix="capiq-build-") as temp_dir:
        subprocess.run(
            [sys.executable, "-m", "build", "--outdir", temp_dir],
            check=True,
        )
        artifacts = sorted(str(path) for path in Path(temp_dir).iterdir())
        if not artifacts:
            raise RuntimeError("Package build produced no artifacts")
        subprocess.run(
            [sys.executable, "-m", "twine", "upload", *artifacts],
            check=True,
        )


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Build and upload capiq-excel")
    parser.add_argument(
        "--build-only",
        action="store_true",
        help="Build into ./dist without uploading",
    )
    args = parser.parse_args()
    upload_app(build_only=args.build_only)
