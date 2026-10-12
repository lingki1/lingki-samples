"""把旧版 .doc 转成 .docx。在临时文件夹里对副本操作，原件不动。"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

SOFFICE_ENV = "HOMEWORK_SOFFICE"
SOFFICE_LOCATIONS = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)
TIMEOUT_SECONDS = 180
NOT_CONVERTED = "LibreOffice 没能转换这个文件，可能已损坏；请用 Word 另存为 .docx 后再交"


@dataclass
class ConversionResult:
    converted: dict[Path, bytes]
    failed: dict[Path, str]


def find_soffice() -> Path | None:
    # 固定的安装位置排在"到处找一找"前面：后者在 Windows 上会先看当前文件夹，
    # 不能让作业文件夹旁边一个碰巧叫 soffice 的东西被当成 LibreOffice 运行
    for candidate in (os.environ.get(SOFFICE_ENV), *SOFFICE_LOCATIONS):
        if candidate and Path(candidate).is_file():
            return Path(candidate)
    found = shutil.which("soffice")
    return Path(found) if found and Path(found).is_file() else None


def _run_soffice(command: list[str], timeout: int) -> str | None:
    """运行 LibreOffice，等它结束。出问题时返回原因，正常返回 None。

    soffice.exe 只是个启动器，真正干活的是它启动的 soffice.bin。所以超时后要连子进程一起结束，
    而且不接它的输出——接了的话，子进程不退出，这边就会一直等下去。
    """
    try:
        process = subprocess.Popen(
            command, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
    except OSError as error:
        return f"LibreOffice 启动失败（{error}）"
    try:
        process.wait(timeout=timeout)
        return None
    except subprocess.TimeoutExpired:
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/T", "/F", "/PID", str(process.pid)],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        else:
            process.kill()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            pass
        return f"LibreOffice 转换超过 {timeout} 秒没有完成"


def convert_doc_files(paths: Sequence[Path]) -> ConversionResult:
    paths = list(paths)
    if not paths:
        return ConversionResult({}, {})
    soffice = find_soffice()
    if soffice is None:
        reason = "这台电脑没有装 LibreOffice，无法转换旧版 .doc；请用 Word 另存为 .docx 后再交"
        return ConversionResult({}, {p: reason for p in paths})

    # 清理临时文件夹时如果 LibreOffice 还占着里面的文件，清不掉就算了，不能因此让整批中断
    with tempfile.TemporaryDirectory(prefix="homework_doc_", ignore_cleanup_errors=True) as tmp:
        tmp_path = Path(tmp)
        in_dir, out_dir, profile = tmp_path / "in", tmp_path / "out", tmp_path / "profile"
        in_dir.mkdir()
        out_dir.mkdir()
        # 副本用数字命名：避开中文路径传给外部程序时的编码问题
        copies: dict[Path, Path] = {}
        result = ConversionResult({}, {})
        for index, path in enumerate(paths):
            copy = in_dir / f"{index}.doc"
            try:
                shutil.copyfile(path, copy)
                copies[path] = copy
            except OSError as error:
                result.failed[path] = f"读不出来，可能正被别的程序打开（{type(error).__name__}）"

        def convert(batch: list[Path]) -> str | None:
            command = [
                str(soffice),
                # 独立的配置目录：老师正开着 LibreOffice 时也不会互相干扰
                f"-env:UserInstallation={profile.as_uri()}",
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                str(out_dir),
                *[str(copies[p]) for p in batch],
            ]
            return _run_soffice(command, TIMEOUT_SECONDS)

        def collect(batch: list[Path]) -> list[Path]:
            missing = []
            for path in batch:
                output = out_dir / f"{copies[path].stem}.docx"
                if output.is_file() and output.stat().st_size > 0:
                    result.converted[path] = output.read_bytes()
                else:
                    missing.append(path)
            return missing

        pending = list(copies)
        failure = convert(pending) if pending else None
        missing = collect(pending)
        if missing and len(pending) > 1:
            # 一起转时有没出来的：可能只是其中一份有问题连累了后面的，逐份再试一次，免得冤枉好文件
            for path in list(missing):
                failure = convert([path])
                if not collect([path]):
                    missing.remove(path)
                else:
                    result.failed[path] = failure or NOT_CONVERTED
        else:
            for path in missing:
                result.failed[path] = failure or NOT_CONVERTED
        return result
