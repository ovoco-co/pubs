import shutil
import subprocess
import sys
from pathlib import Path


SUPPORTED_PANDOC_OUT_FORMAT = "markdown"
PPT_EXTENSION = ".ppt"
PPTX_EXTENSION = ".pptx"


def ps_single_quoted(s: str) -> str:
    # PowerShell single-quoted string escaping: ' -> ''
    return "'" + s.replace("'", "''") + "'"


def ppt_to_pptx_with_powershell_com(ppt_path: Path, pptx_path: Path) -> None:
    ppt_path_ps = ps_single_quoted(str(ppt_path))
    pptx_path_ps = ps_single_quoted(str(pptx_path))

    # Note: PowerPoint COM automation requires STA.
    ps = rf"""
$pptPath = {ppt_path_ps};
$pptxPath = {pptx_path_ps};
$pptApp = $null;
$pres = $null;
try {{
  $pptApp = New-Object -ComObject PowerPoint.Application;
  $pptApp.Visible = $false;
  $pptApp.DisplayAlerts = 0;
  $pres = $pptApp.Presentations.Open($pptPath);
  # 24 = ppSaveAsOpenXMLPresentation
  $pres.SaveAs($pptxPath, 24);
}}
finally {{
  if ($pres -ne $null) {{
    try {{ $pres.Close() | Out-Null }} catch {{ }}
    try {{ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }} catch {{ }}
  }}
  if ($pptApp -ne $null) {{
    try {{ $pptApp.Quit() | Out-Null }} catch {{ }}
    try {{ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) }} catch {{ }}
  }}
  [GC]::Collect();
  [GC]::WaitForPendingFinalizers();
}}
""".strip()

    cmd = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-STA",
        "-Command",
        ps,
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            "PowerPoint COM conversion failed:\n"
            f"stdout:\n{result.stdout}\n\nstderr:\n{result.stderr}".strip()
        )

    if not pptx_path.exists():
        raise RuntimeError(
            f"PowerPoint COM conversion did not produce file: {pptx_path}"
        )


def ppt_to_pptx(ppt_path: Path, pptx_path: Path, temp_dir: Path) -> None:
    # Prefer LibreOffice if present; it can convert .ppt -> .pptx headlessly.
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if soffice:
        try:
            cmd = [
                soffice,
                "--headless",
                "--convert-to",
                PPTX_EXTENSION.lstrip("."),
                "--outdir",
                str(temp_dir),
                str(ppt_path),
            ]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0 and pptx_path.exists():
                return
        except Exception:
            # Fall back to COM conversion below.
            pass

    # Fallback: PowerPoint COM automation (requires Microsoft PowerPoint installed).
    ppt_to_pptx_with_powershell_com(ppt_path, pptx_path)


def convert_with_pandoc_pptx(
    pptx_path: Path,
    output_md: Path,
    *,
    images_root: Path,
    workdir: Path,
) -> None:
    stem = pptx_path.stem
    extract_media_rel = f"images/{stem}"

    cmd = [
        "pandoc",
        str(pptx_path),
        "-o",
        str(output_md),
        "--wrap=none",
        f"--extract-media={extract_media_rel}",
    ]

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        cwd=str(workdir),
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"Pandoc failed for {pptx_path.name}:\n{result.stderr.strip()}"
        )

    if not any(p.is_file() for p in (images_root / stem).rglob("*")):
        # Not fatal—some decks may have no embedded bitmaps.
        print(f"[warn] No extracted images for: {pptx_path.name}")


def main() -> None:
    if len(sys.argv) != 3:
        print("Usage: python folder_ppt_to_md.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_folder = Path(sys.argv[2]).resolve()
    if not input_folder.is_dir():
        print(f"Input folder does not exist: {input_folder}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)
    images_root = output_folder / "images"
    images_root.mkdir(parents=True, exist_ok=True)

    temp_dir = output_folder / ".tmp_pptx"
    temp_dir.mkdir(parents=True, exist_ok=True)

    ppt_files = sorted(
        [p for p in input_folder.iterdir() if p.is_file() and p.suffix.lower() == PPT_EXTENSION]
    )
    if not ppt_files:
        print(f"No {PPT_EXTENSION} files found in: {input_folder}")
        sys.exit(0)

    for ppt_path in ppt_files:
        stem = ppt_path.stem
        output_md = output_folder / f"{stem}.md"
        pptx_path = temp_dir / f"{stem}{PPTX_EXTENSION}"

        print(f"Converting PPT -> PPTX: {ppt_path.name}")
        if pptx_path.exists():
            pptx_path.unlink()
        ppt_to_pptx(ppt_path, pptx_path, temp_dir)

        print(f"Converting PPTX -> MD (with extracted images): {ppt_path.name}")
        convert_with_pandoc_pptx(
            pptx_path,
            output_md,
            images_root=images_root,
            workdir=output_folder,
        )

    print("Done.")


if __name__ == "__main__":
    main()

