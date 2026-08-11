"""
Pandoc integration helper.

Locates a pandoc executable, runs it as a subprocess to convert
Markdown (.md) -> DOCX. Pandoc's built-in texmath translates LaTeX
math into native Word OMML, producing real editable equation objects
(not images, not plain text).
"""
import os
import shutil
import subprocess
from typing import Optional, List


# Known install locations on Windows; checked in order.
_PANDOC_CANDIDATES_WINDOWS: List[str] = [
    os.path.join(
        os.environ.get('LOCALAPPDATA', ''),
        r'Microsoft\WinGet\Packages\JohnMacFarlane.Pandoc_Microsoft.Winget.Source_8wekyb3d8bbwe',
    ),
    os.path.join(os.environ.get('ProgramFiles', r'C:\Program Files'), 'Pandoc'),
    r'C:\Program Files\Pandoc',
]


def find_pandoc() -> Optional[str]:
    """Return absolute path to pandoc.exe, or None if not found."""
    # 1. PATH lookup
    on_path = shutil.which('pandoc')
    if on_path:
        return on_path
    # 2. Known Windows install locations (WinGet package may sit in a
    # versioned subfolder).
    for base in _PANDOC_CANDIDATES_WINDOWS:
        if not base or not os.path.isdir(base):
            continue
        # Search one level deep for the actual exe.
        for entry in os.listdir(base):
            cand = os.path.join(base, entry, 'pandoc.exe')
            if os.path.isfile(cand):
                return cand
        cand = os.path.join(base, 'pandoc.exe')
        if os.path.isfile(cand):
            return cand
    return None


def is_pandoc_available() -> bool:
    return find_pandoc() is not None


def md_to_docx_via_pandoc(md_path: str, docx_path: str,
                          pandoc_exe: Optional[str] = None,
                          extra_args: Optional[List[str]] = None,
                          timeout: int = 120) -> str:
    """
    Convert a Markdown file to DOCX using pandoc. The output document
    will contain native Word OMML equation objects (oMath / oMathPara)
    for any LaTeX math in the source.

    Returns the absolute path to the produced DOCX.
    Raises RuntimeError on failure.
    """
    exe = pandoc_exe or find_pandoc()
    if not exe:
        raise RuntimeError('pandoc executable not found')

    md_path = os.path.abspath(md_path)
    docx_path = os.path.abspath(docx_path)
    os.makedirs(os.path.dirname(docx_path) or '.', exist_ok=True)

    cmd = [exe, md_path, '-o', docx_path]
    if extra_args:
        cmd.extend(extra_args)

    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            encoding='utf-8',
            errors='replace',
        )
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f'pandoc timed out after {timeout}s') from e
    except FileNotFoundError as e:
        raise RuntimeError(f'pandoc executable not usable: {exe}') from e

    if proc.returncode != 0:
        msg = (proc.stderr or proc.stdout or '').strip()
        raise RuntimeError(f'pandoc failed (code {proc.returncode}): {msg}')

    if not os.path.isfile(docx_path):
        raise RuntimeError('pandoc reported success but did not produce the docx file')

    return docx_path
