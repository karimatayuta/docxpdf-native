from __future__ import annotations

from pathlib import Path

import pytest

from docxpdf_native.cli import main
from docxpdf_native.converter import Converter
from docxpdf_native.exceptions import InvalidDocxError, PdfGenerationError


def test_converter_rejects_output_through_symlinked_parent(tmp_path: Path) -> None:
    actual = tmp_path / "actual"
    actual.mkdir()
    linked = tmp_path / "linked"
    linked.symlink_to(actual, target_is_directory=True)

    with pytest.raises(PdfGenerationError, match="symbolic link"):
        Converter._validate_destination(b"docx", linked / "output.pdf")


def test_converter_rejects_hard_link_to_input_as_output(tmp_path: Path) -> None:
    source = tmp_path / "input.docx"
    source.write_bytes(b"docx")
    destination = tmp_path / "output.pdf"
    destination.hardlink_to(source)

    with pytest.raises(InvalidDocxError, match="different files"):
        Converter._validate_destination(source, destination)


@pytest.mark.parametrize(
    ("extra_arguments", "message"),
    [
        (("--layout-json", "output.pdf"), "must use different paths"),
        (("--diagnostics", "output.pdf"), "must use different paths"),
        (
            ("--diagnostics", "report.json", "--layout-json", "report.json"),
            "must use different paths",
        ),
    ],
)
def test_cli_rejects_colliding_output_paths(
    extra_arguments: tuple[str, ...],
    message: str,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    source = tmp_path / "input.docx"
    source.write_bytes(b"not parsed")
    arguments = [str(source), "--output", str(tmp_path / "output.pdf")]
    for option, name in zip(extra_arguments[::2], extra_arguments[1::2], strict=True):
        arguments.extend((option, str(tmp_path / name)))

    exit_code = main(arguments)

    assert exit_code == 2
    assert message in capsys.readouterr().err
