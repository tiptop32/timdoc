from pathlib import Path
from typing import Any

import pytest
import yaml

ROOT = Path(__file__).parents[1]
WORKFLOW = ROOT / ".github" / "workflows" / "windows-build.yml"
README = ROOT / "README.md"
ASSETS = {"dist/Timdoc-Windows.zip", "dist/Timdoc-Setup.exe"}


def _step_with(steps: list[dict[str, Any]], key: str, value: str) -> dict[str, Any]:
    matches = [step for step in steps if value in step.get(key, "")]
    assert len(matches) == 1, f"expected one step with {key} containing {value!r}"
    return matches[0]


def _validate_release_contract(workflow_text: str, readme: str) -> None:
    workflow = yaml.load(workflow_text, Loader=yaml.BaseLoader)
    assert workflow["on"]["push"]["tags"] == ["v*"]
    assert "workflow_dispatch" in workflow["on"]
    assert workflow["permissions"] == {"contents": "read"}

    build = workflow["jobs"]["build"]
    assert build["runs-on"] == "windows-latest"
    assert "permissions" not in build
    build_steps = build["steps"]
    runs = [step["run"] for step in build_steps if "run" in step]
    assert runs.index("uv run pytest") < runs.index(
        "uv run pyinstaller packaging/timdoc.spec --clean --noconfirm"
    )
    assert runs.index("uv run pyinstaller packaging/timdoc.spec --clean --noconfirm") < runs.index(
        "dist\\Timdoc\\Timdoc.exe --self-test"
    )

    upload = _step_with(build_steps, "uses", "actions/upload-artifact@v4")
    assert upload["with"]["name"] == "Timdoc-Windows"
    assert set(upload["with"]["path"].splitlines()) == ASSETS
    assert upload["with"]["if-no-files-found"] == "error"

    release = workflow["jobs"]["release"]
    assert release.get("needs") == "build"
    assert release["if"] == ("github.event_name == 'push' && startsWith(github.ref, 'refs/tags/v')")
    assert release["permissions"] == {"contents": "write"}
    release_steps = release["steps"]
    download = _step_with(release_steps, "uses", "actions/download-artifact@v5")
    assert download["with"] == {"name": "Timdoc-Windows", "path": "dist"}

    publish = _step_with(release_steps, "run", "gh release create")
    assert publish["env"] == {"GH_TOKEN": "${{ github.token }}"}
    command = publish["run"]
    assert command.startswith('gh release create "$GITHUB_REF_NAME" ')
    assert ASSETS <= set(command.split())
    assert all(option in command for option in ('--repo "$GITHUB_REPOSITORY"', "--verify-tag"))
    assert "--generate-notes" in command

    assert "git tag -a v0.1.0" in readme
    assert "git push origin v0.1.0" in readme
    assert all(asset.removeprefix("dist/") in readme for asset in ASSETS)
    assert "Ручной запуск workflow" in readme and "не публикует Release" in readme


@pytest.mark.eval
def test_tagged_windows_build_publishes_installers_to_github_releases() -> None:
    _validate_release_contract(
        WORKFLOW.read_text(encoding="utf-8"),
        README.read_text(encoding="utf-8"),
    )


@pytest.mark.parametrize(
    ("old", "new"),
    [
        ("    needs: build\n", ""),
        ("actions/download-artifact@v5", "actions/checkout@v4"),
        ("dist/Timdoc-Setup.exe", "dist/Missing.exe"),
        (
            "    runs-on: windows-latest\n",
            "    runs-on: windows-latest\n    permissions:\n      contents: write\n",
        ),
        ('gh release create "$GITHUB_REF_NAME"', 'gh release create "v0.0.0"'),
    ],
)
def test_release_contract_rejects_incomplete_workflows(old: str, new: str) -> None:
    workflow = WORKFLOW.read_text(encoding="utf-8").replace(old, new)
    with pytest.raises(AssertionError):
        _validate_release_contract(workflow, README.read_text(encoding="utf-8"))
