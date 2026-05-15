# Contributing

Thanks for your interest in contributing!

## Development setup

```sh
git clone https://github.com/trobichaux/m365-service-comms-mcp.git
cd m365-service-comms-mcp
python -m venv .venv
# Windows: .\.venv\Scripts\Activate.ps1
# Unix:    source .venv/bin/activate
pip install -e ".[dev]"
```

## Required checks before opening a PR

```sh
ruff check .
ruff format --check .
pytest --cov
```

CI runs these on Python 3.11, 3.12, and 3.13. Please make sure they pass
locally first.

## End-to-end verification

The fastest sanity check (no tenant required):

```sh
m365-svc-comms-mcp --demo
```

Then connect any MCP client to it via stdio. The three tools should be
discoverable and callable, returning the canned demo data.

For a real tenant test, see the [README's Quickstart section](README.md#quickstart).

## Commit style

Use [Conventional Commits](https://www.conventionalcommits.org/) — e.g.,
`feat: add list_service_issues tool` or `fix(auth): retry token refresh on
expired cache`.

## Release process (maintainer only)

The `publish.yml` workflow uses
[PyPI trusted publishing](https://docs.pypi.org/trusted-publishers/) (OIDC,
no long-lived API tokens). One-time setup:

1. On <https://pypi.org/manage/account/publishing/>, add a trusted publisher
   for the package `m365-service-comms-mcp` with:
   - Owner: `trobichaux`
   - Repository name: `m365-service-comms-mcp`
   - Workflow filename: `publish.yml`
   - Environment name: `pypi`
2. Repeat at <https://test.pypi.org/manage/account/publishing/> with environment
   `testpypi` for dry-run testing.
3. In the GitHub repo, create the `pypi` and `testpypi` environments under
   **Settings → Environments**.

Per-release flow:

1. Bump the version in `pyproject.toml` and add a section to `CHANGELOG.md`.
2. Open a PR, get it reviewed/merged.
3. (Optional) Dry-run via **Actions → Publish to PyPI → Run workflow → target=testpypi**.
   Verify the package installs from TestPyPI:
   `uvx --index-url https://test.pypi.org/simple/ m365-service-comms-mcp --version`
4. Tag and push: `git tag v0.1.0 && git push origin v0.1.0`. The tag push
   triggers a real PyPI publish via the workflow.
5. Create a GitHub Release pointing at the tag with the `CHANGELOG.md`
   content for that version.

## Adding a new tool

1. Create `src/m365_service_comms_mcp/tools/<name>.py` exposing a
   `register(mcp, client)` function.
2. Add a test file `tests/test_<name>.py`.
3. Wire it into `src/m365_service_comms_mcp/tools/__init__.py`.
4. Update the `## Tools reference` section of `README.md`.
5. Update `CHANGELOG.md` under `[Unreleased]`.

## Security issues

See [SECURITY.md](SECURITY.md) — do not file security issues in the public
tracker.
