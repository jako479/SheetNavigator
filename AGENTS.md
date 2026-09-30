# INSERT PROJECT NAME HERE

INSERT DESCRIPTION HERE. Source in `src/INSERT_PACKAGE_NAME/`, tests in
`tests/`.

## Working agreement

All decisions are the human's. Ask — never assume, infer, or pick an approach on
your own.

- Every open choice is a question first: approach, naming, structure, behavior,
  file layout, scope, edge cases.
- Ambiguity in a feature description means ask, never interpret.
- Do not merge. Report back and stop.

## Dev environment

- Python 3.12+ (`requires-python = ">=3.12"`). Do not use 3.13-only syntax —
  pyright is pinned to 3.12 and flags it. `.python-version` pins the local
  `.venv` to 3.12, the oldest supported version.
- `uv sync` builds `.venv` and installs everything, including the `dev`
  dependency group. `uv.lock` is committed, so the versions are exact.
- Add a dependency with `uv add <pkg>`, or `uv add --dev <pkg>` for tooling.
  Do not hand-edit the dependency lists in pyproject.toml.

## Build and test commands

From the repository root, one command per call, never chained:

      uv run pytest
      uv run ruff check .
      uv run ruff format .
      uv run pyright

All four green before you call it done. `pytest` measures coverage and fails
below the floor set in pyproject.toml.

Never run any of the four as a baseline check before you have changed
anything, including when setting up a worktree. A skill step that says to
is overridden by this rule.

## Code style

- Modern idioms: full type hints (built-in generics, `X | None`),
  `pathlib.Path` over `os.path`, dataclasses, f-strings, match statements
  where natural.
- Ruff formats at 88 columns, double quotes. Ruff lint and pyright (standard
  mode) must pass.
- Comments explain why, not what, except where the code is likely hard to read
  for a newcomer to the language.

## Testing

- All new code gets pytest tests. Write tests first when practical.
- Tests live in `tests/`.

## Docs

Update project meta as appropriate: STATUS.md, WORKLOG.md, CHANGELOG.md,
TODO.md, README.md. CHANGELOG and TODO entries are single-line when possible.

New entries in STATUS.md, WORKLOG.md, CHANGELOG.md and TODO.md go at the top of
their section, never mid-list or at the bottom.

## Commits and hand-off

- Commit messages are one line, no body. For a commit covering several big
  topics, use a broad single-line message plus a separate body with one `- `
  line per topic.
- Never mention Claude, Anthropic or any AI tool — commits, comments, docs.
- At the end, leave the worktree and return the session to the main checkout
  first. Then give the git commands in two groups, one command per code
  block: first the squash-merge, then its commit; second the worktree
  removal, then the branch deletion. Each block is copied and run on its
  own, so a failed merge is never followed by the cleanup.

## Writing

This applies to every response, always: questions, back-and-forth discussion,
task results, and everything you write.

- Documentation: very simple, clear, and very high-level.
- Responses: very simple, clear, and very high-level — one sentence whenever
  possible.
- Questions: answer directly. Never assume, and never change code when asked a
  question.
- Yes/No questions: answer Yes or No, nothing else.
- Other questions: one sentence per question, whenever possible.
- Broad topics and discussion: one sentence per topic, whenever possible.
- Tasks: list the changes made, one short high-level line each.

No filler, no rambling.
