# Sheet Navigator

Excel VSTO add-in that shows a Worksheets pane for any workbook, one pane per
Excel window, with the pane's open state and width remembered per file. Source
in `SheetNavigator/`, solution-level docs at the repo root.

## Working agreement

All decisions are the human's. Ask — never assume, infer, or pick an approach on
your own.

- Every open choice is a question first: approach, naming, structure, behavior,
  file layout, scope, edge cases.
- Ambiguity in a feature description means ask, never interpret.
- Do not merge. Report back and stop.

## Dev environment

- Visual Studio 2022 or later with the **Office/SharePoint development**
  workload, targeting .NET Framework 4.7.2.
- ClickOnce signing needs a certificate that is not in the repo: Project
  Properties, Signing, Create Test Certificate before the first build.
- Any full build registers the add-in with Excel from `bin\Release`; Build,
  Clean Solution unregisters it.

## Build and test commands

From `SheetNavigator\` (the project folder), one command per call, never
chained:

      "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" SheetNavigator.csproj -t:Compile -p:Configuration=Release

That compiles without registering the add-in. There is no test project;
behavior is verified by the human in Excel with Ctrl+F5 (Release, start without
debugging). Never use the Debug configuration or F5.

Never run a build as a baseline check before you have changed anything,
including when setting up a worktree. A skill step that says to is overridden
by this rule.

## Code style

- .NET naming: PascalCase for types, methods and constants (no ALL_CAPS),
  camelCase for parameters, locals and private fields. Descriptive names, never
  Excel's VBA-style `Wb`/`Sh` abbreviations.
- Constants first in a class, nested classes last.
- Every method gets a one- or two-line `<summary>`. Comments explain why, not
  what. No fluff.
- Source files keep their UTF-8 BOM and CRLF endings (see .editorconfig).
- Anything that changes Excel state temporarily (ScreenUpdating, events) is
  restored in a `finally`. COM calls that can fail while Excel is busy are
  wrapped in try/catch.

## Testing

- No test project yet. Verify with the compile command above, then tell the
  human exactly what to check in Excel after Ctrl+F5.

## Docs

Update project meta as appropriate: STATUS.md, CHANGELOG.md, TODO.md, README.md,
release/README.txt. CHANGELOG and TODO entries are single-line when possible.

New entries in STATUS.md, CHANGELOG.md and TODO.md go at the top of their
section, never mid-list or at the bottom.

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
