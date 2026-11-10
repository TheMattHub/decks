# Claude Slides from Claude Code

## Where decks live

Standalone Claude Design (claude.ai/design) closes on 14 December 2026. New decks are Slides
artifacts in Claude. Migrated Claude Design projects become artifacts; their chats and comments do
not move. Find a deck with the Artifact tool's `list` action. A Slides deck reads as the shared Slides
runtime, with the deck's content in its own files.

## Getting a deck

- **A deck exists** (a link from the user, or found with `list`): `read` its `SKILL.md`. That file is
  the type's authoring guide for the current release; follow it for the file format, publishing and
  revising. Read its `artifact-type/reference/*.md` files when it points to them. They change between
  releases, so read them each time instead of relying on a copy.
- **No deck yet**: if the Artifact tool offers the Slides type (a type or intent for slides), create the
  deck with it. Otherwise ask the user to start an empty slide deck in Claude and paste its link, then
  fill it.

## The deck's files, in short

- `project/deck.json` holds the title, the slide order, the sections (one sentence each) and the
  Google Fonts faces.
- `project/slides/<id>.html` holds one `<section>` per slide on a 1920 x 1080 canvas, with inline
  styles only. Speaker notes go in a final `<aside>`. Use `data-transition="fade|push|magic"` on the
  section and `data-build-in="rise|fade|pop"` on pinned children.
- Images are uploaded first as assets (`publish` with `asset: true`). The reply's `/_blob/<id>` goes
  verbatim into `<img src>`.
- Publish two or three slides per call and send only the files that changed.

## What the PowerPoint export keeps and drops

The user runs the export from the deck (Share > Export).

- **Kept**: layout, text as native text boxes, tables as native tables, shapes, images, speaker notes,
  some embedded fonts.
- **Dropped**: videos, transitions and build-ins. `references/powerpoint.md` adds them back.
- **Transitions and build-ins from the deck itself**: fetch the deck's files (`project/deck.json` and
  `project/slides/*.html`) into one folder, each at its deck path (the Artifact tool's `read` with `paths`),
  then `uv run "<skill>/scripts/slides_manifest.py" <folder> manifest.json`. `Invoke-FinishManifest` applies
  the manifest to the export. A structure the script does not recognise stops it with one line (then write
  the finishing calls by hand, as before); unknown effect names are only warned about.
- **Generic names**: shapes are called "Text 3", "Shape 2", "Table 0", and text uses theme colours that
  COM cannot read back. Finishing scripts therefore find shapes by their text and write colours as
  fixed hex values.
- **One layout**: every slide sits on one generic layout, so a template applied afterwards does not
  restyle the slides.

## Gotchas

- When publishing, give long Windows paths, not 8.3 short names such as `JOHNDO~1`. A permission rule
  on the short form blocked a publish.
- Never attach an old Claude Design HTML export (a self-unpacking bundle, base64 and gzip inside JSON,
  several MB) to a prompt: it overflowed the context twice. Read the migrated artifact's files instead,
  or decode only the template block.
- Text wraps only at spaces. Size boxes to the longest word, keep an amount and its unit on one line,
  and use no font smaller than 24 px.
- Claude Code cannot see the deck rendered unless it exports PNGs. When the framing of an image or a
  crop matters, ask the user to look, or check it after the PowerPoint export.
