# Deploying to Streamlit Community Cloud

## Dependency files live at the repository root

| File | Where Streamlit looks | Why it matters |
|---|---|---|
| `packages.txt` | **Repository root only** | Root-only detection. While this file sat in `wc-platform/` it was never read, so no apt package was installed. |
| `requirements.txt` | Root, or beside the entrypoint | The looser rule is why Python dependencies installed while the system libraries silently did not. |

Both are kept at the root so they cannot drift apart again.

## packages.txt must contain bare package names only

Streamlit passes each line of `packages.txt` to `apt-get` verbatim. It does
**not** strip comments. A commented file makes apt try to install packages
called `THIS`, `MUST`, `LIVE`, and so on, and a single apostrophe anywhere in
the file (`entrypoint's`) fails the whole step with
`xargs: unmatched single quote`.

One package name per line. No comments, no blank lines, no quotes, no
apostrophes. Explanations belong here, not in the file.

### What each package is for

- **`libcairo2`** — provides `libcairo.so.2`, which `cairocffi` (under
  `cairosvg`) loads via `dlopen` to rasterise the Mind Map. This is the only
  native dependency CairoSVG has: it imports `cairocffi` and nothing else.
  `libcairo-gobject2`, `libpango*` and `libgdk-pixbuf*` are **WeasyPrint**
  dependencies, not CairoSVG's, and are deliberately not installed.

- **`libreoffice-writer`** — supplies the `soffice` binary used to convert
  the report `.docx` to PDF, and to render pass 1 so the two-pass TOC
  resolver can read real page numbers. The `libreoffice-writer` subset rather
  than the full suite: same binary, far smaller, and the full suite often
  exceeds the Community Cloud build limit.

### Not currently installed: `fonts-dejavu`

The report specifies DejaVu Sans on every element that prints a rupee sign,
because other fonts substitute a missing-glyph box for `₹`. If rendered PDFs
show boxes instead of `₹`, add `fonts-dejavu` as a third bare line.

## Graceful degradation

Neither system dependency is required for the app to function:

- **No LibreOffice** → `PdfUnavailable`; the app ships the editable `.docx`
  alone, with placeholder dots on the contents page, and says so.
- **No libcairo2** → `MindmapUnavailable`; the Mind Map section is skipped
  with a note pointing at the Transaction Snapshot, which carries the same
  recommendations as a table. The report still builds.

The `cairosvg` import is lazy, inside the function that rasterises. A
module-level import made a missing system library break reports for clients
who had no Mind Map at all.

## Python version

Deploy on **Python 3.13**. `matplotlib==3.11.1` publishes no `cp314` wheel, so
on Python 3.14 pip builds it from source — slow, and liable to fail on a
later rebuild. Every other pin has a prebuilt `cp313` wheel.

The Python version cannot be changed on a live app. To change it: note the
subdomain and secrets, delete the app, redeploy with the same subdomain,
choose the version under **Advanced settings**, and re-enter the secrets —
they are not preserved across deletion.

## Deploy steps

1. `share.streamlit.io` → **Create app** → deploy from GitHub; authorise
   access to the private repo.
2. Branch `main`, **main file path `wc-platform/app.py`** (not `app.py`).
3. **Advanced settings → Python 3.13.**
4. **Secrets:**
   ```toml
   APP_PASSWORD = "..."
   ANTHROPIC_API_KEY = "sk-ant-api03-..."
   ```
   Without `APP_PASSWORD` the app deploys but refuses all access, by design —
   it handles real client data on a public URL.
