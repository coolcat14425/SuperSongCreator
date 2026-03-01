diff --git a/README.md b/README.md
new file mode 100644
index 0000000000000000000000000000000000000000..3a95770b82ed43ef9f9e884d66ac576c3b1b419d
--- /dev/null
+++ b/README.md
@@ -0,0 +1,53 @@
+# SuperSongCreator (VB.NET)
+
+A production-oriented VB.NET WinForms app for building PowerPoint lyric decks from plain-text song files.
+
+## Song file format
+
+The parser expects this format:
+
+1. Line 1: Song title.
+2. Line 2: Credits (lyrics/authors/singer line).
+3. Line 3: CCLI song number.
+4. Remaining lines:
+   - `#` ends the current slide and starts a new one.
+   - `##...` marks copyright text (usually last line).
+   - `**...` marks slide notes/comments that will be written into PowerPoint notes for the current slide.
+
+Example:
+
+```text
+WAITING HERE FOR YOU
+Ben Glover | Ben McDonald | Dave Frey
+7036352
+
+If faith can move the mountains
+Let the mountains move
+**Band in by line 2
+We come with expectation
+Waiting here for You, waiting here for You
+#
+...
+##copyright text
+```
+
+## Features implemented
+
+- Load songs from a folder (`*.txt`).
+- Build song groups for services and save/recall groups.
+- Style settings (font name, size, font color, background color).
+- Generate `.pptx` deck and open it after generation.
+- Title slide per song using title style.
+- Error logging to `%AppData%\SuperSongCreator\logs\app-YYYYMMDD.log`.
+- Error alarm popup when failures occur.
+
+## Storage locations
+
+Under `%AppData%\SuperSongCreator`:
+- `settings.json`
+- `songgroups.json`
+- `logs\app-YYYYMMDD.log`
+
+## Build
+
+Open `SuperSongCreator.sln` in Visual Studio 2022+ and build for Windows.
