# SM Line USA Tools

Internal tools served at https://us-sm.github.io/

## Add a tool
1. Create a folder (e.g. `my-tool/`) with an `index.html`.
2. Add an entry to `tools.json`:
   ```json
   { "path": "my-tool/", "name": { "en": "...", "ko": "..." }, "desc": { "en": "...", "ko": "..." }, "tag": "..." }
   ```
3. Optional: put `<a href="../">← All tools</a>` on the tool page and use the
   `usSmLang` localStorage key so the EN/KO choice is shared.
