# Team Tools Site

Static HTML tools hosted via GitHub Pages.

## Adding a new tool

1. Drop your `.html` file into this folder (e.g. `my-new-tool.html`).
2. Paste this line into the `<head>` section of that file, if it's not already there:
   ```html
   <meta name="robots" content="noindex, nofollow">
   ```
3. Add a link to it in `index.html` — copy one of the existing `<li>` blocks and update the href, title, and description.
4. Commit and push. It'll be live in a minute or two.

## Files

- `index.html` — landing page listing all tools
- `robots.txt` — tells search engines / AI crawlers not to index this site
- `transcript-reformatter.html` — (placeholder — replace with your actual tool file)