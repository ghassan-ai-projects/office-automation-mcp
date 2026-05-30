# Office Automation MCP Server

An [MCP](https://modelcontextprotocol.io) server that wraps **python-pptx** into a type-safe, persistent API for generating and editing PowerPoint presentations. Designed for AI agents to create professional branded decks programmatically.

## Features

- **11 MCP tools** — full presentation lifecycle: create, add slides, brand, duplicate, reorder, delete, background, images, export, text editing
- **3 MCP resources** — read deck structure, single slide, brand config
- **8 slide types** — title, content/bullets, 2×2 bento cards, process flow, quotes, fun scale, CTA, thanks
- **Configurable theme system** — switch between themes via JSON config or `PPTX_THEME` env var
- **Built-in themes** — Nature (sage green), Light Professional (blue), Dark (navy/rose)
- **16:9 widescreen** — professional aspect ratio (configurable)
- **macOS integration** — auto-opens in Microsoft PowerPoint after save

## Quick Start

```bash
# Clone
git clone git@github.com:ghassan-ai-projects/office-automation-mcp.git
cd office-automation-mcp

# Run via uv (no install needed)
uv run --with python-pptx --with fastmcp mcp_server.py
```

## Registration (OpenClaw)

Add to OpenClaw MCP config:

```bash
openclaw mcp set office-automation-engine '{
  "command": "uv",
  "args": [
    "run",
    "--with", "python-pptx",
    "--with", "fastmcp",
    "<path-to-repo>/mcp_server.py"
  ]
}'
```

Or add to `openclaw.json`:

```json
{
  "mcp": {
    "servers": {
      "office-automation-engine": {
        "command": "uv",
        "args": [
          "run",
          "--with", "python-pptx",
          "--with", "fastmcp",
          "<path-to-repo>/mcp_server.py"
        ]
      }
    }
  }
}
```

Replace `<path-to-repo>` with your checkout path (e.g. `~/projects/office-automation-mcp` or the absolute path on your machine).

## Theme Configuration

All brand colors, fonts, and defaults are defined in `themes.json` at the project root.

### Available Themes

| Key | Name | Palette |
|-----|------|---------|
| `nature` | Nature | Sage white + forest green (default) |
| `light-professional` | Light Professional | Clean white + blue accent |
| `dark` | Dark | Navy background + rose accent |

### Choosing a Default Theme

Set the `PPTX_THEME` environment variable to any theme key:

```bash
# Always use the dark theme
export PPTX_THEME=dark
uv run --with python-pptx --with fastmcp mcp_server.py
```

```bash
# Or inline for a single run
PPTX_THEME=dark uv run --with python-pptx --with fastmcp mcp_server.py
```

If the env var is unset or references a non-existent theme, the server falls back to `nature`.

### Adding a Custom Theme

Add a new entry under `"themes"` in `themes.json`:

```json
{
  "themes": {
    "ocean": {
      "name": "Ocean",
      "font": "Calibri",
      "colors": {
        "bg": "#E8F4F8",
        "card_bg": "#D1ECF1",
        "card2": "#BEE5EB",
        "border": "#86C4D4",
        "accent": "#0077B6",
        "accent2": "#005F8A",
        "success": "#2EC4B6",
        "text_primary": "#0A1E2A",
        "text_secondary": "#2E5A6B",
        "text_muted": "#5A8A9A",
        "white": "#FFFFFF"
      }
    }
  }
}
```

Your new theme is immediately available to all MCP tools (no restart needed if the server is still running — the JSON is re-read per request via caching).

### Configurable Defaults

The `"defaults"` section in `themes.json` controls slide-level settings:

```json
{
  "defaults": {
    "slide_width_inches": 13.333,
    "slide_height_inches": 7.5,
    "default_font": "Calibri",
    "footer_text": "Confidential",
    "auto_save": true
  }
}
```

| Key | Default | Description |
|-----|---------|-------------|
| `slide_width_inches` | `13.333` | Slide width (16:9) |
| `slide_height_inches` | `7.5` | Slide height (16:9) |
| `default_font` | `"Calibri"` | Fallback font for all slides |
| `footer_text` | `""` | Text appended to slide footers |
| `auto_save` | `true` | Automatically save on every tool call |

## Usage Examples

### Creating a Deck

Using the MCP tool from any OpenClaw agent:

```
Create a 5-slide pitch deck at ~/Documents/pitch.pptx with:
1. Title: "Project Alpha"
2. Content: Key milestones
3. Cards: 4 product features
4. Quote: Customer testimonial
5. CTA: "Get Started Today"
```

Or directly via command line:

```bash
uv run --with python-pptx --with fastmcp mcp_server.py \
  call-tool create_deck '{"path": "~/Documents/demo.pptx"}'
```

## Slide Types & Data Schemas

### `title`
```json
{ "headline": "", "subheadline": "", "tagline": "", "attribution": "", "url": "" }
```

### `content_bullets`
```json
{ "headline": "", "subtitle": "", "bullets": ["..."], "footer": "" }
```

### `cards`
```json
{ "headline": "", "subtitle": "", "cards": [{ "title": "", "description": "", "tags": [""] }] }
```
Supports 2 (1 row) or 4 (2×2 grid).

### `process_flow`
```json
{ "headline": "", "subtitle": "", "steps": [{ "number": 1, "label": "", "description": "" }] }
```
Up to 4 steps with connecting arrows.

### `quote`
```json
{ "headline": "", "subtitle": "", "quotes": [{ "text": "", "attribution": "" }] }
```

### `fun_scale`
```json
{ "headline": "", "subtitle": "", "levels": [{ "time": "", "label": "", "description": "" }] }
```
Late Night Commit Scale — colored severity gradient.

### `cta`
```json
{ "headline": "", "subheadline": "", "button_text": "", "url": "", "contact": "" }
```

### `thanks`
```json
{ "headline": "", "tagline": "", "url": "" }
```

## Tools API

| Tool | Description |
|------|-------------|
| `create_deck(path, template)` | Create empty branded presentation (template defaults to `PPTX_THEME` env var or `"nature"`) |
| `add_slide(path, slide_type, data, index)` | Add a typed slide |
| `apply_brand(path, template, slide_index)` | Apply brand colors (template defaults to `PPTX_THEME` env var or `"nature"`) |
| `read_slide_structure(path)` | Read deck structure as JSON |
| `duplicate_slide(path, source_index, target_index)` | Duplicate a slide |
| `delete_slide(path, index)` | Remove a slide |
| `reorder_slides(path, order)` | Reorder slides |
| `set_slide_background(path, slide_index, color)` | Set slide background |
| `add_image(path, slide_index, image_path, left, top, width, height)` | Add image to slide |
| `export_as_pdf(path, output_path)` | Export as PDF |
| `update_slide_text(path, slide_index, old_text, new_text)` | Find and replace text on a slide |

## Resources

| URI | Description |
|-----|-------------|
| `slides://{path}` | Full slide structure JSON |
| `slides://{path}/{index}` | Single slide JSON |
| `brand://current` | Current brand config (includes `default_theme` and `templates_available`) |

## Project Structure

```
office-automation-mcp/
├── mcp_server.py        # Main MCP server (FastMCP)
├── brand.py             # Theme loader — reads themes.json
├── themes.json          # Theme & defaults configuration
├── utils.py             # Shared helper functions
├── slides/
│   ├── __init__.py
│   ├── title.py         # Title slide builder
│   ├── content.py       # Content/bullets slide builder
│   ├── cards.py         # 2×2 bento card grid builder
│   ├── process.py       # Process flow builder
│   ├── quotes.py        # Quote/testimonial builder
│   ├── funscale.py      # Fun scale builder
│   ├── cta.py           # CTA slide builder
│   └── thanks.py        # Thanks/closing builder
├── examples/
│   └── client-pitch.py  # Full 7-slide pitch deck example
├── pyproject.toml       # Python project config
├── README.md            # This file
└── LICENSE              # MIT
```

## License

MIT — see [LICENSE](LICENSE).
