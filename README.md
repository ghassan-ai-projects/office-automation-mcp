# Office Automation MCP Server

An [MCP](https://modelcontextprotocol.io) server that wraps **python-pptx** into a type-safe, persistent API for generating and editing PowerPoint presentations. Designed for AI agents to create professional branded decks programmatically.

## Features

- **10 MCP tools** — full presentation lifecycle: create, add slides, brand, duplicate, reorder, delete, background, images, export
- **3 MCP resources** — read deck structure, single slide, brand config
- **8 slide types** — title, content/bullets, 2×2 bento cards, process flow, quotes, fun scale, CTA, thanks
- **Nature Theme** — sage white + forest green brand colors
- **16:9 widescreen** — professional aspect ratio
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
| `create_deck(path, template)` | Create empty branded presentation |
| `add_slide(path, slide_type, data, index)` | Add a typed slide |
| `apply_brand(path, template, slide_index)` | Apply brand colors |
| `read_slide_structure(path)` | Read deck structure as JSON |
| `duplicate_slide(path, source_index, target_index)` | Duplicate a slide |
| `delete_slide(path, index)` | Remove a slide |
| `reorder_slides(path, order)` | Reorder slides |
| `set_slide_background(path, slide_index, color)` | Set slide background |
| `add_image(path, slide_index, image_path, left, top, width, height)` | Add image to slide |
| `export_as_pdf(path, output_path)` | Export as PDF |

## Resources

| URI | Description |
|-----|-------------|
| `slides://{path}` | Full slide structure JSON |
| `slides://{path}/{index}` | Single slide JSON |
| `brand://current` | Current brand config |

## Project Structure

```
office-automation-mcp/
├── mcp_server.py        # Main MCP server (FastMCP)
├── brand.py             # Brand color/font definitions
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
