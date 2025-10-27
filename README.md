# PowerPoint Shortcuts Add-in

A productivity-focused PowerPoint add-in that provides custom keyboard shortcuts and professional tools to enhance your presentation workflow.

## Features

### Keyboard Shortcuts
- **Text & Objects**: Paste unformatted, convert text to shapes, split/join textboxes
- **Alignment**: Center, left, right, middle alignment with custom shortcuts
- **Distribution**: Horizontal and vertical object distribution
- **Professional Elements**: Insert footnotes, legends, and stickers
- **Formatting**: Cycle accent colors, apply professional styling

### Task Pane Interface
- Visual access to all shortcuts and functions
- Professional color palette
- Element insertion tools
- Help and documentation

## Installation

### Method 1: GitHub Pages (Recommended)
1. Enable GitHub Pages for this repository
2. Download the `manifest.xml` file
3. Open PowerPoint
4. Go to **Insert** → **My Add-ins** → **Upload My Add-in**
5. Select the `manifest.xml` file

### Method 2: Local Development
1. Clone this repository
2. Run a local web server (e.g., `python -m http.server 8000`)
3. Update manifest.xml URLs to point to localhost
4. Upload manifest.xml to PowerPoint

## Available Shortcuts

| Shortcut | Function |
|----------|----------|
| `Ctrl+Alt+T` | Paste unformatted text |
| `Shift+Alt+Z` | Convert text to autoshape |
| `Alt+Ctrl+J` | Split/join textboxes |
| `Shift+Alt+E` | Make objects same width |
| `Shift+Alt+H` | Make objects same height |
| `Ctrl+Alt+C` | Align center |
| `Ctrl+Alt+L` | Align left |
| `Ctrl+Alt+R` | Align right |
| `Ctrl+Alt+M` | Align middle |
| `Alt+Shift+D` | Distribute horizontally |
| `Alt+Shift+V` | Distribute vertically |
| `Ctrl+Alt+F` | Insert footnote |
| `Ctrl+Alt+G` | Insert legend |
| `Ctrl+Alt+S` | Insert sticker |
| `Shift+Alt+A` | Cycle accent colors |
| `Ctrl+Alt+P` | Green print export |
| `Ctrl+Alt+Q` | Show quick keys |
| `Ctrl+Alt+Y` | Reset elements |
| `Ctrl+Alt+K` | Show task pane |

## Technical Details

- **Platform**: Office Add-ins (Web-based)
- **Compatibility**: PowerPoint 2016 and later
- **Framework**: Vanilla JavaScript with Office.js
- **Manifest Version**: 3.0

## Limitations

Due to Office Add-in platform restrictions, some shortcuts cannot be implemented:
- Single digit shortcuts (Ctrl+1, Ctrl+2, etc.)
- Shortcuts missing second modifier (Alt+Q)
- Some advanced PowerPoint API features

## Development

### File Structure
```
├── manifest.xml           # Add-in manifest
├── src/
│   ├── commands/          # Keyboard shortcut implementations
│   ├── taskpane/          # User interface
│   └── shared/            # Shared utilities and functions
└── assets/                # Icons and resources
```

### Building
No build process required - this is a vanilla JavaScript add-in.

### Testing
1. Start local web server
2. Update manifest.xml with local URLs
3. Sideload in PowerPoint for testing

## License

This project is open source and available under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

