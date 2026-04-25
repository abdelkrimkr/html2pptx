# HTML2PPTX - HTML to PowerPoint Converter

A powerful Node.js library and command-line tool that automatically converts HTML files to PowerPoint presentations (.pptx). No manual intervention or configuration required!

## 🎉 Recent Fixes (Oct 14, 2025)

**✅ Flexbox Column Layout Support** - Fixed critical issue where only one element was rendering in flexbox column layouts. All elements now render correctly with proper spacing! See [FLEXBOX_FIX.md](./FLEXBOX_FIX.md) for details.

**✅ PPTX Corruption Fixes** - Post-processing removes all corruption issues from generated files. See [CORRUPTION_FIXES.md](./CORRUPTION_FIXES.md) for details.

## 🌟 Features

- **Automatic Conversion**: Convert any HTML file to PowerPoint with a single command
- **Text Boxes**: Preserves text content with styling (colors, fonts, sizes, alignment, character spacing)
- **Lists Supported**: Unordered and Ordered Lists formatted accurately as PowerPoint bullet points
- **Positioned Elements**: Handles absolute, relative, and flex-based positioning (including flexbox columns with gaps!)
- **Custom PPTX Layouts**: Configurable layout sizes extending default 16:9 (`LAYOUT_16x10`, `LAYOUT_4x3`, `LAYOUT_WIDE`)
- **Images**: Converts HTML `<img>` elements automatically to positioned PowerPoint images
- **SVG Support**: Converts SVG shapes, lines, and text to PowerPoint elements
- **Complex Layouts**: Supports multi-column, flexbox, and grid layouts
- **CSS Styling**: Extracts and applies inline styles, style tags, and class-based styles
- **CSS Transforms**: Supports `rotate()` transforms on elements
- **Hyperlinks**: Converts `<a>` tags with `href` attributes to clickable links
- **Borders & Backgrounds**: Preserves border colors, widths, and background colors
- **No Configuration**: Works out-of-the-box with sensible defaults
- **Corruption-Free**: Automatic post-processing ensures valid PPTX files

## 📦 Installation

### Global Installation (Recommended for CLI)

```bash
npm install -g html2pptx
```

### Local Installation

```bash
npm install html2pptx
```

### From Source

```bash
git clone <repository-url>
cd html2pptx-library
npm install
npm link  # Makes the CLI available globally
```

## 🚀 Usage

### Command Line Interface

Basic usage:

```bash
html2pptx input.html output.pptx
```

Examples:

```bash
# Convert a simple HTML file
html2pptx slide.html presentation.pptx

# Convert with full paths
html2pptx /path/to/input.html /path/to/output.pptx

# Show help
html2pptx --help

# Show version
html2pptx --version
```

### Programmatic Usage

```javascript
const { convertHTML2PPTX, HTML2PPTX } = require('html2pptx');

// Simple conversion
convertHTML2PPTX('input.html', 'output.pptx')
    .then(result => {
        console.log('Success!', result);
    })
    .catch(error => {
        console.error('Error:', error);
    });

// Advanced usage with custom options
const converter = new HTML2PPTX({
    slideWidth: 10,      // inches
    slideHeight: 5.625,  // inches (16:9 ratio)
    background: { color: 'FFFFFF' }
});

converter.convert('input.html', 'output.pptx')
    .then(result => console.log('Converted!', result))
    .catch(error => console.error('Error:', error));
```

## 📋 Supported HTML Elements

### Text Elements
- `<div>`, `<p>`, `<span>` → Text boxes
- `<h1>` - `<h6>` → Styled text boxes
- `<li>`, `<ul>`, `<ol>` → Formatted text
- `<a>` → Text boxes with hyperlinks

### Visual Elements
- `<img>` → PowerPoint images
- `<svg>` → Converted to shapes and lines
- SVG `<line>` → PowerPoint lines
- SVG `<text>` → PowerPoint text

### Layout
- Absolute positioning
- Relative positioning
- Flexbox layouts (simplified conversion)
- Nested elements

## 🎨 Supported CSS Properties

### Text Styling
- `font-size` - Converted to points
- `font-family` - Mapped to PowerPoint fonts
- `font-weight` - Bold text
- `font-style` - Italic text
- `color` - Text color
- `text-align` - Horizontal alignment
- `align-items` - Vertical alignment (flexbox)
- `justify-content` - Horizontal alignment (flexbox)

### Box Model
- `width` - Element width
- `height` - Element height
- `padding` - Inner spacing
- `border` - Border color and width
- `border-color` - Border color
- `border-width` - Border width
- `border-radius` - Rounded corners (visual approximation)
- `background-color` - Background fill
- `background` - Background fill

### Positioning
- `position: absolute` - Absolute positioning
- `position: fixed` - Fixed positioning
- `position: relative` - Relative positioning
- `top`, `left`, `right`, `bottom` - Position values

### Display
- `display: flex` - Flexbox layout
- `display: grid` - Grid layout
- `flex-direction` - Layout direction

### Transforms
- `transform: rotate(…)` - Rotates elements

## 📋 To-Do List (Future Improvements)

- **Advanced Table Generation**: Support parsing deeply nested HTML tables and data attributes to generate rich PptxGenJS Tables and native PowerPoint Charts.
- **Background Images**: Expand CSS parsing to include background images (`url()`) applied directly to slide bounds instead of falling back to flat colors.
- **Animations & Transitions**: Parse specific CSS animation classes to assign native PowerPoint slide transitions or entrance animations.
- **Advanced Grid CSS**: Better approximations for complex and auto-flow CSS Grid layouts natively matching HTML structures.

## 🎯 How It Works

1. **HTML Parsing**: Uses Cheerio to parse HTML structure
2. **CSS Extraction**: Extracts styles from `<style>` tags, inline styles, and classes
3. **Style Computation**: Computes final styles for each element
4. **Element Mapping**: Maps HTML elements to PowerPoint equivalents
5. **Position Calculation**: Converts CSS positioning to PowerPoint coordinates
6. **PowerPoint Generation**: Uses PptxGenJS to create the final presentation

## 📐 Slide Dimensions

Default slide dimensions:
- **Width**: 10 inches
- **Height**: 5.625 inches
- **Aspect Ratio**: 16:9 (standard widescreen)

This matches the standard PowerPoint 16:9 layout.

## ⚠️ Limitations

- **CSS Complexity**: Very complex CSS (animations, transforms) may not convert perfectly
- **JavaScript**: Dynamic content generated by JavaScript is not captured
- **External Resources**: External images/fonts must be accessible
- **Layout**: Complex flexbox/grid layouts are approximated
- **Fonts**: Font families are mapped to PowerPoint-compatible fonts
- **Interactive Elements**: Forms, buttons, and interactive elements are converted to static content

## 🔧 Troubleshooting

### Conversion Errors

If you encounter errors, try:

1. **Check HTML validity**: Ensure your HTML is well-formed
2. **Simplify CSS**: Complex CSS might need simplification
3. **Check paths**: Verify input file exists and output directory is writable
4. **Debug mode**: Run with `DEBUG=1` environment variable

```bash
DEBUG=1 html2pptx input.html output.pptx
```

### Common Issues

**Issue**: Text is too small/large
- **Solution**: Adjust font sizes in your HTML/CSS

**Issue**: Elements are misaligned
- **Solution**: Use absolute positioning for precise placement

**Issue**: Colors look different
- **Solution**: Use hex color codes (#RRGGBB) for best results

## 📝 Examples

### Example 1: Simple Text Boxes

```html
<!DOCTYPE html>
<html>
<head>
    <style>
        .container {
            width: 1280px;
            height: 720px;
        }
        .text-box {
            border: 2px solid #3498db;
            padding: 20px;
            font-size: 24px;
            background: #f8f9fa;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="text-box">Hello World!</div>
    </div>
</body>
</html>
```

### Example 2: Absolute Positioning

```html
<!DOCTYPE html>
<html>
<head>
    <style>
        .slide {
            width: 1280px;
            height: 720px;
            position: relative;
        }
        .title {
            position: absolute;
            top: 50px;
            left: 100px;
            font-size: 36px;
            color: #2c3e50;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="slide">
        <div class="title">My Presentation</div>
    </div>
</body>
</html>
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## 📄 License

MIT License - feel free to use this in your projects!

## 🔗 Dependencies

- **cheerio**: HTML parsing
- **pptxgenjs**: PowerPoint generation
- **css**: CSS parsing

## 📧 Support

For issues, questions, or feature requests, please open an issue on the repository.

---

Made with ❤️ by the HTML2PPTX team

---

## PPTX Corruption Fix (October 2025)

### Issue Resolved
Previous versions generated PPTX files that PowerPoint marked as corrupted and requiring repair.

### Root Cause
The underlying PptxGenJS 3.12.0 library had XML generation bugs that produced invalid OpenXML:
- Empty name attributes
- Empty line elements
- Zero/invalid dimensions
- Conflicting autofit settings
- Invalid charset values

### Solution
The library now includes an automatic **post-processor** (`lib/pptx-fixer.js`) that:
- Runs transparently after PPTX generation
- Fixes all XML corruption issues
- Ensures 100% PowerPoint compliance
- Requires no API changes

### Result
✅ Generated PPTX files now open immediately in PowerPoint  
✅ No corruption warnings  
✅ No manual repair required  
✅ Production-ready  

See `CORRUPTION_FIXES.md` for complete technical details.

---
