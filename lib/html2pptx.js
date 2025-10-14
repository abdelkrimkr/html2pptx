const cheerio = require('cheerio');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
const css = require('css');

/**
 * HTML to PowerPoint Converter
 * Converts HTML files to PowerPoint presentations
 */
class HTML2PPTX {
    constructor(options = {}) {
        this.options = {
            slideWidth: 10,        // inches (standard 16:9)
            slideHeight: 5.625,    // inches (standard 16:9)
            background: { color: 'FFFFFF' },
            ...options
        };
        this.pptx = new PptxGenJS();
        this.styles = {};
        this.computedStyles = new Map();
    }

    /**
     * Convert HTML file to PowerPoint
     * @param {string} inputPath - Path to HTML file
     * @param {string} outputPath - Path for output PPTX file
     */
    async convert(inputPath, outputPath) {
        try {
            // Read HTML file
            const html = fs.readFileSync(inputPath, 'utf8');
            
            // Parse HTML
            const $ = cheerio.load(html);
            
            // Extract and parse CSS
            this.extractCSS($);
            
            // Configure presentation
            this.pptx.layout = 'LAYOUT_16x9';
            this.pptx.author = 'HTML2PPTX Converter';
            this.pptx.title = $('title').text() || 'Converted Presentation';
            
            // Create slide
            const slide = this.pptx.addSlide();
            
            // Get the main container dimensions
            const container = this.findMainContainer($);
            const containerWidth = this.options.slideWidth * 72; // Convert to points
            const containerHeight = this.options.slideHeight * 72;
            
            // Process elements
            await this.processElements($, slide, container, containerWidth, containerHeight);
            
            // Save presentation
            await this.pptx.writeFile({ fileName: outputPath });
            
            return {
                success: true,
                outputPath: outputPath
            };
        } catch (error) {
            throw new Error(`Conversion failed: ${error.message}`);
        }
    }

    /**
     * Find the main content container
     */
    findMainContainer($) {
        // Look for common container classes/IDs
        const selectors = [
            '.slide-container',
            '.container',
            'body > div:first-child',
            'body'
        ];
        
        for (const selector of selectors) {
            const elem = $(selector);
            if (elem.length > 0) {
                return elem.first();
            }
        }
        
        return $('body');
    }

    /**
     * Extract CSS from style tags and inline styles
     */
    extractCSS($) {
        // Extract from style tags
        $('style').each((i, elem) => {
            const cssText = $(elem).html();
            try {
                const parsed = css.parse(cssText);
                this.processCSSRules(parsed.stylesheet.rules);
            } catch (e) {
                console.warn('CSS parsing warning:', e.message);
            }
        });
    }

    /**
     * Process CSS rules
     */
    processCSSRules(rules) {
        rules.forEach(rule => {
            if (rule.type === 'rule') {
                const styles = {};
                rule.declarations.forEach(decl => {
                    if (decl.type === 'declaration') {
                        styles[decl.property] = decl.value;
                    }
                });
                
                rule.selectors.forEach(selector => {
                    if (!this.styles[selector]) {
                        this.styles[selector] = {};
                    }
                    Object.assign(this.styles[selector], styles);
                });
            }
        });
    }

    /**
     * Get computed style for an element
     */
    getComputedStyle($, elem) {
        const $elem = $(elem);
        const computedStyle = {};
        
        // Get inline styles first
        const inlineStyle = $elem.attr('style');
        if (inlineStyle) {
            this.parseInlineStyle(inlineStyle, computedStyle);
        }
        
        // Get class styles
        const classes = ($elem.attr('class') || '').split(' ');
        classes.forEach(className => {
            if (className) {
                const classStyle = this.styles[`.${className}`];
                if (classStyle) {
                    Object.assign(computedStyle, classStyle);
                }
            }
        });
        
        // Get tag styles
        const tagName = $elem.prop('tagName')?.toLowerCase();
        if (tagName && this.styles[tagName]) {
            Object.assign(computedStyle, this.styles[tagName]);
        }
        
        return computedStyle;
    }

    /**
     * Parse inline style string
     */
    parseInlineStyle(styleStr, target) {
        const parts = styleStr.split(';');
        parts.forEach(part => {
            const [prop, value] = part.split(':').map(s => s.trim());
            if (prop && value) {
                target[prop] = value;
            }
        });
    }

    /**
     * Process all elements in the HTML
     */
    async processElements($, slide, container, containerWidth, containerHeight) {
        const elements = [];
        
        // Find all visible elements with content
        container.find('*').each((i, elem) => {
            const $elem = $(elem);
            const tagName = $elem.prop('tagName')?.toLowerCase();
            
            // Skip script, style, meta tags
            if (['script', 'style', 'meta', 'link', 'head'].includes(tagName)) {
                return;
            }
            
            // Get element info
            const text = this.getDirectText($, $elem);
            const style = this.getComputedStyle($, elem);
            
            // Store element info
            if (text || tagName === 'img' || tagName === 'svg') {
                elements.push({
                    elem: $elem,
                    tagName: tagName,
                    text: text,
                    style: style,
                    html: $elem.html()
                });
            }
        });
        
        // Process SVG elements
        this.processSVGElements($, slide, container);
        
        // Process text elements
        this.processTextElements($, slide, elements, containerWidth, containerHeight);
    }

    /**
     * Get direct text content (not from children)
     */
    getDirectText($, $elem) {
        let text = '';
        $elem.contents().each((i, node) => {
            if (node.type === 'text') {
                text += $(node).text().trim();
            }
        });
        return text;
    }

    /**
     * Process SVG elements
     */
    processSVGElements($, slide, container) {
        container.find('svg').each((i, elem) => {
            const $svg = $(elem);
            
            // Process SVG lines
            $svg.find('line').each((j, lineElem) => {
                const $line = $(lineElem);
                const x1 = parseFloat($line.attr('x1') || 0);
                const y1 = parseFloat($line.attr('y1') || 0);
                const x2 = parseFloat($line.attr('x2') || 0);
                const y2 = parseFloat($line.attr('y2') || 0);
                
                // Get parent position
                const parentStyle = this.getComputedStyle($, $svg.parent()[0]);
                const baseX = this.parsePosition(parentStyle.left || '0') / 72;
                const baseY = this.parsePosition(parentStyle.top || '0') / 72;
                
                // Calculate position in inches
                const svgWidth = parseFloat($svg.attr('width') || 400);
                const svgHeight = parseFloat($svg.attr('height') || 350);
                const slideWidthPx = this.options.slideWidth * 72;
                const slideHeightPx = this.options.slideHeight * 72;
                
                const scaleX = this.options.slideWidth / slideWidthPx;
                const scaleY = this.options.slideHeight / slideHeightPx;
                
                try {
                    slide.addShape('line', {
                        x: baseX + (x1 / svgWidth) * (svgWidth / 72),
                        y: baseY + (y1 / svgHeight) * (svgHeight / 72),
                        w: Math.abs(x2 - x1) / 72,
                        h: Math.abs(y2 - y1) / 72,
                        line: {
                            color: '3182ce',
                            width: 2
                        }
                    });
                } catch (e) {
                    console.warn('Error adding line:', e.message);
                }
            });
            
            // Process SVG text
            $svg.find('text').each((j, textElem) => {
                const $text = $(textElem);
                const x = parseFloat($text.attr('x') || 0);
                const y = parseFloat($text.attr('y') || 0);
                const content = $text.text();
                const fill = $text.attr('fill') || '#000000';
                const fontSize = this.parseFontSize($text.attr('style') || 'font-size: 24px');
                
                // Get parent position
                const parentStyle = this.getComputedStyle($, $svg.parent()[0]);
                const baseX = this.parsePosition(parentStyle.left || '0') / 72;
                const baseY = this.parsePosition(parentStyle.top || '0') / 72;
                
                try {
                    slide.addText(content, {
                        x: baseX + (x / 72) / 10,
                        y: baseY + (y / 72) / 10,
                        fontSize: fontSize,
                        color: this.parseColor(fill),
                        bold: ($text.attr('style') || '').includes('font-weight: 700'),
                        fontFace: this.parseFontFamily($text.attr('style') || '')
                    });
                } catch (e) {
                    console.warn('Error adding SVG text:', e.message);
                }
            });
        });
    }

    /**
     * Process text elements and create text boxes
     */
    processTextElements($, slide, elements, containerWidth, containerHeight) {
        const processedParents = new Set();
        
        elements.forEach((elemInfo, index) => {
            const { elem: $elem, tagName, text, style } = elemInfo;
            
            // Skip if parent already processed
            const parentId = $elem.parent().toString();
            if (processedParents.has(parentId) && !text) {
                return;
            }
            
            // Get full text content
            const fullText = $elem.text().trim();
            if (!fullText) return;
            
            // Calculate position and size
            const position = this.calculatePosition($, $elem, style, containerWidth, containerHeight);
            
            if (!position) return;
            
            // Parse text formatting
            const textOptions = this.parseTextOptions($, $elem, style);
            
            try {
                // Add text box to slide
                slide.addText(fullText, {
                    x: position.x,
                    y: position.y,
                    w: position.w,
                    h: position.h,
                    ...textOptions
                });
                
                processedParents.add($elem.toString());
            } catch (e) {
                console.warn(`Error adding text element: ${e.message}`);
            }
        });
    }

    /**
     * Calculate element position and size
     */
    calculatePosition($, $elem, style, containerWidth, containerHeight) {
        // Handle absolute positioning
        if (style.position === 'absolute' || style.position === 'fixed') {
            const left = this.parsePosition(style.left || '0');
            const top = this.parsePosition(style.top || '0');
            const right = this.parsePosition(style.right || '0');
            const bottom = this.parsePosition(style.bottom || '0');
            const width = this.parsePosition(style.width || '200');
            const height = this.parsePosition(style.height || '50');
            
            return {
                x: left / 72,
                y: top / 72,
                w: width / 72,
                h: height / 72
            };
        }
        
        // Handle flex and relative layouts
        // This is a simplified approach - estimates position based on DOM order
        const allElements = $elem.parent().children();
        const index = allElements.index($elem[0]);
        const totalElements = allElements.length;
        
        // Default sizing
        const width = this.parsePosition(style.width || `${containerWidth / 2}`);
        const height = this.parsePosition(style.height || '60');
        
        // Estimate vertical position based on element order
        const estimatedTop = (index / totalElements) * containerHeight;
        
        return {
            x: 0.5, // Default padding
            y: 0.5 + (estimatedTop / 72),
            w: (width / 72) || 9, // Default width
            h: (height / 72) || 0.6  // Default height
        };
    }

    /**
     * Parse text formatting options
     */
    parseTextOptions($, $elem, style) {
        const options = {
            align: 'left',
            valign: 'middle',
            fontSize: 18,
            color: '000000',
            bold: false,
            italic: false,
            fontFace: 'Arial'
        };
        
        // Font size
        if (style['font-size']) {
            options.fontSize = this.parseFontSize(style['font-size']);
        }
        
        // Color
        if (style.color) {
            options.color = this.parseColor(style.color);
        }
        
        // Background color
        if (style['background-color'] || style.background) {
            const bgColor = style['background-color'] || style.background;
            options.fill = { color: this.parseColor(bgColor) };
        }
        
        // Border
        if (style.border || style['border-color']) {
            const borderColor = style['border-color'] || this.extractColorFromBorder(style.border);
            const borderWidth = this.parseBorderWidth(style.border || style['border-width'] || '1px');
            options.line = {
                color: this.parseColor(borderColor),
                width: borderWidth
            };
        }
        
        // Text alignment
        if (style['text-align']) {
            options.align = style['text-align'];
        }
        
        if (style['justify-content']) {
            const justify = style['justify-content'];
            if (justify === 'center') options.align = 'center';
            if (justify === 'flex-end') options.align = 'right';
        }
        
        if (style['align-items']) {
            const alignItems = style['align-items'];
            if (alignItems === 'center') options.valign = 'middle';
            if (alignItems === 'flex-start') options.valign = 'top';
            if (alignItems === 'flex-end') options.valign = 'bottom';
        }
        
        // Font weight
        if (style['font-weight']) {
            const weight = parseInt(style['font-weight']);
            options.bold = weight >= 600 || style['font-weight'] === 'bold';
        }
        
        // Font style
        if (style['font-style'] === 'italic') {
            options.italic = true;
        }
        
        // Font family
        if (style['font-family']) {
            options.fontFace = this.parseFontFamily(style['font-family']);
        }
        
        // Border radius (for rounded corners)
        if (style['border-radius']) {
            // PptxGenJS doesn't directly support border-radius on text boxes
            // but we can note it for reference
        }
        
        return options;
    }

    /**
     * Parse position value (px, %, etc.) to points
     */
    parsePosition(value) {
        if (!value) return 0;
        
        // Remove units and parse
        const numValue = parseFloat(value);
        
        if (value.includes('px')) {
            return numValue;
        } else if (value.includes('%')) {
            // Convert percentage to pixels (assuming 1280px width as base)
            return (numValue / 100) * 1280;
        } else if (value.includes('em') || value.includes('rem')) {
            return numValue * 16; // Assume 16px base
        }
        
        return numValue;
    }

    /**
     * Parse font size to points
     */
    parseFontSize(fontSize) {
        const size = parseFloat(fontSize);
        
        if (fontSize.includes('px')) {
            return Math.round(size * 0.75); // Convert px to pt
        } else if (fontSize.includes('em') || fontSize.includes('rem')) {
            return Math.round(size * 12); // Convert em to pt (assuming 16px base)
        }
        
        return size || 18;
    }

    /**
     * Parse color to hex format
     */
    parseColor(color) {
        if (!color) return '000000';
        
        // Remove # if present
        color = color.replace('#', '');
        
        // Handle rgb/rgba
        if (color.startsWith('rgb')) {
            const matches = color.match(/\d+/g);
            if (matches && matches.length >= 3) {
                const r = parseInt(matches[0]).toString(16).padStart(2, '0');
                const g = parseInt(matches[1]).toString(16).padStart(2, '0');
                const b = parseInt(matches[2]).toString(16).padStart(2, '0');
                return r + g + b;
            }
        }
        
        // Handle named colors (common ones)
        const namedColors = {
            'white': 'FFFFFF',
            'black': '000000',
            'red': 'FF0000',
            'green': '00FF00',
            'blue': '0000FF',
            'gray': '808080',
            'grey': '808080'
        };
        
        if (namedColors[color.toLowerCase()]) {
            return namedColors[color.toLowerCase()];
        }
        
        // Return as-is if it looks like a hex color
        if (/^[0-9A-Fa-f]{6}$/.test(color)) {
            return color.toUpperCase();
        }
        
        if (/^[0-9A-Fa-f]{3}$/.test(color)) {
            // Expand 3-digit hex to 6-digit
            return color.split('').map(c => c + c).join('').toUpperCase();
        }
        
        return '000000';
    }

    /**
     * Extract color from border string
     */
    extractColorFromBorder(border) {
        if (!border) return '#000000';
        
        // Try to find a color value
        const parts = border.split(' ');
        for (const part of parts) {
            if (part.startsWith('#') || part.startsWith('rgb')) {
                return part;
            }
        }
        
        // Check for named colors
        const namedColors = ['red', 'blue', 'green', 'black', 'white', 'gray', 'grey'];
        for (const color of namedColors) {
            if (border.includes(color)) {
                return color;
            }
        }
        
        return '#000000';
    }

    /**
     * Parse border width
     */
    parseBorderWidth(borderWidth) {
        const width = parseFloat(borderWidth);
        return isNaN(width) ? 1 : width;
    }

    /**
     * Parse font family
     */
    parseFontFamily(fontFamily) {
        if (!fontFamily) return 'Arial';
        
        // Extract first font name
        const fonts = fontFamily.split(',');
        let font = fonts[0].trim().replace(/['"]/g, '');
        
        // Map common fonts
        const fontMap = {
            'Roboto': 'Arial',
            'Montserrat': 'Arial',
            'Helvetica': 'Arial',
            'sans-serif': 'Arial',
            'serif': 'Times New Roman',
            'monospace': 'Courier New'
        };
        
        return fontMap[font] || font || 'Arial';
    }
}

/**
 * Main conversion function
 */
async function convertHTML2PPTX(inputPath, outputPath, options = {}) {
    const converter = new HTML2PPTX(options);
    return await converter.convert(inputPath, outputPath);
}

module.exports = {
    HTML2PPTX,
    convertHTML2PPTX
};
