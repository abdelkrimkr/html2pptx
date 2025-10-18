const cheerio = require('cheerio');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
const css = require('css');
const { fixPPTX } = require('./pptx-fixer');

/**
 * HTML to PowerPoint Converter
 * Converts HTML files to PowerPoint presentations
 * 
 * ROOT CAUSE FIXES:
 * 1. Font Size: Uses proper scaling from HTML dimensions to slide dimensions
 * 2. Positioning: Simplified coordinate system with proper scaling
 * 3. Element Processing: Only processes leaf text elements, skips containers
 * 4. Dimension Scaling: Consistent scaling factor throughout
 * 5. CSS Parsing: Better handling of inline styles and common patterns
 * 6. Background: Support for gradients and solid colors
 * 7. PPTX Corruption: Post-processes generated files to fix PptxGenJS bugs
 *    - Empty name attributes
 *    - Empty line elements
 *    - Zero dimensions
 *    - Conflicting autofit settings
 *    - Invalid charset values
 */
class HTML2PPTX {
    constructor(options = {}) {
        this.options = {
            slideWidth: 10,        // inches (standard 16:9)
            slideHeight: 5.625,    // inches (standard 16:9)
            htmlWidth: 1280,       // Default HTML container width in pixels
            htmlHeight: 720,       // Default HTML container height in pixels
            background: { color: 'FFFFFF' },
            ...options
        };
        this.pptx = new PptxGenJS();
        this.styles = {};
        this.computedStyles = new Map();
        
        // Calculate scaling factors (HTML pixels to PowerPoint inches)
        // This is the ROOT FIX for dimension scaling issues
        this.scaleX = this.options.slideWidth / this.options.htmlWidth;
        this.scaleY = this.options.slideHeight / this.options.htmlHeight;
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
            
            // Get the main container
            const container = this.findMainContainer($);
            
            // Update HTML dimensions if specified in container
            this.updateDimensionsFromContainer($, container);
            
            // Recalculate scaling factors if dimensions changed
            this.scaleX = this.options.slideWidth / this.options.htmlWidth;
            this.scaleY = this.options.slideHeight / this.options.htmlHeight;
            
            // Process background
            this.processBackground($, slide, container);
            
            // Process elements - REDESIGNED to handle hierarchy properly
            await this.processElements($, slide, container);
            
            // Save presentation
            await this.pptx.writeFile({ fileName: outputPath });
            
            // Post-process to fix PptxGenJS corruption issues
            console.log('[HTML2PPTX] Post-processing PPTX to fix corruption issues...');
            await fixPPTX(outputPath);
            console.log('[HTML2PPTX] PPTX file fixed successfully');
            
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
            '.slide',
            '.container',
            'main',
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
     * Update HTML dimensions from container style
     * ROOT FIX: Extract actual dimensions from HTML instead of using hardcoded values
     */
    updateDimensionsFromContainer($, container) {
        const style = this.getComputedStyle($, container[0]);
        
        if (style.width) {
            const width = this.parsePixelValue(style.width);
            if (width > 0) {
                this.options.htmlWidth = width;
            }
        }
        
        if (style.height || style['min-height']) {
            const height = this.parsePixelValue(style.height || style['min-height']);
            if (height > 0) {
                this.options.htmlHeight = height;
            }
        }
    }
    
    /**
     * Process background (gradient or solid color)
     * ROOT FIX: Handle CSS gradients which were previously ignored
     */
    processBackground($, slide, container) {
        const style = this.getComputedStyle($, container[0]);
        
        if (style.background) {
            const bg = style.background;
            
            // Check for gradient
            if (bg.includes('linear-gradient') || bg.includes('radial-gradient')) {
                // Extract colors from gradient
                const colors = this.extractGradientColors(bg);
                
                if (colors.length >= 2) {
                    // PowerPoint doesn't support CSS gradients directly,
                    // so we'll use a solid color (middle of gradient) or create a shape
                    // For now, use the first color
                    slide.background = { color: this.parseColor(colors[0]) };
                }
            } else {
                // Solid color
                slide.background = { color: this.parseColor(bg) };
            }
        } else if (style['background-color']) {
            slide.background = { color: this.parseColor(style['background-color']) };
        }
    }
    
    /**
     * Extract colors from gradient string
     */
    extractGradientColors(gradient) {
        const colors = [];
        
        // Match hex colors
        const hexMatches = gradient.match(/#[0-9a-fA-F]{3,6}/g);
        if (hexMatches) {
            colors.push(...hexMatches);
        }
        
        // Match rgb/rgba colors
        const rgbMatches = gradient.match(/rgba?\([^)]+\)/g);
        if (rgbMatches) {
            colors.push(...rgbMatches);
        }
        
        return colors;
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
        
        // Get tag styles first (lowest priority)
        const tagName = $elem.prop('tagName')?.toLowerCase();
        if (tagName && this.styles[tagName]) {
            Object.assign(computedStyle, this.styles[tagName]);
        }
        
        // Get class styles
        const classes = ($elem.attr('class') || '').split(' ');
        classes.forEach(className => {
            if (className) {
                // Check parsed CSS first
                const classStyle = this.styles[`.${className}`];
                if (classStyle) {
                    Object.assign(computedStyle, classStyle);
                }
                
                // Handle common Tailwind classes (for CDN-loaded Tailwind)
                const tailwindStyle = this.getTailwindStyle(className);
                if (tailwindStyle) {
                    Object.assign(computedStyle, tailwindStyle);
                }
            }
        });
        
        // Check for nth-child pseudo-selectors
        const parent = $elem.parent();
        if (parent.length > 0) {
            const siblings = parent.children();
            const index = siblings.index($elem[0]);
            
            // Check all stored selectors for nth-child matches
            for (const selector in this.styles) {
                if (selector.includes(':nth-child')) {
                    // Extract the class and nth-child index
                    const match = selector.match(/(.+):nth-child\((\d+)\)/);
                    if (match) {
                        const baseSelector = match[1];
                        const nthIndex = parseInt(match[2]) - 1; // Convert to 0-based
                        
                        // Check if this element matches
                        if (index === nthIndex) {
                            // Check if base selector matches (class or tag)
                            let matches = false;
                            if (baseSelector.startsWith('.')) {
                                const className = baseSelector.substring(1);
                                matches = classes.includes(className);
                            } else if (baseSelector === tagName) {
                                matches = true;
                            }
                            
                            if (matches) {
                                Object.assign(computedStyle, this.styles[selector]);
                            }
                        }
                    }
                }
            }
        }
        
        // Get inline styles last (highest priority)
        const inlineStyle = $elem.attr('style');
        if (inlineStyle) {
            this.parseInlineStyle(inlineStyle, computedStyle);
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
     * Get Tailwind CSS utility class styles
     * Handles common Tailwind classes when CSS is loaded from CDN
     */
    getTailwindStyle(className) {
        const tailwindMap = {
            // Layout
            'flex': { 'display': 'flex' },
            'flex-row': { 'flex-direction': 'row' },
            'flex-col': { 'flex-direction': 'column' },
            'items-center': { 'align-items': 'center' },
            'items-start': { 'align-items': 'flex-start' },
            'items-end': { 'align-items': 'flex-end' },
            'justify-center': { 'justify-content': 'center' },
            'justify-start': { 'justify-content': 'flex-start' },
            'justify-end': { 'justify-content': 'flex-end' },
            'justify-between': { 'justify-content': 'space-between' },
            
            // Text alignment
            'text-center': { 'text-align': 'center' },
            'text-left': { 'text-align': 'left' },
            'text-right': { 'text-align': 'right' },
            
            // Gap (simplified - Tailwind has many gap-* classes)
            'gap-0': { 'gap': '0' },
            'gap-1': { 'gap': '0.25rem' },
            'gap-2': { 'gap': '0.5rem' },
            'gap-4': { 'gap': '1rem' },
            'gap-6': { 'gap': '1.5rem' },
            'gap-8': { 'gap': '2rem' },
            
            // Padding (simplified)
            'p-0': { 'padding': '0' },
            'p-4': { 'padding': '1rem' },
            'p-8': { 'padding': '2rem' },
            'p-16': { 'padding': '4rem' },
        };
        
        return tailwindMap[className] || null;
    }

    /**
     * Process all elements in the HTML
     * ROOT FIX: Completely redesigned to handle element hierarchy correctly
     * - Only processes leaf text elements (elements with text but no text-bearing children)
     * - Uses semantic understanding of HTML structure
     * - Properly handles centering and layout
     */
    async processElements($, slide, container) {
        // Find all leaf text elements (elements that contain text directly, not through children)
        const textElements = this.findLeafTextElements($, container);
        
        // Group elements by their semantic role and position
        const groupedElements = this.groupElementsByLayout($, textElements);
        
        // Process each element group
        for (const group of groupedElements) {
            await this.processTextElement($, slide, group);
        }
        
        // Process SVG elements (if any)
        this.processSVGElements($, slide, container);
    }
    
    /**
     * Find leaf text elements
     * ROOT FIX: Only find elements that directly contain text, not containers
     */
    findLeafTextElements($, container) {
        const leafElements = [];
        
        // Text-bearing elements to consider
        const textTags = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'span', 'div', 'li', 'td', 'th', 'label', 'a'];
        
        container.find('*').each((i, elem) => {
            const $elem = $(elem);
            const tagName = $elem.prop('tagName')?.toLowerCase();
            
            // Skip non-text elements
            if (!textTags.includes(tagName)) {
                return;
            }
            
            // Get direct text content (not from children)
            const directText = this.getDirectText($, $elem);
            
            // Check if this element has text-bearing children
            const hasTextChildren = $elem.children().filter((j, child) => {
                const childTag = $(child).prop('tagName')?.toLowerCase();
                return textTags.includes(childTag) && $(child).text().trim().length > 0;
            }).length > 0;
            
            // If element has direct text OR has text but no text-bearing children, it's a leaf
            if (directText && !hasTextChildren) {
                leafElements.push({
                    elem: $elem,
                    tagName: tagName,
                    text: $elem.text().trim()
                });
            } else if (!hasTextChildren && $elem.text().trim().length > 0) {
                // Element with only text (no child elements with text)
                leafElements.push({
                    elem: $elem,
                    tagName: tagName,
                    text: $elem.text().trim()
                });
            }
        });
        
        return leafElements;
    }
    
    /**
     * Group elements by their layout position
     * ROOT FIX: Better understanding of element relationships
     */
    groupElementsByLayout($, elements) {
        // For now, return elements as-is
        // In the future, we could group elements that should be in the same text box
        return elements;
    }
    
    /**
     * Process a single text element
     * ROOT FIX: Simplified positioning using proper scaling
     */
    async processTextElement($, slide, elementInfo) {
        const { elem: $elem, tagName, text } = elementInfo;
        
        if (!text || text.length === 0) return;
        
        const style = this.getComputedStyle($, $elem[0]);
        
        // Calculate position using improved method
        const position = this.calculateElementPosition($, $elem, style);
        
        if (!position) {
            console.warn(`Could not calculate position for element: ${text.substring(0, 50)}...`);
            return;
        }
        
        // Parse text formatting
        const textOptions = this.parseTextOptions($, $elem, style, tagName);
        
        try {
            // Add text box to slide
            slide.addText(text, {
                x: position.x,
                y: position.y,
                w: position.w,
                h: position.h,
                ...textOptions
            });
        } catch (e) {
            console.warn(`Error adding text element: ${e.message}`);
        }
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
     * Calculate element position and size
     * ROOT FIX: Completely rewritten to properly track document flow and sibling positions
     * - Tracks previous siblings to calculate vertical stacking
     * - Properly handles nested containers
     * - Distinguishes container centering from element positioning
     */
    calculateElementPosition($, $elem, style) {
        // Strategy: Calculate absolute position by:
        // 1. Walking up to find positioning contexts (absolute, flex containers)
        // 2. Calculating offset from previous siblings in document flow
        // 3. Accumulating positions from parent chain
        // 4. Converting to PowerPoint coordinates using scaling
        
        let x = 0, y = 0, w = 0, h = 0;
        
        // Get element's own dimensions
        const fontSize = this.parseFontSize(style['font-size'] || '16px');
        const lineHeight = style['line-height'] ? this.parsePixelValue(style['line-height']) : fontSize * 1.2;
        
        // Estimate element dimensions based on content
        const text = $elem.text().trim();
        const estimatedLines = Math.ceil(text.length / 50); // Rough estimate
        
        // Get parent and parent style
        const parent = $elem.parent();
        const parentStyle = parent.length > 0 ? this.getComputedStyle($, parent[0]) : {};
        
        // Calculate width and height first
        if (style.width) {
            w = this.parsePixelValue(style.width);
        } else if (parentStyle.width) {
            const parentWidth = this.parsePixelValue(parentStyle.width);
            const parentPadding = this.parsePixelValue(parentStyle.padding || '0') * 2;
            w = parentWidth - parentPadding;
        } else {
            w = this.options.htmlWidth * 0.8; // 80% of slide width by default
        }
        
        if (style.height) {
            h = this.parsePixelValue(style.height);
        } else {
            h = lineHeight * estimatedLines + fontSize * 0.5; // Add some padding
        }
        
        // Calculate position - NEW LOGIC that tracks document flow
        const position = this.calculateAbsolutePosition($, $elem, parentStyle, w, h);
        x = position.x;
        y = position.y;
        

        
        // Handle flex-specific dimensions if needed
        if (parentStyle.display === 'flex' && parentStyle['flex-direction'] === 'column') {
            const gap = this.parsePixelValue(parentStyle.gap || '0');
            const siblings = parent.children();
            
            if (style.flex === '1' || style.flex === '1 1 0%' || style.flex) {
                // Calculate flex height
                const flexCount = siblings.filter((j, el) => {
                    const s = this.getComputedStyle($, el);
                    return s.flex === '1' || s.flex === '1 1 0%' || s.flex;
                }).length;
                
                const parentHeight = this.parsePixelValue(parentStyle.height || parentStyle['min-height'] || '720px');
                const parentPadding = this.parsePixelValue(parentStyle.padding || '0') * 2;
                const totalGaps = gap * (siblings.length - 1);
                const availableHeight = parentHeight - parentPadding - totalGaps;
                
                h = availableHeight / flexCount;
            }
        }
        
        // Apply scaling to convert HTML pixels to PowerPoint inches
        return {
            x: x * this.scaleX,
            y: y * this.scaleY,
            w: w * this.scaleX,
            h: h * this.scaleY
        };
    }
    
    /**
     * Calculate absolute position of an element in HTML coordinate space
     * NEW METHOD: Properly tracks document flow and sibling offsets
     */
    calculateAbsolutePosition($, $elem, parentStyle, elemWidth, elemHeight) {
        let x = 0, y = 0;

        // Calculate offset from previous siblings in document flow
        const parent = $elem.parent();
        if (parent.length > 0) {
            const siblings = parent.children();
            const index = siblings.index($elem[0]);

            // Calculate Y offset from previous siblings
            if (index > 0) {
                const gap = this.parsePixelValue(parentStyle.gap || '0');

                for (let i = 0; i < index; i++) {
                    const sibling = siblings.eq(i);
                    const siblingStyle = this.getComputedStyle($, sibling[0]);

                    // Calculate sibling height
                    let siblingHeight = 0;
                    if (siblingStyle.height) {
                        siblingHeight = this.parsePixelValue(siblingStyle.height);
                    } else {
                        const siblingText = sibling.text().trim();
                        const siblingFontSize = this.parseFontSize(siblingStyle['font-size'] || '16px');
                        const siblingLineHeight = siblingStyle['line-height'] ?
                            this.parsePixelValue(siblingStyle['line-height']) : siblingFontSize * 1.2;
                        const siblingLines = Math.ceil(siblingText.length / 50);
                        siblingHeight = siblingLineHeight * siblingLines + siblingFontSize * 0.5;
                    }

                    // Add sibling height and margin
                    y += siblingHeight;

                    // Add margin-bottom of sibling
                    if (siblingStyle['margin-bottom']) {
                        y += this.parsePixelValue(siblingStyle['margin-bottom']);
                    }

                    // Add margin-top of current element (only once, after last sibling)
                    if (i === index - 1) {
                        const currentStyle = this.getComputedStyle($, $elem[0]);
                        if (currentStyle['margin-top']) {
                            y += this.parsePixelValue(currentStyle['margin-top']);
                        }
                    }

                    // Add gap if parent is flex
                    if (parentStyle.display === 'flex' && parentStyle['flex-direction'] === 'column') {
                        y += gap;
                    }
                }
            }
        }

        // Add parent padding
        if (parentStyle.padding) {
            const padding = this.parsePixelValue(parentStyle.padding);
            x += padding;
            y += padding;
        }

        // Recursively add parent's position
        if (parent.length > 0 && parent.prop('tagName')?.toLowerCase() !== 'body') {
            const grandParent = parent.parent();
            const grandParentStyle = grandParent.length > 0 ? this.getComputedStyle($, grandParent[0]) : {};
            const parentComputedStyle = this.getComputedStyle($, parent[0]);

            // Get parent width and height for calculations
            const parentWidth = this.parsePixelValue(parentComputedStyle.width || this.options.htmlWidth.toString());
            const parentHeight = this.parsePixelValue(parentComputedStyle.height || parentComputedStyle['min-height'] || '0');

            const parentPos = this.calculateAbsolutePosition($, parent, grandParentStyle, parentWidth, parentHeight);
            x += parentPos.x;
            y += parentPos.y;
        }

        // Handle text-align center (horizontal centering within parent)
        const currentStyle = this.getComputedStyle($, $elem[0]);
        if (currentStyle['text-align'] === 'center' || parentStyle['text-align'] === 'center') {
            // Center horizontally within parent
            const parentWidth = this.parsePixelValue(parentStyle.width || this.options.htmlWidth.toString());
            const parentPadding = this.parsePixelValue(parentStyle.padding || '0') * 2;
            const availableWidth = parentWidth - parentPadding;

            // Adjust x to center the element
            x = x - this.parsePixelValue(parentStyle.padding || '0') + (availableWidth - elemWidth) / 2;
        }

        // Handle flex centering at container level
        // If the current element's parent is centered, adjust the position
        if (parent.length > 0 && parent.prop('tagName')?.toLowerCase() !== 'body') {
            const grandParent = parent.parent();
            const grandParentStyle = grandParent.length > 0 ? this.getComputedStyle($, grandParent[0]) : {};

            // Check if parent is centered by its parent (flex container)
            if (grandParentStyle.display === 'flex') {
                const grandParentWidth = this.parsePixelValue(grandParentStyle.width || this.options.htmlWidth.toString());
                const grandParentHeight = this.parsePixelValue(grandParentStyle.height || grandParentStyle['min-height'] || this.options.htmlHeight.toString());
                const parentComputedStyle = this.getComputedStyle($, parent[0]);
                const parentActualWidth = this.parsePixelValue(parentComputedStyle.width || grandParentWidth.toString());

                // Horizontal centering from grandparent flex
                if (grandParentStyle['justify-content'] === 'center' || grandParentStyle['align-items'] === 'center') {
                    // The parent is centered, so adjust x accordingly
                    // x is already relative to parent, so no adjustment needed for horizontal
                    // But we need to ensure elements are centered within their immediate parent
                }

                // Vertical centering from grandparent flex
                if (grandParentStyle['justify-content'] === 'center' && grandParentStyle['flex-direction'] === 'column') {
                    // Parent is vertically centered in grandparent
                    // Calculate total height of all siblings
                    const siblings = parent.children();
                    let totalContentHeight = 0;
                    siblings.each((i, sib) => {
                        const sibStyle = this.getComputedStyle($, sib);
                        const sibText = $(sib).text().trim();
                        const sibFontSize = this.parseFontSize(sibStyle['font-size'] || '16px');
                        const sibLineHeight = sibStyle['line-height'] ?
                            this.parsePixelValue(sibStyle['line-height']) : sibFontSize * 1.2;
                        const sibLines = Math.ceil(sibText.length / 50);
                        const sibHeight = sibStyle.height ?
                            this.parsePixelValue(sibStyle.height) :
                            sibLineHeight * sibLines + sibFontSize * 0.5;
                        totalContentHeight += sibHeight;

                        // Add margins
                        if (sibStyle['margin-top']) {
                            totalContentHeight += this.parsePixelValue(sibStyle['margin-top']);
                        }
                        if (sibStyle['margin-bottom']) {
                            totalContentHeight += this.parsePixelValue(sibStyle['margin-bottom']);
                        }
                    });

                    // Center the block vertically
                    const verticalOffset = (grandParentHeight - totalContentHeight) / 2;
                    if (verticalOffset > 0) {
                        y += verticalOffset;
                    }
                }
            }
        }

        return { x, y };
    }

    /**
     * Parse text formatting options
     * ROOT FIX: Better font size scaling and semantic understanding of HTML tags
     */
    parseTextOptions($, $elem, style, tagName) {
        const options = {
            align: 'center', // Default to center for better appearance
            valign: 'middle',
            fontSize: 18,
            color: 'FFFFFF', // Default to white (common for dark backgrounds)
            bold: false,
            italic: false,
            fontFace: 'Arial',
            autoFit: true, // Let PowerPoint adjust font size if needed
            shrinkText: true
        };
        
        // Semantic font sizes based on HTML tags
        // ROOT FIX: Use proper font size scaling that maintains visual hierarchy
        const tagFontSizes = {
            'h1': 44,
            'h2': 36, 
            'h3': 28,
            'h4': 24,
            'h5': 20,
            'h6': 18,
            'p': 18,
            'span': 18,
            'div': 18
        };
        
        // Start with semantic default
        if (tagFontSizes[tagName]) {
            options.fontSize = tagFontSizes[tagName];
        }
        
        // Override with actual style if present
        if (style['font-size']) {
            options.fontSize = this.parseFontSize(style['font-size']);
        }
        
        // Color
        if (style.color) {
            options.color = this.parseColor(style.color);
        }
        
        // Background color - don't set by default to avoid white boxes
        if (style['background-color'] && style['background-color'] !== 'transparent') {
            options.fill = { color: this.parseColor(style['background-color']) };
        } else if (style.background && !style.background.includes('gradient') && style.background !== 'transparent') {
            options.fill = { color: this.parseColor(style.background) };
        }
        
        // Border - handle shorthand and individual properties
        let borderColor = null;
        let borderWidth = 1;
        
        if (style.border && style.border !== 'none') {
            borderColor = this.extractColorFromBorder(style.border);
            borderWidth = this.parseBorderWidth(style.border);
        }
        
        // Individual border properties override shorthand
        if (style['border-color']) {
            borderColor = style['border-color'];
        }
        if (style['border-width']) {
            borderWidth = this.parseBorderWidth(style['border-width']);
        }
        
        if (borderColor && borderColor !== 'transparent') {
            options.line = {
                color: this.parseColor(borderColor),
                width: borderWidth,
                dashType: this.parseBorderStyle(style.border || style['border-style'])
            };
        }
        
        // Text alignment
        if (style['text-align']) {
            options.align = style['text-align'];
        }
        
        // Check parent for flex alignment
        const parent = $elem.parent();
        if (parent.length > 0) {
            const parentStyle = this.getComputedStyle($, parent[0]);
            
            if (parentStyle['justify-content']) {
                const justify = parentStyle['justify-content'];
                if (justify === 'center') options.align = 'center';
                if (justify === 'flex-end') options.align = 'right';
                if (justify === 'flex-start') options.align = 'left';
            }
            
            if (parentStyle['align-items']) {
                const alignItems = parentStyle['align-items'];
                if (alignItems === 'center') options.valign = 'middle';
                if (alignItems === 'flex-start') options.valign = 'top';
                if (alignItems === 'flex-end') options.valign = 'bottom';
            }
        }
        
        // Font weight
        if (style['font-weight']) {
            const weight = parseInt(style['font-weight']) || 400;
            options.bold = weight >= 600 || style['font-weight'] === 'bold';
        }
        
        // Semantic bold (h1, h2, etc. are often bold)
        if (['h1', 'h2'].includes(tagName) && !style['font-weight']) {
            options.bold = true;
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
            const borderRadius = this.parsePixelValue(style['border-radius']);
            if (borderRadius > 0) {
                options.shape = this.pptx.ShapeType.roundRect;
                options.rectRadius = borderRadius * this.scaleX; // Scale border radius
            }
        }
        
        return options;
    }

    /**
     * Parse position value (px, %, etc.) to points
     * Kept for backward compatibility but prefer parsePixelValue
     */
    parsePosition(value) {
        return this.parsePixelValue(value);
    }
    
    /**
     * Parse CSS value to pixels
     * ROOT FIX: Better unit conversion
     */
    parsePixelValue(value) {
        if (!value) return 0;
        
        const strValue = String(value);
        const numValue = parseFloat(strValue);
        
        if (isNaN(numValue)) return 0;
        
        if (strValue.includes('px')) {
            return numValue;
        } else if (strValue.includes('%')) {
            // Percentage relative to container width
            return (numValue / 100) * this.options.htmlWidth;
        } else if (strValue.includes('em')) {
            return numValue * 16; // 1em = 16px typically
        } else if (strValue.includes('rem')) {
            return numValue * 16; // 1rem = 16px typically
        } else if (strValue.includes('pt')) {
            // Points to pixels (1pt = 1.333px approximately)
            return numValue * 1.333;
        } else if (strValue.includes('vh')) {
            return (numValue / 100) * this.options.htmlHeight;
        } else if (strValue.includes('vw')) {
            return (numValue / 100) * this.options.htmlWidth;
        }
        
        // Default: treat as pixels
        return numValue;
    }

    /**
     * Parse font size to PowerPoint points
     * ROOT FIX: Proper font size scaling instead of fixed multiplier
     */
    parseFontSize(fontSize) {
        if (!fontSize) return 18;
        
        const strValue = String(fontSize);
        const numValue = parseFloat(strValue);
        
        if (isNaN(numValue)) return 18;
        
        // PowerPoint uses points for font size
        // We want to maintain the visual appearance, not do a literal px->pt conversion
        // The key insight: HTML px on screen ~= PowerPoint pt in presentation
        // So we use a scaling factor that maintains relative sizes
        
        if (strValue.includes('px')) {
            // Use a scaling factor based on typical screen DPI (96) and PowerPoint DPI (72)
            // But we want to preserve visual hierarchy, so we use a gentler scaling
            return Math.round(numValue * 0.85); // Slightly reduce to fit better in slides
        } else if (strValue.includes('em')) {
            // 1em = 16px typically, then convert to pt
            return Math.round(numValue * 16 * 0.85);
        } else if (strValue.includes('rem')) {
            return Math.round(numValue * 16 * 0.85);
        } else if (strValue.includes('pt')) {
            // Already in points
            return Math.round(numValue);
        }
        
        // Default: assume it's a number in points
        return Math.round(numValue);
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
     * Parse border style
     */
    parseBorderStyle(borderStyle) {
        if (!borderStyle) return 'solid';
        if (borderStyle.includes('dashed')) return 'dash';
        if (borderStyle.includes('dotted')) return 'dot';
        if (borderStyle.includes('dashdot')) return 'dashDot';
        if (borderStyle.includes('longdash')) return 'lgDash';
        if (borderStyle.includes('longdashdot')) return 'lgDashDot';
        if (borderStyle.includes('longdashdotdot')) return 'lgDashDotDot';
        return 'solid';
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
