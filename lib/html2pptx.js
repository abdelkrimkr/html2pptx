const cheerio = require('cheerio');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
const css = require('css');
const { fixPPTX } = require('./pptx-fixer');
const { getElementDimensions } = require('./dimension-fetcher');

class HTML2PPTX {
    constructor(options = {}) {
        this.options = {
            slideWidth: 10,
            slideHeight: 5.625,
            htmlWidth: 1280,
            htmlHeight: 720,
            background: { color: 'FFFFFF' },
            ...options
        };
        this.pptx = new PptxGenJS();
        this.styles = {};
        this.scaleX = this.options.slideWidth / this.options.htmlWidth;
        this.scaleY = this.options.slideHeight / this.options.htmlHeight;
    }

    async convert(inputPath, outputPath) {
        try {
            const html = fs.readFileSync(inputPath, 'utf8');
            const $ = cheerio.load(html);
            this.extractCSS($);
            this.pptx.layout = 'LAYOUT_16x9';
            this.pptx.author = 'HTML2PPTX Converter';
            this.pptx.title = $('title').text() || 'Converted Presentation';
            const slide = this.pptx.addSlide({
                margin: 0
            });
            const container = this.findMainContainer($);
            this.updateDimensionsFromContainer($, container);
            this.scaleX = this.options.slideWidth / this.options.htmlWidth;
            this.scaleY = this.options.slideHeight / this.options.htmlHeight;
            this.processBackground($, slide, container);
            
            const elementDimensions = await getElementDimensions(inputPath);
            await this.processElements(slide, elementDimensions);

            await this.pptx.writeFile({ fileName: outputPath });
            await fixPPTX(outputPath);
            return { success: true, outputPath: outputPath };
        } catch (error) {
            console.error('Conversion failed:', error);
            throw new Error(`Conversion failed: ${error.message}`);
        }
    }

    findMainContainer($) {
        const selectors = ['.slide-container', '.slide', '.container', 'main', 'body > div:first-child', 'body'];
        for (const selector of selectors) {
            const elem = $(selector);
            if (elem.length > 0) return elem.first();
        }
        return $('body');
    }

    updateDimensionsFromContainer($, container) {
        const style = this.getComputedStyle($, container[0]);
        if (style.width) {
            const width = this.parsePixelValue(style.width);
            if (width > 0) this.options.htmlWidth = width;
        }
        if (style.height || style['min-height']) {
            const height = this.parsePixelValue(style.height || style['min-height']);
            if (height > 0) this.options.htmlHeight = height;
        }
    }

    processBackground($, slide, container) {
        const style = this.getComputedStyle($, container[0]);
        const bg = style.background || style['background-color'];
        if (bg) {
            if (bg.includes('linear-gradient') || bg.includes('radial-gradient')) {
                const colors = this.extractGradientColors(bg);
                if (colors.length >= 2) slide.background = { color: this.parseColor(colors[0]) };
            } else {
                slide.background = { color: this.parseColor(bg) };
            }
        }
    }

    extractGradientColors(gradient) {
        const colors = [];
        const hexMatches = gradient.match(/#[0-9a-fA-F]{3,6}/g);
        if (hexMatches) colors.push(...hexMatches);
        const rgbMatches = gradient.match(/rgba?\([^)]+\)/g);
        if (rgbMatches) colors.push(...rgbMatches);
        return colors;
    }

    extractCSS($) {
        $('style').each((i, elem) => {
            try {
                this.processCSSRules(css.parse($(elem).html()).stylesheet.rules);
            } catch (e) {
                console.warn('CSS parsing warning:', e.message);
            }
        });
    }

    processCSSRules(rules) {
        rules.forEach(rule => {
            if (rule.type === 'rule') {
                const styles = {};
                rule.declarations.forEach(decl => {
                    if (decl.type === 'declaration') styles[decl.property] = decl.value;
                });
                rule.selectors.forEach(selector => {
                    if (!this.styles[selector]) this.styles[selector] = {};
                    Object.assign(this.styles[selector], styles);
                });
            }
        });
    }

    getComputedStyle($, elem) {
        const $elem = $(elem);
        const computedStyle = {};
        const tagName = $elem.prop('tagName')?.toLowerCase();
        if (tagName && this.styles[tagName]) Object.assign(computedStyle, this.styles[tagName]);
        ($elem.attr('class') || '').split(' ').forEach(className => {
            if (className) {
                const classStyle = this.styles[`.${className}`];
                if (classStyle) Object.assign(computedStyle, classStyle);
            }
        });
        const inlineStyle = $elem.attr('style');
        if (inlineStyle) this.parseInlineStyle(inlineStyle, computedStyle);
        return computedStyle;
    }

    parseInlineStyle(styleStr, target) {
        styleStr.split(';').forEach(part => {
            const [prop, value] = part.split(':').map(s => s.trim());
            if (prop && value) target[prop] = value;
        });
    }

    async processElements(slide, elements) {
        for (const elem of elements) {
            if (!elem || elem.w === 0 || elem.h === 0) continue;

            const finalPosition = {
                x: elem.x * this.scaleX,
                y: elem.y * this.scaleY,
                w: elem.w * this.scaleX,
                h: elem.h * this.scaleY,
            };

            const textOptions = this.parseTextOptions(elem.style, elem.tag.toLowerCase(), finalPosition, elem.hyperlink);
            
            try {
                if (elem.text && elem.text.trim().length > 0) {
                    slide.addText(elem.text.trim(), { ...finalPosition, ...textOptions });
                }
            } catch (e) {
                console.warn(`Error adding element: ${e.message}`);
            }
        }
    }

    parseTextOptions(style, tagName, position, hyperlink) {
        const options = {
            align: 'center', valign: 'middle', fontSize: 18, color: '000000',
            bold: false, italic: false, fontFace: 'Arial', autoFit: true, shrinkText: true
        };

        const tagFontSizes = { 'h1': 44, 'h2': 36, 'h3': 28, 'h4': 24, 'h5': 20, 'h6': 18, 'p': 18, 'span': 18, 'div': 18, 'a': 18 };
        if (tagFontSizes[tagName]) options.fontSize = tagFontSizes[tagName];
        
        if (style.fontSize) options.fontSize = this.parseFontSize(style.fontSize);
        if (style.color) options.color = this.parseColor(style.color);
        if (style.backgroundColor && style.backgroundColor !== 'transparent' && style.backgroundColor !== 'rgba(0, 0, 0, 0)') {
            options.fill = { color: this.parseColor(style.backgroundColor) };
        }

        let borderColor = null;
        let borderWidth = this.parseBorderWidth(style.border);
        
        if (style.border && style.border !== 'none') {
            borderColor = this.extractColorFromBorder(style.border);
        }
        
        if (borderColor && borderColor !== 'transparent' && borderWidth > 0) {
            options.line = {
                color: this.parseColor(borderColor),
                width: borderWidth * 0.75, // Convert px to points
                dashType: this.parseBorderStyle(style.border)
            };
        }

        if (style.textAlign) options.align = style.textAlign;
        if (style.fontWeight) options.bold = (parseInt(style.fontWeight) || 400) >= 600 || style.fontWeight === 'bold';
        if (['h1', 'h2'].includes(tagName) && !style.fontWeight) options.bold = true;
        if (style.fontStyle === 'italic') options.italic = true;
        if (style.fontFamily) options.fontFace = this.parseFontFamily(style.fontFamily);
        if (style.borderRadius) {
            const borderRadius = this.parsePixelValue(style.borderRadius);
            if (borderRadius > 0) {
                options.shape = this.pptx.ShapeType.roundRect;
                options.rectRadius = borderRadius * this.scaleX;
            }
        }
        if (style.transform) {
            const rotateMatch = style.transform.match(/rotate\(([^)]+)\)/);
            if (rotateMatch && rotateMatch[1]) {
                const angle = parseFloat(rotateMatch[1]);
                if (!isNaN(angle)) options.rotate = angle;
            }
            const scaleMatch = style.transform.match(/scale\(([^)]+)\)/);
            if (scaleMatch && scaleMatch[1]) {
                const scale = parseFloat(scaleMatch[1]);
                if (!isNaN(scale)) {
                    const newWidth = position.w * scale;
                    const newHeight = position.h * scale;
                    position.x += (position.w - newWidth) / 2;
                    position.y += (position.h - newHeight) / 2;
                    position.w = newWidth;
                    position.h = newHeight;
                }
            }
        }
        if (hyperlink) options.hyperlink = { url: hyperlink };
        return options;
    }

    parsePixelValue(value, parentDimension) {
        if (!value) return 0;
        const strValue = String(value);
        const numValue = parseFloat(strValue);
        if (isNaN(numValue)) return 0;
        if (strValue.includes('px')) return numValue;
        if (strValue.includes('%')) return (numValue / 100) * (parentDimension || this.options.htmlWidth);
        if (strValue.includes('em') || strValue.includes('rem')) return numValue * 16;
        if (strValue.includes('pt')) return numValue * 1.333;
        if (strValue.includes('vh')) return (numValue / 100) * this.options.htmlHeight;
        if (strValue.includes('vw')) return (numValue / 100) * this.options.htmlWidth;
        return numValue;
    }

    parseFontSize(fontSize) {
        if (!fontSize) return 18;
        const strValue = String(fontSize);
        const numValue = parseFloat(strValue);
        if (isNaN(numValue)) return 18;
        if (strValue.includes('px')) return Math.round(numValue * 0.75);
        if (strValue.includes('em') || strValue.includes('rem')) return Math.round(numValue * 16 * 0.75);
        if (strValue.includes('pt')) return Math.round(numValue);
        return Math.round(numValue);
    }

    parseColor(color) {
        if (!color) return '000000';
        color = color.replace('#', '').trim();
        if (color.startsWith('rgb')) {
            const matches = color.match(/\d+/g);
            if (matches && matches.length >= 3) {
                const r = parseInt(matches[0]).toString(16).padStart(2, '0');
                const g = parseInt(matches[1]).toString(16).padStart(2, '0');
                const b = parseInt(matches[2]).toString(16).padStart(2, '0');
                return r + g + b;
            }
        }
        const namedColors = { 'white': 'FFFFFF', 'black': '000000', 'red': 'FF0000', 'green': '00FF00', 'blue': '0000FF', 'gray': '808080', 'grey': '808080', 'transparent': 'FFFFFF' };
        if (namedColors[color.toLowerCase()]) return namedColors[color.toLowerCase()];
        if (/^[0-9A-Fa-f]{6}$/.test(color)) return color.toUpperCase();
        if (/^[0-9A-Fa-f]{3}$/.test(color)) return color.split('').map(c => c + c).join('').toUpperCase();
        return '000000';
    }

    extractColorFromBorder(border) {
        if (!border) return '#000000';
        const parts = border.split(' ');
        for (const part of parts) {
            if (part.startsWith('#') || part.startsWith('rgb')) return part;
        }
        const namedColors = ['red', 'blue', 'green', 'black', 'white', 'gray', 'grey'];
        for (const color of namedColors) {
            if (border.includes(color)) return color;
        }
        return '#000000';
    }

    parseBorderWidth(border) {
        if (!border || border === 'none') return 0;
        const parts = String(border).split(' ');
        const widthPart = parts.find(p => p.endsWith('px'));
        if (widthPart) {
            const width = parseFloat(widthPart);
            return isNaN(width) ? 0 : width;
        }
        return 0;
    }

    parseBorderStyle(borderStyle) {
        if (!borderStyle) return 'solid';
        if (borderStyle.includes('dashed')) return 'dash';
        if (borderStyle.includes('dotted')) return 'dot';
        return 'solid';
    }

    parseFontFamily(fontFamily) {
        if (!fontFamily) return 'Arial';
        const fonts = fontFamily.split(',');
        let font = fonts[0].trim().replace(/['"]/g, '');
        const fontMap = { 'Roboto': 'Arial', 'Montserrat': 'Arial', 'Helvetica': 'Arial', 'sans-serif': 'Arial', 'serif': 'Times New Roman', 'monospace': 'Courier New' };
        return fontMap[font] || font || 'Arial';
    }
}

async function convertHTML2PPTX(inputPath, outputPath, options = {}) {
    const converter = new HTML2PPTX(options);
    return await converter.convert(inputPath, outputPath);
}

module.exports = { HTML2PPTX, convertHTML2PPTX };