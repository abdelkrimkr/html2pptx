const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

async function getElementDimensions(htmlFilePath) {
    let browser;
    try {
        browser = await puppeteer.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] });
        const page = await browser.newPage();

        const absolutePath = path.resolve(htmlFilePath);
        const htmlContent = fs.readFileSync(absolutePath, 'utf8');

        await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

        const dimensions = await page.evaluate(() => {
            // Force border-box sizing for predictable layout calculations
            const styleNode = document.createElement('style');
            styleNode.innerHTML = 'body, div, span, p, h1, h2, h3, h4, h5, h6, a, li { box-sizing: border-box; }';
            document.head.appendChild(styleNode);

            const container = document.querySelector('.slide');
            if (!container) return [];

            const containerRect = container.getBoundingClientRect();
            const containerStyle = window.getComputedStyle(container);
            const paddingLeft = parseFloat(containerStyle.paddingLeft) || 0;
            const paddingTop = parseFloat(containerStyle.paddingTop) || 0;

            // Use querySelectorAll to get all potentially relevant elements
            const elements = Array.from(container.querySelectorAll('div, span, p, h1, h2, h3, h4, h5, h6, a, li'));

            return elements.filter(el => {
                // Robust filtering for visible elements within the slide container
                if (el === container) return false; // Exclude the container itself
                if (!container.contains(el)) return false; // Must be a descendant
                const style = window.getComputedStyle(el);
                if (style.display === 'none' || style.visibility === 'hidden' || parseFloat(style.opacity) === 0) {
                    return false;
                }
                const rect = el.getBoundingClientRect();
                if (rect.width === 0 || rect.height === 0) {
                    return false; // Exclude elements with no area
                }
                return true;
            }).map(el => {
                const rect = el.getBoundingClientRect();
                const style = window.getComputedStyle(el);
                return {
                    tag: el.tagName,
                    id: el.id,
                    className: el.className,
                    x: rect.left - containerRect.left - paddingLeft,
                    y: rect.top - containerRect.top - paddingTop,
                    w: rect.width,
                    h: rect.height,
                    text: el.innerText,
                    style: {
                        color: style.color,
                        backgroundColor: style.backgroundColor,
                        fontSize: style.fontSize,
                        fontWeight: style.fontWeight,
                        fontStyle: style.fontStyle,
                        fontFamily: style.fontFamily,
                        textAlign: style.textAlign,
                        border: style.border,
                        borderRadius: style.borderRadius,
                        transform: style.transform,
                    },
                    hyperlink: el.closest('a') ? el.closest('a').href : null,
                };
            });
        });

        return dimensions;
    } finally {
        if (browser) {
            await browser.close();
        }
    }
}

module.exports = { getElementDimensions };