const AdmZip = require('adm-zip');
const xml2js = require('xml2js');

class PptxVerifier {
    constructor(pptxPath) {
        this.pptxPath = pptxPath;
        this.zip = new AdmZip(pptxPath);
        this.parser = new xml2js.Parser({ explicitArray: false });
    }

    async getSlide(slideNumber) {
        const slideEntry = this.zip.getEntry(`ppt/slides/slide${slideNumber}.xml`);
        if (!slideEntry) {
            throw new Error(`Slide ${slideNumber} not found.`);
        }
        const slideXml = slideEntry.getData().toString('utf8');
        return await this.parser.parseStringPromise(slideXml);
    }

    async getShapeCount(slideNumber) {
        const slide = await this.getSlide(slideNumber);
        const shapes = slide['p:sld']['p:cSld']['p:spTree']['p:sp'];
        return Array.isArray(shapes) ? shapes.length : (shapes ? 1 : 0);
    }

    async getShapeText(slideNumber, shapeIndex) {
        const slide = await this.getSlide(slideNumber);
        const shape = slide['p:sld']['p:cSld']['p:spTree']['p:sp'][shapeIndex];
        if (!shape) {
            throw new Error(`Shape ${shapeIndex} not found on slide ${slideNumber}.`);
        }
        const textBody = shape['p:txBody'];
        if (!textBody) {
            return '';
        }
        const paragraphs = textBody['a:p'];
        if (!paragraphs) {
            return '';
        }
        if (Array.isArray(paragraphs)) {
            return paragraphs.map(p => p['a:r']['a:t']).join('\n');
        }
        return paragraphs['a:r']['a:t'];
    }

    async getShapePosition(slideNumber, shapeIndex) {
        const slide = await this.getSlide(slideNumber);
        const shape = slide['p:sld']['p:cSld']['p:spTree']['p:sp'][shapeIndex];
        if (!shape) {
            throw new Error(`Shape ${shapeIndex} not found on slide ${slideNumber}.`);
        }
        const xfrm = shape['p:spPr']['a:xfrm'];
        return {
            x: parseInt(xfrm['a:off']['$']['x']),
            y: parseInt(xfrm['a:off']['$']['y']),
            w: parseInt(xfrm['a:ext']['$']['cx']),
            h: parseInt(xfrm['a:ext']['$']['cy']),
        };
    }

    async findShapeByText(slideNumber, text) {
        const slide = await this.getSlide(slideNumber);
        const shapes = slide['p:sld']['p:cSld']['p:spTree']['p:sp'];
        const shapeArray = Array.isArray(shapes) ? shapes : [shapes];

        for (const shape of shapeArray) {
            const textBody = shape['p:txBody'];
            if (textBody) {
                const paragraphs = textBody['a:p'];
                if (paragraphs) {
                    const shapeText = Array.isArray(paragraphs)
                        ? paragraphs.map(p => p['a:r']['a:t']).join('\n')
                        : paragraphs['a:r']['a:t'];

                    if (shapeText === text) {
                        const xfrm = shape['p:spPr']['a:xfrm'];
                        return {
                            x: parseInt(xfrm['a:off']['$']['x']),
                            y: parseInt(xfrm['a:off']['$']['y']),
                            w: parseInt(xfrm['a:ext']['$']['cx']),
                            h: parseInt(xfrm['a:ext']['$']['cy']),
                        };
                    }
                }
            }
        }
        return null;
    }

    async verifyNoOverlaps(slideNumber) {
        const slide = await this.getSlide(slideNumber);
        const shapes = slide['p:sld']['p:cSld']['p:spTree']['p:sp'];
        const shapeArray = Array.isArray(shapes) ? shapes : [shapes];
        const positions = [];

        for (const shape of shapeArray) {
            const xfrm = shape['p:spPr']['a:xfrm'];
            positions.push({
                x: parseInt(xfrm['a:off']['$']['x']),
                y: parseInt(xfrm['a:off']['$']['y']),
                w: parseInt(xfrm['a:ext']['$']['cx']),
                h: parseInt(xfrm['a:ext']['$']['cy']),
            });
        }

        for (let i = 0; i < positions.length; i++) {
            for (let j = i + 1; j < positions.length; j++) {
                const pos1 = positions[i];
                const pos2 = positions[j];

                const overlap = !(pos1.x + pos1.w < pos2.x ||
                                  pos2.x + pos2.w < pos1.x ||
                                  pos1.y + pos1.h < pos2.y ||
                                  pos2.y + pos2.h < pos1.y);

                if (overlap) {
                    return false; // Found an overlap
                }
            }
        }

        return true; // No overlaps found
    }

    async verifyContentFit(slideNumber) {
        const slide = await this.getSlide(slideNumber);
        const shapes = slide['p:sld']['p:cSld']['p:spTree']['p:sp'];
        const shapeArray = Array.isArray(shapes) ? shapes : [shapes];

        for (const shape of shapeArray) {
            const textBody = shape['p:txBody'];
            if (textBody) {
                const shapeProps = shape['p:spPr']['a:xfrm'];
                const shapeWidth = parseInt(shapeProps['a:ext']['$']['cx']);
                const shapeHeight = parseInt(shapeProps['a:ext']['$']['cy']);

                const paragraphs = textBody['a:p'];
                const shapeText = Array.isArray(paragraphs)
                    ? paragraphs.map(p => p['a:r']['a:t']).join('\n')
                    : paragraphs['a:r']['a:t'];

                // This is a rough approximation. A more accurate method would require
                // complex font metrics.
                const estimatedTextWidth = shapeText.length * 10000; // Assuming 10000 EMUs per char
                const estimatedTextHeight = (shapeText.split('\n').length) * 200000; // Assuming 200000 EMUs per line

                if (estimatedTextWidth > shapeWidth || estimatedTextHeight > shapeHeight) {
                    return false; // Content likely overflows
                }
            }
        }
        return true; // All content fits
    }

    async verifySlideBounds(slideNumber, slideWidth, slideHeight) {
        const slide = await this.getSlide(slideNumber);
        const shapes = slide['p:sld']['p:cSld']['p:spTree']['p:sp'];
        const shapeArray = Array.isArray(shapes) ? shapes : [shapes];

        const slideWidthEmu = slideWidth * 914400;
        const slideHeightEmu = slideHeight * 914400;

        for (const shape of shapeArray) {
            const xfrm = shape['p:spPr']['a:xfrm'];
            const x = parseInt(xfrm['a:off']['$']['x']);
            const y = parseInt(xfrm['a:off']['$']['y']);
            const w = parseInt(xfrm['a:ext']['$']['cx']);
            const h = parseInt(xfrm['a:ext']['$']['cy']);

            if (x < 0 || y < 0 || x + w > slideWidthEmu || y + h > slideHeightEmu) {
                return false; // Shape is out of bounds
            }
        }
        return true; // All shapes are within bounds
    }

    async calculateQualityScore(slideNumber, slideWidth, slideHeight) {
        let score = 10;
        let deductions = '';

        if (!(await this.verifyNoOverlaps(slideNumber))) {
            score -= 4;
            deductions += '[-4 points for overlapping elements] ';
        }
        if (!(await this.verifyContentFit(slideNumber))) {
            score -= 3;
            deductions += '[-3 points for content overflow] ';
        }
        if (!(await this.verifySlideBounds(slideNumber, slideWidth, slideHeight))) {
            score -= 3;
            deductions += '[-3 points for elements out of bounds]';
        }

        console.log(`  Quality Score: ${score}/10 ${deductions}`);
        return score;
    }
}

module.exports = { PptxVerifier };
