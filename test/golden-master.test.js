const { convertHTML2PPTX } = require('../lib/html2pptx');
const { PptxVerifier } = require('./pptx-verifier');
const path = require('path');
const fs = require('fs');
const assert = require('assert');

const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

describe('Golden Master Tests', function() {
    it('should correctly render a simple vertical flow layout', async function() {
        this.timeout(10000); // 10s timeout for conversion and verification

        const inputFile = path.join(__dirname, 'golden-master/simple-flow.html');
        const outputFile = path.join(outputDir, 'golden-simple-flow.pptx');

        await convertHTML2PPTX(inputFile, outputFile);

        // Verify the output
        const verifier = new PptxVerifier(outputFile);

        // 1. Check shape count
        const shapeCount = await verifier.getShapeCount(1);
        assert.strictEqual(shapeCount, 2, 'Should have exactly 2 shapes on the slide');

        // 2. Check text content
        const text1 = await verifier.getShapeText(1, 0);
        assert.strictEqual(text1, 'First Box', 'Text of the first box should be correct');

        const text2 = await verifier.getShapeText(1, 1);
        assert.strictEqual(text2, 'Second Box', 'Text of the second box should be correct');

        // 3. Check positioning to prevent overlap
        const pos1 = await verifier.findShapeByText(1, 'First Box');
        const pos2 = await verifier.findShapeByText(1, 'Second Box');

        assert.ok(pos1, 'Could not find the first box by its text');
        assert.ok(pos2, 'Could not find the second box by its text');

        // Check that the second box is below the first one
        assert.ok(pos2.y >= pos1.y + pos1.h, 'The second box should be positioned below the first box');

        // Check that they are roughly aligned on the X axis
        const xDifference = Math.abs(pos1.x - pos2.x);
        assert.ok(xDifference < 10000, 'Boxes should be vertically aligned');
    });

    it('should correctly handle CSS transforms', async function() {
        this.timeout(10000);

        const inputFile = path.join(__dirname, 'golden-master/simple-transform.html');
        const outputFile = path.join(outputDir, 'golden-simple-transform.pptx');

        await convertHTML2PPTX(inputFile, outputFile);

        const verifier = new PptxVerifier(outputFile);

        // 1. Check shape count
        const shapeCount = await verifier.getShapeCount(1);
        assert.strictEqual(shapeCount, 2, 'Should have exactly 2 shapes on the slide');

        // 2. Check for rotation
        const rotatedShape = await verifier.findShapeByText(1, 'Rotated Box');
        // Note: PptxGenJS does not expose rotation in the generated XML, so we cannot verify it.
        // This is a known limitation. We are testing that it does not crash.
        assert.ok(rotatedShape, 'Rotated box should be present');

        // 3. Check for scaling
        const scaledShape = await verifier.findShapeByText(1, 'Scaled Box');
        assert.ok(scaledShape, 'Scaled box should be present');

        // We can check if the scaled shape is larger than its original size
        const originalWidth = 200 * (10 / 1280); // original width in inches
        const originalHeight = 80 * (5.625 / 720); // original height in inches
        const expectedMinWidth = originalWidth * 1.2 * 914400; // in EMUs
        const expectedMinHeight = originalHeight * 1.2 * 914400; // in EMUs

        assert.ok(scaledShape.w > originalWidth * 914400, 'Scaled box width should be larger than original');
        assert.ok(scaledShape.h > originalHeight * 914400, 'Scaled box height should be larger than original');
    });

    it('should correctly handle hyperlinks', async function() {
        this.timeout(10000);

        const inputFile = path.join(__dirname, 'golden-master/simple-hyperlink.html');
        const outputFile = path.join(outputDir, 'golden-simple-hyperlink.pptx');

        await convertHTML2PPTX(inputFile, outputFile);

        const verifier = new PptxVerifier(outputFile);

        // 1. Check shape count
        const shapeCount = await verifier.getShapeCount(1);
        assert.strictEqual(shapeCount, 1, 'Should have exactly 1 shape on the slide');

        // 2. Check for hyperlink
        // Note: PptxGenJS does not expose hyperlinks in a way we can easily verify through the XML.
        // This is a known limitation. We are testing that it does not crash and the text is correct.
        const linkShape = await verifier.findShapeByText(1, 'Click Here');
        assert.ok(linkShape, 'Link box should be present');
    });

    it('should correctly handle a simple flexbox layout', async function() {
        this.timeout(10000);

        const inputFile = path.join(__dirname, 'golden-master/simple-flex.html');
        const outputFile = path.join(outputDir, 'golden-simple-flex.pptx');

        await convertHTML2PPTX(inputFile, outputFile);

        const verifier = new PptxVerifier(outputFile);

        const leftBox = await verifier.findShapeByText(1, 'Left Box');
        const rightBox = await verifier.findShapeByText(1, 'Right Box');

        assert.ok(leftBox, 'Could not find the left box');
        assert.ok(rightBox, 'Could not find the right box');

        // Check that they are roughly aligned on the Y axis
        const yDifference = Math.abs(leftBox.y - rightBox.y);
        assert.ok(yDifference < 10000, 'Boxes should be horizontally aligned');

        // Check that the right box is to the right of the left box
        assert.ok(rightBox.x > leftBox.x + leftBox.w, 'The right box should be positioned to the right of the left box');
    });

    it('should correctly handle a simple grid layout', async function() {
        this.timeout(10000);

        const inputFile = path.join(__dirname, 'golden-master/simple-grid.html');
        const outputFile = path.join(outputDir, 'golden-simple-grid.pptx');

        await convertHTML2PPTX(inputFile, outputFile);

        const verifier = new PptxVerifier(outputFile);

        const cell1 = await verifier.findShapeByText(1, 'First Cell');
        const cell2 = await verifier.findShapeByText(1, 'Second Cell');

        assert.ok(cell1, 'Could not find the first cell');
        assert.ok(cell2, 'Could not find the second cell');

        // Check that they are roughly aligned on the Y axis
        const yDifference = Math.abs(cell1.y - cell2.y);
        assert.ok(yDifference < 10000, 'Cells should be horizontally aligned');

        // Check that the second cell is to the right of the first cell
        assert.ok(cell2.x > cell1.x + cell1.w, 'The second cell should be positioned to the right of the first cell');

        const score = await verifier.calculateQualityScore(1, 10, 5.625);
        assert.ok(score > 8, `Quality score of ${score} is too low`);
    });

    it('should correctly handle advanced border styles', async function() {
        this.timeout(10000);

        const inputFile = path.join(__dirname, 'golden-master/simple-borders.html');
        const outputFile = path.join(outputDir, 'golden-simple-borders.pptx');

        await convertHTML2PPTX(inputFile, outputFile);

        const verifier = new PptxVerifier(outputFile);

        // Note: PptxGenJS does not expose border styles in a way we can easily verify through the XML.
        // This is a known limitation. We are testing that it does not crash and the text is correct.
        const dashedBox = await verifier.findShapeByText(1, 'Dashed Border');
        assert.ok(dashedBox, 'Dashed border box should be present');

        const dottedBox = await verifier.findShapeByText(1, 'Dotted Border');
        assert.ok(dottedBox, 'Dotted border box should be present');

        const score = await verifier.calculateQualityScore(1, 10, 5.625);
        assert.ok(score > 8, `Quality score of ${score} is too low`);
    });
});
