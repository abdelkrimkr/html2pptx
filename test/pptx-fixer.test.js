const assert = require('assert');
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
const { PPTXFixer } = require('../lib/pptx-fixer');

async function runFixerTests() {
    console.log('🧪 Running PPTX Fixer Tests\n');

    let passed = 0;
    let failed = 0;

    function test(name, fn) {
        try {
            fn();
            console.log(`✅ PASSED: ${name}`);
            passed++;
        } catch (e) {
            console.log(`❌ FAILED: ${name}`);
            console.log(`   Error: ${e.message}`);
            failed++;
        }
    }

    async function testAsync(name, fn) {
        try {
            await fn();
            console.log(`✅ PASSED: ${name}`);
            passed++;
        } catch (e) {
            console.log(`❌ FAILED: ${name}`);
            console.log(`   Error: ${e.message}`);
            failed++;
        }
    }

    const fixer = new PPTXFixer();

    // Fix 1
    test('fixXML should fix empty name attributes in p:cNvPr', () => {
        fixer.fixCount = 0;
        const result1 = fixer.fixXML('<p:cNvPr id="5" name=""/>', 'test.xml');
        assert.strictEqual(result1, '<p:cNvPr id="5" name="Shape 5"/>');

        const result2 = fixer.fixXML('<p:cNvPr id="12" name="">hello</p:cNvPr>', 'test.xml');
        assert.strictEqual(result2, '<p:cNvPr id="12" name="Shape 12">hello</p:cNvPr>');
    });

    // Fix 2
    test('fixXML should remove empty a:ln elements', () => {
        fixer.fixCount = 0;
        const result1 = fixer.fixXML('<a:ln></a:ln>', 'test.xml');
        assert.strictEqual(result1, '');

        const result2 = fixer.fixXML('<a:ln/>', 'test.xml');
        assert.strictEqual(result2, '');
    });

    // Fix 3
    test('fixXML should fix zero dimensions in a:ext', () => {
        fixer.fixCount = 0;
        // Note: cy="1" and cy="500" are updated to cy="10000" by Fix 5 sequentially.
        const result1 = fixer.fixXML('<a:ext cx="0" cy="0"/>', 'test.xml');
        assert.strictEqual(result1, '<a:ext cx="1" cy="10000"/>');

        const result2 = fixer.fixXML('<a:ext cx="0" cy="500"/>', 'test.xml');
        assert.strictEqual(result2, '<a:ext cx="1" cy="10000"/>');

        const result3 = fixer.fixXML('<a:ext cx="500" cy="0"/>', 'test.xml');
        assert.strictEqual(result3, '<a:ext cx="500" cy="10000"/>');
    });

    // Fix 4
    test('fixXML should fix conflicting autofit settings', () => {
        fixer.fixCount = 0;
        const input = '<a:bodyPr><a:normAutofit/><a:spAutoFit/></a:bodyPr>';
        const result = fixer.fixXML(input, 'test.xml');
        assert.strictEqual(result, '<a:bodyPr><a:normAutofit/></a:bodyPr>');
    });

    // Fix 5
    test('fixXML should fix very small cy dimensions', () => {
        fixer.fixCount = 0;
        const result1 = fixer.fixXML('cy="500"', 'test.xml');
        assert.strictEqual(result1, 'cy="10000"');

        const result2 = fixer.fixXML('cy="0"', 'test.xml'); // Not modified by this rule
        assert.strictEqual(result2, 'cy="0"');

        const result3 = fixer.fixXML('cy="20000"', 'test.xml');
        assert.strictEqual(result3, 'cy="20000"');
    });

    // Fix 6
    test('fixXML should fix empty or invalid charset attributes', () => {
        fixer.fixCount = 0;
        const result1 = fixer.fixXML('charset="-122"', 'test.xml');
        assert.strictEqual(result1, 'charset="0"');

        const result2 = fixer.fixXML('charset="-120"', 'test.xml');
        assert.strictEqual(result2, 'charset="0"');

        const result3 = fixer.fixXML('charset="1"', 'test.xml');
        assert.strictEqual(result3, 'charset="1"');
    });

    // Integration Test
    await testAsync('fixPPTX should fix corrupted files inside zip', async () => {
        const outputDir = path.join(__dirname, 'output');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const testPptxPath = path.join(outputDir, 'test-corrupt.pptx');
        const zip = new AdmZip();

        // Add a mock XML file with some corruptions
        const corruptXML = '<p:cNvPr id="7" name=""/><a:ln/><a:ext cx="0" cy="0"/><a:bodyPr><a:normAutofit/><a:spAutoFit/></a:bodyPr>cy="500"charset="-122"';
        zip.addFile('ppt/slides/slide1.xml', Buffer.from(corruptXML, 'utf8'));
        zip.writeZip(testPptxPath);

        const result = await fixer.fixPPTX(testPptxPath);
        assert.strictEqual(result.success, true);
        assert.strictEqual(result.fixCount > 0, true);

        // Verify it was fixed
        const fixedZip = new AdmZip(testPptxPath);
        const fixedContent = fixedZip.getEntry('ppt/slides/slide1.xml').getData().toString('utf8');

        assert.ok(!fixedContent.includes('name=""'));
        assert.ok(!fixedContent.includes('<a:ln/>'));
        assert.ok(!fixedContent.includes('cx="0" cy="0"'));
        assert.ok(!fixedContent.includes('<a:spAutoFit/>'));
        assert.ok(!fixedContent.includes('cy="500"'));
        assert.ok(!fixedContent.includes('charset="-122"'));
    });

    console.log('\n' + '='.repeat(50));
    console.log(`PPTX Fixer Test Results: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(50));

    if (failed > 0) {
        process.exit(1);
    }
}

if (require.main === module) {
    runFixerTests().catch(error => {
        console.error('PPTX Fixer test suite error:', error);
        process.exit(1);
    });
}

module.exports = { runFixerTests };
