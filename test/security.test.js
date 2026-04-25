const { validatePath, isPathSafe } = require('../lib/security-utils');
const { HTML2PPTX } = require('../lib/html2pptx');
const path = require('path');
const assert = require('assert');
const fs = require('fs');

async function runSecurityTests() {
    console.log('🧪 Running Security Tests\n');

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

    // validatePath tests
    test('validatePath should accept valid path', () => {
        assert.strictEqual(validatePath('test.html'), true);
    });

    test('validatePath should reject null bytes', () => {
        assert.throws(() => validatePath('test.html\0'), /Path contains null bytes/);
    });

    // isPathSafe tests
    test('isPathSafe should allow any path if baseDir is null', () => {
        assert.strictEqual(isPathSafe('/etc/passwd', null), true);
    });

    test('isPathSafe should allow path within baseDir', () => {
        const baseDir = path.resolve('./examples');
        const filePath = path.resolve('./examples/check.html');
        assert.strictEqual(isPathSafe(filePath, baseDir), true);
    });

    test('isPathSafe should reject path outside baseDir', () => {
        const baseDir = path.resolve('./examples');
        const filePath = path.resolve('./package.json');
        assert.strictEqual(isPathSafe(filePath, baseDir), false);
    });

    test('isPathSafe should reject path traversal attempt', () => {
        const baseDir = path.resolve('./examples');
        const filePath = path.join(baseDir, '../package.json');
        assert.strictEqual(isPathSafe(filePath, baseDir), false);
    });

    // HTML2PPTX integration tests
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    await testAsync('HTML2PPTX should reject null bytes in input path', async () => {
        const converter = new HTML2PPTX();
        try {
            await converter.convert('test.html\0', 'output.pptx');
            assert.fail('Should have thrown an error');
        } catch (e) {
            assert.ok(e.message.includes('Path contains null bytes'));
        }
    });

    await testAsync('HTML2PPTX should respect baseDir and reject outside paths', async () => {
        const baseDir = path.join(__dirname, 'output');
        const converter = new HTML2PPTX({ baseDir });
        const inputPath = path.resolve(__dirname, '../examples/check.html');
        const outputPath = path.join(baseDir, 'out.pptx');

        try {
            await converter.convert(inputPath, outputPath);
            assert.fail('Should have thrown an error because input is outside baseDir');
        } catch (e) {
            assert.ok(e.message.includes('Access denied: input path is outside of base directory'));
        }
    });

    console.log('\n' + '='.repeat(50));
    console.log(`Security Test Results: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(50));

    if (failed > 0) {
        process.exit(1);
    }
}

if (require.main === module) {
    runSecurityTests().catch(error => {
        console.error('Security test suite error:', error);
        process.exit(1);
    });
}

module.exports = { runSecurityTests };
