const fs = require('fs');
const path = require('path');

/**
 * This test file uses code extraction because the project's dependencies
 * are not correctly installed in the environment, preventing a standard require().
 */

const filePath = path.join(__dirname, '../lib/html2pptx.js');
let content = fs.readFileSync(filePath, 'utf8');

// Extract the parseBorderStyle function body from the source file
const functionMatch = content.match(/parseBorderStyle\(borderStyle\) \{([\s\S]*?)\n    \}/);
if (!functionMatch) {
    console.error('Could not find parseBorderStyle function in source file');
    process.exit(1);
}

const functionBody = functionMatch[1];
const extractedParseBorderStyle = new Function('borderStyle', functionBody);


// Extract the extractGradientColors function body from the source file
const extractGradientMatch = content.match(/extractGradientColors\(gradient\) \{([\s\S]*?)\n    \}/);
if (!extractGradientMatch) {
    console.error('Could not find extractGradientColors function in source file');
    process.exit(1);
}

const extractGradientBody = extractGradientMatch[1];
const extractedGradientColors = new Function('gradient', extractGradientBody);

function testExtractGradientColors() {
    console.log('\n🧪 Testing extractGradientColors (extracted from source)...');

    const testCases = [
        { input: 'linear-gradient(#ff0000, #00ff00)', expected: ['#ff0000', '#00ff00'], desc: 'Standard hex colors' },
        { input: 'radial-gradient(circle, rgba(255,0,0,0.5), rgba(0,255,0,0.8))', expected: ['rgba(255,0,0,0.5)', 'rgba(0,255,0,0.8)'], desc: 'Standard rgba colors' },
        { input: 'linear-gradient(rgb(255, 0, 0), rgb(0, 255, 0))', expected: ['rgb(255, 0, 0)', 'rgb(0, 255, 0)'], desc: 'Standard rgb colors' },
        { input: 'linear-gradient(#f00, #0f0)', expected: ['#f00', '#0f0'], desc: 'Short hex colors' },
        { input: 'linear-gradient(#ff0000, rgba(0, 255, 0, 0.5))', expected: ['#ff0000', 'rgba(0, 255, 0, 0.5)'], desc: 'Mixed hex and rgba' },
        { input: 'linear-gradient(to right, red, blue)', expected: [], desc: 'Named colors (not supported by regex currently)' },
        { input: '', expected: [], desc: 'Empty string' },
        { input: null, expectedError: true, desc: 'Null input' },
        { input: undefined, expectedError: true, desc: 'Undefined input' }
    ];

    let passed = 0;
    let failed = 0;

    testCases.forEach(tc => {
        try {
            const result = extractedGradientColors(tc.input);
            if (tc.expectedError) {
                console.log(`  ❌ ${tc.desc}: Expected error but got result`);
                failed++;
            } else if (JSON.stringify(result) !== JSON.stringify(tc.expected)) {
                throw new Error(`expected ${JSON.stringify(tc.expected)}, got ${JSON.stringify(result)}`);
            } else {
                console.log(`  ✅ ${tc.desc}`);
                passed++;
            }
        } catch (err) {
            if (tc.expectedError) {
                console.log(`  ✅ ${tc.desc} (caught expected error)`);
                passed++;
            } else {
                console.log(`  ❌ ${tc.desc}: ${err.message}`);
                failed++;
            }
        }
    });

    console.log(`\nResults: ${passed} passed, ${failed} failed\n`);
    return failed;
}

function testParseBorderStyle() {
    console.log('🧪 Testing parseBorderStyle (extracted from source)...');

    const testCases = [
        { input: null, expected: 'solid', desc: 'null input' },
        { input: undefined, expected: 'solid', desc: 'undefined input' },
        { input: '', expected: 'solid', desc: 'empty string' },
        { input: 'solid', expected: 'solid', desc: 'solid style' },
        { input: 'dashed', expected: 'dash', desc: 'dashed style' },
        { input: 'dotted', expected: 'dot', desc: 'dotted style' },
        { input: 'double', expected: 'dblPt', desc: 'double style' },
        { input: 'dashdot', expected: 'dashDot', desc: 'dashdot style' },
        { input: 'longdash', expected: 'lgDash', desc: 'longdash style' },
        { input: 'longdashdot', expected: 'lgDashDot', desc: 'longdashdot style' },
        { input: 'longdashdotdot', expected: 'lgDashDotDot', desc: 'longdashdotdot style' },
        { input: '1px solid black', expected: 'solid', desc: 'complex solid' },
        { input: '2px dashed red', expected: 'dash', desc: 'complex dashed' },
        { input: 'none', expected: 'solid', desc: 'none (fallback to solid)' },
        { input: 'unknown', expected: 'solid', desc: 'unknown style' }
    ];

    let passed = 0;
    let failed = 0;

    testCases.forEach(tc => {
        try {
            const result = extractedParseBorderStyle(tc.input);
            if (result !== tc.expected) {
                throw new Error(`expected ${tc.expected}, got ${result}`);
            }
            console.log(`  ✅ ${tc.desc}`);
            passed++;
        } catch (err) {
            console.log(`  ❌ ${tc.desc}: ${err.message}`);
            failed++;
        }
    });

    console.log(`\nResults: ${passed} passed, ${failed} failed\n`);
    return failed;
}

const failedParseBorderStyle = testParseBorderStyle();
const failedExtractGradient = testExtractGradientColors();
const totalFailed = failedParseBorderStyle + failedExtractGradient;

if (totalFailed > 0) {
    process.exit(1);
} else {
    process.exit(0);
}
