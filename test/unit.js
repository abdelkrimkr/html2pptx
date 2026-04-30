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

// Extract the parseBorderWidth function body from the source file
const widthFunctionMatch = content.match(/parseBorderWidth\(borderWidth\) \{([\s\S]*?)\n    \}/);
if (!widthFunctionMatch) {
    console.error('Could not find parseBorderWidth function in source file');
    process.exit(1);
}
const widthFunctionBody = widthFunctionMatch[1];
const extractedParseBorderWidth = new Function('borderWidth', widthFunctionBody);

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

let totalFailed = testParseBorderStyle();

function testParseBorderWidth() {
    console.log('🧪 Testing parseBorderWidth (extracted from source)...');

    const testCases = [
        { input: null, expected: 1, desc: 'null input' },
        { input: undefined, expected: 1, desc: 'undefined input' },
        { input: '', expected: 1, desc: 'empty string' },
        { input: '1px', expected: 1, desc: '1px' },
        { input: '2', expected: 2, desc: '2' },
        { input: '1.5em', expected: 1.5, desc: '1.5em' },
        { input: 'thin', expected: 1, desc: 'non-numeric thin' },
        { input: 'thick', expected: 1, desc: 'non-numeric thick' },
        { input: 3, expected: 3, desc: 'number 3' },
        { input: 0, expected: 0, desc: 'number 0' },
        { input: '0px', expected: 0, desc: '0px' },
    ];

    let passed = 0;
    let failed = 0;

    testCases.forEach(tc => {
        try {
            const result = extractedParseBorderWidth(tc.input);
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

totalFailed += testParseBorderWidth();

if (totalFailed > 0) {
    process.exit(1);
} else {
    process.exit(0);
}
