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


// Extract the parseColor function body from the source file
const colorMatch = content.match(/parseColor\(color\) \{([\s\S]*?)\n    \}/);
if (!colorMatch) {
    console.error('Could not find parseColor function in source file');
    process.exit(1);
}

const parseColorBody = colorMatch[1];
const extractedParseColor = new Function('color', parseColorBody);

function testParseColor() {
    console.log('🧪 Testing parseColor (extracted from source)...');

    const testCases = [
        { input: null, expected: '000000', desc: 'null input' },
        { input: undefined, expected: '000000', desc: 'undefined input' },
        { input: '', expected: '000000', desc: 'empty string' },
        { input: '#FFFFFF', expected: 'FFFFFF', desc: 'hex with hash' },
        { input: 'FFFFFF', expected: 'FFFFFF', desc: 'hex without hash' },
        { input: '#fff', expected: 'FFFFFF', desc: '3-digit hex with hash' },
        { input: '000', expected: '000000', desc: '3-digit hex without hash' },
        { input: '#12345678', expected: '123456', desc: '8-digit hex with hash' },
        { input: '12345678', expected: '123456', desc: '8-digit hex without hash' },
        { input: 'rgb(255, 0, 0)', expected: 'FF0000', desc: 'rgb format' },
        { input: 'rgba(0, 255, 0, 0.5)', expected: '00FF00', desc: 'rgba format' },
        { input: 'white', expected: 'FFFFFF', desc: 'named color (white)' },
        { input: 'BLACK', expected: '000000', desc: 'named color uppercase' },
        { input: 'ReD', expected: 'FF0000', desc: 'named color mixed case' },
        { input: 'invalid', expected: '000000', desc: 'invalid string' }
    ];

    let passed = 0;
    let failed = 0;

    testCases.forEach(tc => {
        try {
            const result = extractedParseColor(tc.input);
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
totalFailed += testParseColor();

if (totalFailed > 0) {
    process.exit(1);
} else {
    process.exit(0);
}
