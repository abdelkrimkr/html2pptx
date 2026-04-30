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

// Extract the parseFontFamily function body from the source file
const fontFamilyMatch = content.match(/parseFontFamily\(fontFamily\) \{([\s\S]*?)\n    \}/);
if (!fontFamilyMatch) {
    console.error('Could not find parseFontFamily function in source file');
    process.exit(1);
}

const fontFamilyBody = fontFamilyMatch[1];
const extractedParseFontFamily = new Function('fontFamily', fontFamilyBody);

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

function testParseFontFamily() {
    console.log('🧪 Testing parseFontFamily (extracted from source)...');

    const testCases = [
        { input: null, expected: 'Arial', desc: 'null input' },
        { input: undefined, expected: 'Arial', desc: 'undefined input' },
        { input: '', expected: 'Arial', desc: 'empty string' },
        { input: 'Roboto', expected: 'Arial', desc: 'mapped font (Roboto)' },
        { input: 'Montserrat', expected: 'Arial', desc: 'mapped font (Montserrat)' },
        { input: 'serif', expected: 'Times New Roman', desc: 'mapped font (serif)' },
        { input: 'monospace', expected: 'Courier New', desc: 'mapped font (monospace)' },
        { input: 'Comic Sans', expected: 'Comic Sans', desc: 'unknown font' },
        { input: '"Open Sans", sans-serif', expected: 'Open Sans', desc: 'quoted font with fallback' },
        { input: "'Helvetica Neue', Helvetica, Arial, sans-serif", expected: 'Helvetica Neue', desc: 'single quotes multiple fonts' },
        { input: '  Georgia  ', expected: 'Georgia', desc: 'leading/trailing spaces' }
    ];

    let passed = 0;
    let failed = 0;

    testCases.forEach(tc => {
        try {
            const result = extractedParseFontFamily(tc.input);
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
totalFailed += testParseFontFamily();

if (totalFailed > 0) {
    process.exit(1);
} else {
    process.exit(0);
}
