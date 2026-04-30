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

// Extract the parseFontFamily function body from the source file
const functionMatchFont = content.match(/parseFontFamily\(fontFamily\) \{([\s\S]*?)\n    \}/);
if (!functionMatchFont) {
    console.error('Could not find parseFontFamily function in source file');
    process.exit(1);
}

const functionBodyFont = functionMatchFont[1];
const extractedParseFontFamily = new Function('fontFamily', functionBodyFont);

function testParseFontFamily() {
    console.log('\n🧪 Testing parseFontFamily (extracted from source)...');

    const testCases = [
        { input: null, expected: 'Arial', desc: 'null input' },
        { input: undefined, expected: 'Arial', desc: 'undefined input' },
        { input: '', expected: 'Arial', desc: 'empty string' },
        { input: 'Roboto', expected: 'Arial', desc: 'Roboto mapped to Arial' },
        { input: 'Montserrat', expected: 'Arial', desc: 'Montserrat mapped to Arial' },
        { input: 'Helvetica', expected: 'Arial', desc: 'Helvetica mapped to Arial' },
        { input: 'sans-serif', expected: 'Arial', desc: 'sans-serif mapped to Arial' },
        { input: 'serif', expected: 'Times New Roman', desc: 'serif mapped to Times New Roman' },
        { input: 'monospace', expected: 'Courier New', desc: 'monospace mapped to Courier New' },
        { input: 'Verdana', expected: 'Verdana', desc: 'unmapped font returns itself' },
        { input: '"Times New Roman"', expected: 'Times New Roman', desc: 'font with double quotes' },
        { input: "'Courier New'", expected: 'Courier New', desc: 'font with single quotes' },
        { input: '"Roboto", sans-serif', expected: 'Arial', desc: 'comma-separated list, first font quoted and mapped' },
        { input: 'Georgia, serif', expected: 'Georgia', desc: 'comma-separated list, first font unmapped' }
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

const totalFailed = testParseBorderStyle() + testParseFontFamily();

if (totalFailed > 0) {
    process.exit(1);
} else {
    process.exit(0);
}
