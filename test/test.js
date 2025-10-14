const { convertHTML2PPTX } = require('../lib/html2pptx');
const path = require('path');
const fs = require('fs');

async function runTests() {
    console.log('🧪 Running HTML2PPTX Tests\n');
    
    // Create output directory
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }
    
    const tests = [
        {
            name: 'Test 1: Simple Text Boxes (5 Text Boxes 16_9.html)',
            input: path.join(__dirname, '../examples/5 Text Boxes 16_9.html'),
            output: path.join(outputDir, 'test1-textboxes.pptx')
        },
        {
            name: 'Test 2: Complex Layout (check.html)',
            input: path.join(__dirname, '../examples/check.html'),
            output: path.join(outputDir, 'test2-complex.pptx')
        }
    ];
    
    let passed = 0;
    let failed = 0;
    
    for (const test of tests) {
        console.log(`Running: ${test.name}`);
        console.log(`  Input:  ${test.input}`);
        console.log(`  Output: ${test.output}`);
        
        try {
            const startTime = Date.now();
            await convertHTML2PPTX(test.input, test.output);
            const duration = Date.now() - startTime;
            
            // Check if file was created
            if (fs.existsSync(test.output)) {
                const stats = fs.statSync(test.output);
                console.log(`  ✅ PASSED (${duration}ms, ${Math.round(stats.size / 1024)}KB)\n`);
                passed++;
            } else {
                console.log(`  ❌ FAILED: Output file not created\n`);
                failed++;
            }
        } catch (error) {
            console.log(`  ❌ FAILED: ${error.message}\n`);
            failed++;
        }
    }
    
    console.log('='.repeat(50));
    console.log(`Test Results: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(50));
    
    if (failed === 0) {
        console.log('\n✅ All tests passed!');
        console.log(`\nOutput files saved to: ${outputDir}`);
    } else {
        console.log('\n❌ Some tests failed.');
        process.exit(1);
    }
}

runTests().catch(error => {
    console.error('Test suite error:', error);
    process.exit(1);
});
