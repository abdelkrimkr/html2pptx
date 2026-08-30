const { HTML2PPTX } = require('../lib/html2pptx');
const cheerio = require('cheerio');

async function runCSSTests() {
    console.log('🧪 Running CSS Tests\n');

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

    test('extractCSS should handle malformed CSS without crashing and log a warning', () => {
        const converter = new HTML2PPTX();
        const $ = cheerio.load('<style> { malformed css !@# </style>');

        let warningLogged = false;
        const originalWarn = console.warn;

        try {
            console.warn = (msg, ...args) => {
                if (msg.includes('CSS parsing warning:')) {
                    warningLogged = true;
                }
            };

            converter.extractCSS($);

            if (!warningLogged) {
                throw new Error('Expected CSS parsing warning to be logged');
            }
        } finally {
            console.warn = originalWarn;
        }
    });

    console.log('\n' + '='.repeat(50));
    console.log(`CSS Test Results: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(50));

    if (failed > 0) {
        throw new Error('CSS tests failed');
    }
}

if (require.main === module) {
    runCSSTests().catch(error => {
        console.error('CSS test suite error:', error);
        process.exit(1);
    });
}

module.exports = { runCSSTests };
