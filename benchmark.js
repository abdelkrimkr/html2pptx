const { convertHTML2PPTX } = require('./lib/html2pptx');
const path = require('path');
const fs = require('fs');

async function runBenchmark() {
    if (!fs.existsSync(path.join(__dirname, 'test/output'))) {
        fs.mkdirSync(path.join(__dirname, 'test/output'), { recursive: true });
    }

    const start = process.hrtime.bigint();
    for (let i = 0; i < 50; i++) {
        await convertHTML2PPTX(path.join(__dirname, 'examples/check.html'), path.join(__dirname, `test/output/bench_${i}.pptx`));
    }
    const end = process.hrtime.bigint();
    const durationMs = Number(end - start) / 1000000;

    console.log(`Benchmark completed in ${durationMs.toFixed(2)} ms`);
    console.log(`Average time per conversion: ${(durationMs / 50).toFixed(2)} ms`);
}

runBenchmark().catch(console.error);
