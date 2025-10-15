const fs = require('fs');
const cheerio = require('cheerio');

const html = fs.readFileSync('/home/ubuntu/Uploads/1.html', 'utf8');
const $ = cheerio.load(html);

console.log('=== TEXT ELEMENTS IN 1.html ===\n');

$('.text-center').children().each((i, elem) => {
    const $elem = $(elem);
    const tagName = $elem.prop('tagName').toLowerCase();
    const text = $elem.text().trim();
    console.log(`${i+1}. <${tagName}> "${text}"`);
});

console.log('\n=== PARENT STRUCTURE ===\n');
const parent = $('.text-center').parent();
console.log('Parent classes:', parent.attr('class'));
console.log('Text-center classes:', $('.text-center').attr('class'));
