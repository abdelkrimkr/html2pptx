const cheerio = require('cheerio');
const fs = require('fs');
const html = fs.readFileSync('/home/ubuntu/Uploads/1.html', 'utf8');
const $ = cheerio.load(html);

const textCenter = $('.text-center');
console.log('text-center element found:', textCenter.length);
console.log('text-center classes:', textCenter.attr('class'));

// Check Tailwind classes
const container = $('.slide-container');
console.log('\nslide-container classes:', container.attr('class'));

// Since Tailwind is external CSS, we need to handle common Tailwind classes
console.log('\nKnown Tailwind utility classes:');
console.log('- text-center -> text-align: center');
console.log('- flex -> display: flex');
console.log('- flex-col -> flex-direction: column');
console.log('- items-center -> align-items: center');
console.log('- justify-center -> justify-content: center');
