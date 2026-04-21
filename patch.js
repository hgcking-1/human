const os = require('os');
const originalHostname = os.hostname;
os.hostname = () => 'DESKTOP';
console.log('os.hostname patched to: DESKTOP');
