import os from 'node:os';
import { createRequire } from 'node:module';

// os.hostname()이 'laptop'을 반환하도록 강제 수정
os.hostname = () => 'laptop';
console.log('os.hostname patched to: laptop');

const vercelPath = 'C:/Users/user/AppData/Roaming/npm/node_modules/vercel/dist/vc.js';

// Vercel CLI 실행
import(`file:///${vercelPath}`);
