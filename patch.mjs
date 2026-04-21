import os from 'os';
import * as nodeOs from 'node:os';

const DESKTOP = 'DESKTOP';
os.hostname = () => DESKTOP;
nodeOs.hostname = () => DESKTOP;

console.log('os and node:os patched.');
