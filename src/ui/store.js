// Penyimpanan lokal per komputer (localStorage). Semua akses dibungkus try/catch.

const PREFIX = 'master.';

export function load(key, fallback) {
  try {
    const raw = window.localStorage.getItem(PREFIX + key);
    return raw === null ? fallback : JSON.parse(raw);
  } catch (e) {
    return fallback;
  }
}

export function save(key, value) {
  try {
    window.localStorage.setItem(PREFIX + key, JSON.stringify(value));
    return true;
  } catch (e) {
    return false;
  }
}

export const DEFAULT_SETTINGS = {
  apiKey: '',
  model: 'claude-opus-5-5',
  effort: 'medium',
  mode: 'auto', // auto | ai | offline
  codePolicy: 'ask', // ask | always | never
};

export function getSettings() {
  return Object.assign({}, DEFAULT_SETTINGS, load('settings', {}));
}

export function getProfile() {
  return Object.assign({ nama: '', jabatan: '', instansi: '', alamat: '', kota: '', telepon: '', email: '' }, load('profile', {}));
}
