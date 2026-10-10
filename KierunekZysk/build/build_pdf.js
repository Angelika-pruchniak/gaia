// Generuje PDF-y z plików HTML w ../src (Chromium przez Playwright)
const { chromium } = require('playwright');
const path = require('path');
const files = [
  ['01_Przewodnik.html', '1_Kierunek_Zysk_Przewodnik.pdf'],
  ['03_Strategia_na_1_stronie.html', '3_Strategia_na_1_stronie.pdf'],
  ['04_Planer_90_dni.html', '4_Planer_90_dni.pdf'],
  ['05_Skrypty_sprzedazowe.html', '5_Skrypty_i_szablony_sprzedazowe.pdf'],
  ['06_Audyt_i_bank_testow.html', '6_Audyt_i_bank_40_testow.pdf'],
];
(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROMIUM || undefined });
  const page = await browser.newPage();
  for (const [src, out] of files) {
    await page.goto('file://' + path.resolve(__dirname, '../src', src), { waitUntil: 'networkidle' });
    await page.evaluate(() => document.fonts.ready);
    await page.pdf({ path: path.resolve(__dirname, '../PDF', out), preferCSSPageSize: true, printBackground: true });
    console.log('OK', out);
  }
  await browser.close();
})();
