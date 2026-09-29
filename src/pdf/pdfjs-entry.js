// Pacote separado do pdf.js, carregado só quando um PDF é aberto.
// O worker roda na própria página: o protocolo file:// não permite Workers.
import * as pdfjsLib from 'pdfjs-dist/build/pdf.mjs';
import * as worker from 'pdfjs-dist/build/pdf.worker.mjs';

globalThis.pdfjsWorker = worker;
window.RedatorPdf = pdfjsLib;
