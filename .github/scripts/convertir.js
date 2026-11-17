'use strict';
// Convierte los archivos de ventas .XLS/.XLSX de una carpeta a CSV (separador ";", Windows-1252),
// que es el formato que la app de COMACA 360 lee por partes y con poca memoria.
//
// Uso:  node convertir.js <carpeta_entrada> <carpeta_salida>
//
// El CSV se genera para que la app obtenga EXACTAMENTE las mismas filas que leyendo el Excel
// (comprobado con el archivo real: mismas filas, mismos campos y misma fecha máxima).
// Si algo no cuadra (sin filas, sin las columnas esperadas), el programa falla y NO se publica nada.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Windows-1252: los códigos 0x80-0x9F no coinciden con Latin-1.
const CP1252 = { '€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84, '…': 0x85, '†': 0x86, '‡': 0x87, 'ˆ': 0x88, '‰': 0x89, 'Š': 0x8A, '‹': 0x8B, 'Œ': 0x8C, 'Ž': 0x8E, '‘': 0x91, '’': 0x92, '“': 0x93, '”': 0x94, '•': 0x95, '–': 0x96, '—': 0x97, '˜': 0x98, '™': 0x99, 'š': 0x9A, '›': 0x9B, 'œ': 0x9C, 'ž': 0x9E, 'Ÿ': 0x9F };
function aCp1252(texto) {
  const out = Buffer.alloc(texto.length * 2);
  let n = 0;
  for (const ch of texto) {
    const c = ch.codePointAt(0);
    if (c < 0x80 || (c >= 0xA0 && c <= 0xFF)) out[n++] = c;
    else if (CP1252[ch] !== undefined) out[n++] = CP1252[ch];
    else out[n++] = 0x3F; // '?'
  }
  return out.subarray(0, n);
}

function convertirArchivo(rutaEntrada) {
  const wb = XLSX.read(fs.readFileSync(rutaEntrada), { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const filas = XLSX.utils.sheet_to_json(ws); // lo mismo que lee la app
  if (filas.length === 0) throw new Error('el archivo no tiene filas');
  const cabeceras = Object.keys(filas[0]).join('|');
  if (!/fecha/i.test(cabeceras) || !/monto|total/i.test(cabeceras)) {
    throw new Error('no se encuentran las columnas esperadas (Fecha y Monto/Total)');
  }
  // rawNumbers: números sin formato de celda (igual que la lectura del Excel); blankrows:false: sin filas vacías (igual que sheet_to_json)
  const csv = XLSX.utils.sheet_to_csv(ws, { FS: ';', rawNumbers: true, blankrows: false });
  const lineas = csv.split('\n').filter(l => l.trim() !== '').length;
  if (lineas < filas.length + 1) throw new Error(`el CSV tiene menos líneas (${lineas}) que filas el Excel (${filas.length + 1})`);
  return { datos: aCp1252(csv), filas: filas.length };
}

if (require.main === module) {
  const [,, dirEntrada, dirSalida] = process.argv;
  if (!dirEntrada || !dirSalida) { console.error('Uso: node convertir.js <carpeta_entrada> <carpeta_salida>'); process.exit(2); }
  const archivos = fs.readdirSync(dirEntrada).filter(f => /^VENTAS.*\.(xls|xlsx)$/i.test(f));
  if (archivos.length === 0) { console.error('No hay archivos VENTAS*.XLS en la carpeta'); process.exit(1); }
  fs.mkdirSync(dirSalida, { recursive: true });
  let fallos = 0;
  for (const f of archivos) {
    try {
      const { datos, filas } = convertirArchivo(path.join(dirEntrada, f));
      const salida = path.join(dirSalida, path.parse(f).name + '.csv');
      fs.writeFileSync(salida, datos);
      console.log(`OK    ${f} -> ${path.basename(salida)}  (${filas} filas, ${(datos.length / 1e6).toFixed(1)} MB)`);
    } catch (e) {
      fallos++;
      console.error(`ERROR ${f}: ${e.message}`);
    }
  }
  process.exit(fallos ? 1 : 0);
}

module.exports = { convertirArchivo, aCp1252 };
