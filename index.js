const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
let files = [];
let lastResults = [];

const camposExtras = ["PS", "Cupones", "Pagos", "Cobranzas", "TC", "CPD", "Opera a Crédito", "Acuerdo", "Cuenta MS"];

dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropzone.style.backgroundColor = '#f0f0f0';
});

dropzone.addEventListener('dragleave', () => {
  dropzone.style.backgroundColor = '';
});

dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.style.backgroundColor = '';
  files = Array.from(e.dataTransfer.files);
  processFiles(files);
});

fileInput.addEventListener('change', (e) => {
  files = Array.from(e.target.files);
  processFiles(files);
});

function readExcel(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const workbook = XLSX.read(e.target.result, { type: 'binary' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet);
      resolve(data);
    };
    reader.readAsBinaryString(file);
  });
}

async function processFiles(files) {
  if (files.length < 2) {
    alert('Por favor sube dos archivos.');
    return;
  }

  const [lastCrossFile, newCrossFile] = files;
  const lastCrossRaw = await readExcel(lastCrossFile);
  const newCrossRaw = await readExcel(newCrossFile);

  const cleanCrossData = (data, label) => data.map(row => {
    const base = {
      Codigo_Cliente: row['Codigo_Cliente'],
      [`Nombre_${label}`]: row['Nombre'],
      [`Segmento_${label}`]: row['Segmento'],
      [`ICS_BE_${label}`]: row['ICS - BE'] || 0
    };
    camposExtras.forEach(campo => {
      base[`${campo}_${label}`] = row[campo] ?? 0;
    });
    return base;
  });

  const lastCrossData = cleanCrossData(lastCrossRaw, 'last');
  const newCrossData = cleanCrossData(newCrossRaw, 'new');

  const allCodigos = [...new Set([...lastCrossData, ...newCrossData].map(c => c.Codigo_Cliente))];

  const crossComparison = allCodigos.map(codigo => {
    const last = lastCrossData.find(c => c.Codigo_Cliente === codigo) || {};
    const current = newCrossData.find(c => c.Codigo_Cliente === codigo) || {};

    const ICS_BE_last = last.ICS_BE_last ?? 0;
    const ICS_BE_new = current.ICS_BE_new ?? 0;
    const diferencia = ICS_BE_new - ICS_BE_last;

    let cambio;
    if (last.ICS_BE_last == null && current.ICS_BE_new != null) cambio = "Cuenta nueva";
    else if (last.ICS_BE_last != null && current.ICS_BE_new == null) cambio = "Cuenta eliminada";
    else if (diferencia > 0) cambio = "Aumentó";
    else if (diferencia < 0) cambio = "Disminuyó";
    else cambio = "Igual";

    const extras = {};
    camposExtras.forEach(campo => {
      const valLast = last[`${campo}_last`] ?? 0;
      const valNew = current[`${campo}_new`] ?? 0;
      const diff = valNew - valLast;
      extras[campo] = {
        last: valLast,
        new: valNew,
        diff
      };
    });

    return {
      Codigo_Cliente: codigo,
      Nombre: last.Nombre_last || current.Nombre_new,
      Segmento: last.Segmento_last || current.Segmento_new,
      ICS_BE_last,
      ICS_BE_new,
      Diferencia: diferencia,
      Cambio: cambio,
      extras
    };
  });

  const resumen = {};
  let totalNeto = 0;

  crossComparison.forEach(r => {
    totalNeto += r.Diferencia;
    if (!resumen[r.Cambio]) {
      resumen[r.Cambio] = { Cantidad: 0, Total_Diferencia: 0 };
    }
    resumen[r.Cambio].Cantidad += 1;
    resumen[r.Cambio].Total_Diferencia += r.Diferencia;
  });

  crossComparison.sort((a, b) => {
    const order = {
      'Disminuyó': 0,
      'Aumentó': 1,
      'Igual': 2,
      'Cuenta nueva': 3,
      'Cuenta eliminada': 4
    };
    const aOrder = order[a.Cambio] ?? 5;
    const bOrder = order[b.Cambio] ?? 5;
    return aOrder - bOrder;
  });

  showResults(crossComparison, resumen, totalNeto);
  lastResults = crossComparison; // Para exportar después

}

function showResults(crossComparison, resumen, totalNeto) {
  const output = document.getElementById('output');
  output.innerHTML = '';

  const table = document.createElement('table');
  const header = `<tr>
    <th>numero</th><th>Nombre</th><th>Segmento</th><th>ICS anterior</th>
    <th>cambio</th><th>ICS nuevo</th>
    ${camposExtras.map(c => `<th>${c} anterior</th><th>Cambio</th><th>${c} nuevo</th>`).join('')}
  </tr>`;

  table.innerHTML = header + crossComparison.map(row => {
    const extrasHtml = camposExtras.map(campo => {
      const data = row.extras[campo];
      return `
        <td>${data.last}</td>
        <td style="background-color: ${data.diff > 0 ? '#4CAF50' : data.diff < 0 ? '#f44336' : 'transparent'}; color: ${data.diff !== 0 ? 'white' : 'black'}; font-weight: bold; text-align: center;">
          ${data.diff > 0 ? '+' + data.diff : data.diff < 0 ? data.diff : '0'}
        </td>
        <td>${data.new}</td>
      `;
    }).join('');

    return `
      <tr>
        <td>${row.Codigo_Cliente}</td>
        <td>${row.Nombre}</td>
        <td>${row.Segmento}</td>
        <td>${row.ICS_BE_last}</td>
        <td style="
          background-color: ${(row.Cambio === 'Aumentó' ? '#4CAF50' : row.Cambio === 'Disminuyó' ? '#f44336' : 'transparent')};
          color: ${(row.Cambio === 'Aumentó' || row.Cambio === 'Disminuyó') ? 'white' : 'black'};
          font-weight: bold;
          text-align: center;
        ">
          ${row.Cambio === 'Aumentó' ? '+' + row.Diferencia : row.Cambio === 'Disminuyó' ? row.Diferencia : row.Cambio}
        </td>
        <td>${row.ICS_BE_new}</td>
        ${extrasHtml}
      </tr>
    `;
  }).join('');

  output.appendChild(table);

  const resumenTable = document.createElement('table');
  resumenTable.innerHTML = `<tr><th>Cambio</th><th>Cantidad</th><th>Total Diferencia</th></tr>` +
    Object.entries(resumen).map(([cambio, info]) => `
      <tr><td>${cambio}</td><td>${info.Cantidad}</td><td>${info.Total_Diferencia}</td></tr>
    `).join('');
  output.appendChild(document.createElement('hr'));
  output.appendChild(resumenTable);

  const neto = document.createElement('p');
  neto.innerHTML = `<strong>Total Neto:</strong> ${totalNeto}`;
  output.appendChild(neto);
}
function exportToExcel() {
  const headers = [
    "Codigo_Cliente", "Nombre", "Segmento", "ICS_BE_last", "Diferencia", "ICS_BE_new",
    ...camposExtras.flatMap(c => [`${c}_last`, `${c}_diff`, `${c}_new`])
  ];

  const rows = lastResults.map(row => {
    const data = {
      Codigo_Cliente: row.Codigo_Cliente,
      Nombre: row.Nombre,
      Segmento: row.Segmento,
      ICS_BE_last: row.ICS_BE_last,
      Diferencia: row.Diferencia,
      ICS_BE_new: row.ICS_BE_new
    };

    camposExtras.forEach(c => {
      data[`${c}_last`] = row.extras[c].last;
      data[`${c}_diff`] = row.extras[c].diff;
      data[`${c}_new`] = row.extras[c].new;
    });

    return data;
  });

  const worksheet = XLSX.utils.json_to_sheet(rows, { header: headers });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Comparación");

  XLSX.writeFile(workbook, "comparacion_clientes.xlsx");
}
