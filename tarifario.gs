function normalizeText(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function parseNumber(value) {
  if (typeof value === "number") return value;
  const cleaned = String(value || "").replace(/[^0-9.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function esMXN(value) {
  const moneda = normalizeText(value);
  return moneda === "MXN" || moneda === "PESOS" || moneda === "MEX";
}

function formatCurrency(value, moneda) {
  const amount = Number(value) || 0;
  if (esMXN(moneda)) {
    return `MX$ ${amount.toFixed(2)}`;
  }
  return `USD ${amount.toFixed(2)}`;
}

function getHeaderIndexMap(headers) {
  const normalizedHeaders = headers.map(header => normalizeText(header));
  const findHeaderIndex = (aliases, startIndex = 0) => {
    const normalizedAliases = aliases.map(alias => normalizeText(alias));
    for (let i = Math.max(0, startIndex); i < normalizedHeaders.length; i += 1) {
      if (normalizedAliases.includes(normalizedHeaders[i])) {
        return i;
      }
    }
    return -1;
  };

  const idx = {
    tipoOperacion: findHeaderIndex(["TIPO_DE_OPERACION", "TIPO DE OPERACION"]),
    tipoViaje: findHeaderIndex(["TIPO_DE_VIAJE", "TIPO DE VIAJE"]),
    tipoUnidad: findHeaderIndex(["TIPO_UNIDAD", "TIPO UNIDAD"]),
    transportista: findHeaderIndex(["TRANSPORTISTA"]),
    cliente: findHeaderIndex(["CLIENTE"]),
    paisOrigen: findHeaderIndex(["PAIS_ORIGEN", "PAIS ORIGEN"]),
    estadoOrigen: findHeaderIndex(["ESTADO_ORIGEN", "ESTADO ORIGEN"]),
    ciudadOrigen: findHeaderIndex(["CIUDAD_ORIGEN", "CIUDAD ORIGEN"]),
    paisDestino: findHeaderIndex(["PAIS_DESTINO", "PAIS DESTINO"]),
    estadoDestino: findHeaderIndex(["ESTADO_DESTINO", "ESTADO DESTINO"]),
    ciudadDestino: findHeaderIndex(["CIUDAD_DESTINO", "CIUDAD DESTINO"]),
    valorPeso: findHeaderIndex(["VALOR PESO", "VALOR_PESO"]),
    monedaAllIn: findHeaderIndex(["MONEDA", "MONEDA ALL IN"]),
    allIn: findHeaderIndex(["ALL IN", "ALL_IN"]),
    precioSencillo: findHeaderIndex(["PRECIO VIAJE SENCILLO", "TARIFA VIAJE SENCILLO"]),
    precioSencilloAlt: -1,
    monedaSencillo: findHeaderIndex(["MONEDA_BASE", "MONEDA BASE", "MONEDA VIAJE SENCILLO"]),
    precioRedondo: findHeaderIndex(["PRECIO VIAJE REDONDO", "TARIFA VIAJE FULL", "TARIFA VIAJE REDONDO"]),
    precioRedondoAlt: -1,
    monedaRedondo: -1,
    fechaFin: findHeaderIndex(["FECHA_VIGENCIA_FIN", "FECHA VIGENCIA FIN"])
  };

  idx.precioSencilloAlt = findHeaderIndex(
    ["PRECIO VIAJE SENCILLO", "TARIFA VIAJE SENCILLO"],
    idx.precioSencillo + 1
  );
  idx.precioRedondoAlt = findHeaderIndex(
    ["PRECIO VIAJE REDONDO", "TARIFA VIAJE FULL", "TARIFA VIAJE REDONDO"],
    idx.precioRedondo + 1
  );
  idx.monedaRedondo = findHeaderIndex(
    ["MONEDA_BASE_REDONDO", "MONEDA BASE REDONDO", "MONEDA VIAJE REDONDO", "MONEDA_BASE", "MONEDA BASE"],
    idx.monedaSencillo + 1
  );
  if (idx.monedaRedondo < 0) {
    idx.monedaRedondo = idx.monedaSencillo;
  }

  return { idx, normalizedHeaders };
}

function seleccionarTarifaLocal(row, opts) {
  const { AGIndex, AHIndex, AIIndex, AJIndex, idx, tipoCambio, normalizeText, esMXN, parseNumber } = opts;

  const sencillo = AGIndex >= 0 ? parseNumber(row[AGIndex]) : 0;
  const monedaSencillo = AHIndex >= 0 ? row[AHIndex] : "";
  const redondo = AIIndex >= 0 ? parseNumber(row[AIIndex]) : 0;
  const monedaRedondo = AJIndex >= 0 ? row[AJIndex] : "";
  const tipoViaje = idx && idx.tipoViaje >= 0 ? normalizeText(row[idx.tipoViaje]) : "";

  if (tipoViaje.includes("SENCILLO") && sencillo > 0 && monedaSencillo) {
    return { tarifa: sencillo, moneda: monedaSencillo };
  }

  if (tipoViaje.includes("SENCILLO") && (redondo > 0 && monedaRedondo)) {
    return { tarifa: redondo, moneda: monedaRedondo };
  }

  if ((tipoViaje.includes("REDONDO") || tipoViaje.includes("FULL")) && redondo > 0 && monedaRedondo) {
    return { tarifa: redondo, moneda: monedaRedondo };
  }

  if ((tipoViaje.includes("REDONDO") || tipoViaje.includes("FULL")) && sencillo > 0 && monedaSencillo) {
    return { tarifa: sencillo, moneda: monedaSencillo };
  }

  const opciones = [];

  if (sencillo > 0 && monedaSencillo) {
    opciones.push({
      tarifa: sencillo,
      moneda: monedaSencillo,
      usd: esMXN(monedaSencillo) ? sencillo / tipoCambio : sencillo
    });
  }

  if (redondo > 0 && monedaRedondo) {
    opciones.push({
      tarifa: redondo,
      moneda: monedaRedondo,
      usd: esMXN(monedaRedondo) ? redondo / tipoCambio : redondo
    });
  }

  if (opciones.length === 0) {
    return null;
  }

  opciones.sort((a, b) => a.usd - b.usd);
  return opciones[0];
}

function generarReporteFinal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTarifario = ss.getSheetByName("Tarifario estandar");
  const shReporte = ss.getSheetByName("Reporte Final");

  if (!shTarifario || !shReporte) {
    throw new Error("No se encontró la hoja 'Tarifario estandar' o 'Reporte Final'.");
  }

  const rawTipoCambio = shReporte.getRange("C2").getValue();
  const parsedTipoCambio = Number(String(rawTipoCambio || "").replace(/[^0-9.]/g, ""));
  const tipoCambio = Number.isFinite(parsedTipoCambio) && parsedTipoCambio > 0 ? parsedTipoCambio : 1;

  const filtros = {
    tipoOperacion: shReporte.getRange("C3").getValue(),
    tipoViaje: shReporte.getRange("C4").getValue(),
    tipoUnidad: shReporte.getRange("C5").getValue(),
    transportista: shReporte.getRange("C6").getValue(),
    cliente: shReporte.getRange("C7").getValue(),
    paisOrigen: shReporte.getRange("C8").getValue(),
    estadoOrigen: shReporte.getRange("C9").getValue(),
    ciudadOrigen: shReporte.getRange("C10").getValue(),
    paisDestino: shReporte.getRange("C11").getValue(),
    estadoDestino: shReporte.getRange("C12").getValue(),
    ciudadDestino: shReporte.getRange("C13").getValue(),
    valorPeso: shReporte.getRange("C14").getValue(),
    mostrarFechas: shReporte.getRange("C15").getValue()
  };

  const lastRow = shTarifario.getLastRow();
  if (lastRow < 2) return;

  const lastCol = shTarifario.getLastColumn();
  const headers = shTarifario.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = shTarifario.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const { idx, normalizedHeaders } = getHeaderIndexMap(headers);

  const resultados = [];
  const hoy = new Date();
  let vencidas = 0;

  data.forEach(row => {
    const match = (value, filtro) => !filtro || normalizeText(value) === normalizeText(filtro);

    if (
      (idx.tipoOperacion >= 0 && !match(row[idx.tipoOperacion], filtros.tipoOperacion)) ||
      (idx.tipoViaje >= 0 && !match(row[idx.tipoViaje], filtros.tipoViaje)) ||
      (idx.tipoUnidad >= 0 && !match(row[idx.tipoUnidad], filtros.tipoUnidad)) ||
      (idx.transportista >= 0 && !match(row[idx.transportista], filtros.transportista)) ||
      (idx.cliente >= 0 && !match(row[idx.cliente], filtros.cliente)) ||
      (idx.paisOrigen >= 0 && !match(row[idx.paisOrigen], filtros.paisOrigen)) ||
      (idx.estadoOrigen >= 0 && !match(row[idx.estadoOrigen], filtros.estadoOrigen)) ||
      (idx.ciudadOrigen >= 0 && !match(row[idx.ciudadOrigen], filtros.ciudadOrigen)) ||
      (idx.paisDestino >= 0 && !match(row[idx.paisDestino], filtros.paisDestino)) ||
      (idx.estadoDestino >= 0 && !match(row[idx.estadoDestino], filtros.estadoDestino)) ||
      (idx.ciudadDestino >= 0 && !match(row[idx.ciudadDestino], filtros.ciudadDestino)) ||
      (idx.valorPeso >= 0 && !match(row[idx.valorPeso], filtros.valorPeso))
    ) {
      return;
    }

    const tipoOperacion = idx.tipoOperacion >= 0 ? normalizeText(row[idx.tipoOperacion]) : "";
    const esInternacional = tipoOperacion.includes("IMPORT") || tipoOperacion.includes("EXPORT");

    let tarifaSeleccionada = 0;
    let monedaSeleccionada = "";

    if (esInternacional) {
      tarifaSeleccionada = idx.allIn >= 0 ? parseNumber(row[idx.allIn]) : 0;
      monedaSeleccionada = idx.monedaAllIn >= 0 ? row[idx.monedaAllIn] : "";
    } else {
      const seleccion = seleccionarTarifaLocal(row, {
        AGIndex: idx.precioSencillo,
        AHIndex: idx.monedaSencillo,
        AIIndex: idx.precioRedondo,
        AJIndex: idx.monedaRedondo,
        idx,
        tipoCambio,
        normalizeText,
        esMXN,
        parseNumber
      });
      if (!seleccion) {
        return;
      }
      tarifaSeleccionada = seleccion.tarifa;
      monedaSeleccionada = seleccion.moneda;
    }

    if (tarifaSeleccionada <= 0 || !monedaSeleccionada) {
      return;
    }

    const tarifaUSD = esMXN(monedaSeleccionada)
      ? parseNumber(tarifaSeleccionada) / tipoCambio
      : parseNumber(tarifaSeleccionada);

    const fechaFin = idx.fechaFin >= 0 ? row[idx.fechaFin] : "";
    const fechaFinDate = fechaFin ? new Date(fechaFin) : null;
    const estaVencida = fechaFinDate instanceof Date && !Number.isNaN(fechaFinDate) && fechaFinDate < hoy;
    if (estaVencida) {
      vencidas += 1;
    }

    resultados.push({
      row,
      tarifaSeleccionada,
      monedaSeleccionada,
      tarifaUSD,
      estaVencida
    });
  });

  const startRow = 20;
  const startCol = 1;
  const outputRows = shReporte.getMaxRows() - startRow + 1;
  const outputCols = shReporte.getMaxColumns();

  shReporte.getRange(startRow, startCol, outputRows, outputCols)
    .clearContent()
    .clearFormat();

  const reportColumns = [
    { label: "TRANSPORTISTA", aliases: ["TRANSPORTISTA"] },
    { label: "TIPO_UNIDAD", aliases: ["TIPO_UNIDAD", "TIPO UNIDAD"] },
    { label: "PRECIO VIAJE SENCILLO", aliases: ["PRECIO VIAJE SENCILLO", "TARIFA VIAJE SENCILLO"] },
    { label: "PRECIO VIAJE REDONDO", aliases: ["PRECIO VIAJE REDONDO", "TARIFA VIAJE FULL"] },
    { label: "ALL IN", aliases: ["ALL IN", "ALL_IN"] },
    { label: "MONEDA", aliases: ["MONEDA", "MONEDA ALL IN", "MONEDA_BASE", "MONEDA BASE"] }
  ];
  const getColumnIndex = aliases => {
    const normalizedAliases = aliases.map(alias => normalizeText(alias));
    return normalizedHeaders.findIndex(header => normalizedAliases.includes(header));
  };
  const reportIndexes = reportColumns.map(column => getColumnIndex(column.aliases));

  const outputHeaders = [
    ...reportColumns.map(column => column.label),
    "MEJOR_TARIFA"
  ];
  shReporte.getRange(startRow, startCol, 1, outputHeaders.length).setValues([outputHeaders]);

  if (resultados.length > 0) {
    const groupKey = row => {
      const parts = [
        idx.tipoOperacion >= 0 ? row[idx.tipoOperacion] : "",
        idx.tipoViaje >= 0 ? row[idx.tipoViaje] : "",
        idx.tipoUnidad >= 0 ? row[idx.tipoUnidad] : "",
        idx.paisOrigen >= 0 ? row[idx.paisOrigen] : "",
        idx.estadoOrigen >= 0 ? row[idx.estadoOrigen] : "",
        idx.ciudadOrigen >= 0 ? row[idx.ciudadOrigen] : "",
        idx.paisDestino >= 0 ? row[idx.paisDestino] : "",
        idx.estadoDestino >= 0 ? row[idx.estadoDestino] : "",
        idx.ciudadDestino >= 0 ? row[idx.ciudadDestino] : ""
      ];
      return parts.map(normalizeText).join("|");
    };

    const mejores = new Map();
    resultados.forEach(item => {
      const key = groupKey(item.row);
      const actual = mejores.get(key);
      if (!actual || item.tarifaUSD < actual.tarifaUSD) {
        mejores.set(key, item);
      }
    });

    const outputRowsData = resultados.map((item) => {
      const { row, tarifaSeleccionada, monedaSeleccionada, tarifaUSD, estaVencida } = item;
      const mejorTarifa = mejores.get(groupKey(row)) === item ? "SI" : "NO";

      return [
        ...reportIndexes.map(index => (index >= 0 ? row[index] : "")),
        mejorTarifa
      ];
    });

    shReporte.getRange(startRow + 1, startCol, outputRowsData.length, outputHeaders.length)
      .setValues(outputRowsData);

    const headerRange = shReporte.getRange(startRow, startCol, outputRowsData.length + 1, outputHeaders.length);
    headerRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

    shReporte.getRange(startRow, startCol, 1, outputHeaders.length)
      .setBackground("#1F4E79")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    const dataRange = shReporte.getRange(startRow + 1, startCol, outputRowsData.length, outputHeaders.length);
    dataRange.setVerticalAlignment("middle").setHorizontalAlignment("center");

  }

  shReporte.autoResizeColumns(1, outputHeaders.length);

  if (vencidas > 0) {
    SpreadsheetApp.getUi().alert(`Hay ${vencidas} tarifas vencidas en el reporte.`);
  }
}

function generarMejoresOpciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTarifario = ss.getSheetByName("Tarifario estandar");
  const shReporte = ss.getSheetByName("Reporte Final");

  if (!shTarifario || !shReporte) {
    throw new Error("No se encontró la hoja 'Tarifario estandar' o 'Reporte Final'.");
  }

  const rawTipoCambio = shReporte.getRange("C2").getValue();
  const parsedTipoCambio = Number(String(rawTipoCambio || "").replace(/[^0-9.]/g, ""));
  const tipoCambio = Number.isFinite(parsedTipoCambio) && parsedTipoCambio > 0 ? parsedTipoCambio : 1;

  const filtros = {
    tipoOperacion: shReporte.getRange("C3").getValue(),
    tipoViaje: shReporte.getRange("C4").getValue(),
    tipoUnidad: shReporte.getRange("C5").getValue(),
    transportista: shReporte.getRange("C6").getValue(),
    cliente: shReporte.getRange("C7").getValue(),
    paisOrigen: shReporte.getRange("C8").getValue(),
    estadoOrigen: shReporte.getRange("C9").getValue(),
    ciudadOrigen: shReporte.getRange("C10").getValue(),
    paisDestino: shReporte.getRange("C11").getValue(),
    estadoDestino: shReporte.getRange("C12").getValue(),
    ciudadDestino: shReporte.getRange("C13").getValue(),
    valorPeso: shReporte.getRange("C14").getValue()
  };

  const lastRow = shTarifario.getLastRow();
  if (lastRow < 2) return;

  const lastCol = shTarifario.getLastColumn();
  const headers = shTarifario.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = shTarifario.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const normalizeText = value => String(value || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");

  const normalizedHeaders = headers.map(header => normalizeText(header));
  const findHeaderIndex = aliases => {
    const normalizedAliases = aliases.map(alias => normalizeText(alias));
    return normalizedHeaders.findIndex(header => normalizedAliases.includes(header));
  };

  const idx = {
    tipoOperacion: findHeaderIndex(["TIPO_DE_OPERACION", "TIPO DE OPERACION"]),
    tipoViaje: findHeaderIndex(["TIPO_DE_VIAJE", "TIPO DE VIAJE"]),
    tipoUnidad: findHeaderIndex(["TIPO_UNIDAD", "TIPO UNIDAD"]),
    transportista: findHeaderIndex(["TRANSPORTISTA"]),
    cliente: findHeaderIndex(["CLIENTE"]),
    paisOrigen: findHeaderIndex(["PAIS_ORIGEN", "PAIS ORIGEN"]),
    estadoOrigen: findHeaderIndex(["ESTADO_ORIGEN", "ESTADO ORIGEN"]),
    ciudadOrigen: findHeaderIndex(["CIUDAD_ORIGEN", "CIUDAD ORIGEN"]),
    paisDestino: findHeaderIndex(["PAIS_DESTINO", "PAIS DESTINO"]),
    estadoDestino: findHeaderIndex(["ESTADO_DESTINO", "ESTADO DESTINO"]),
    ciudadDestino: findHeaderIndex(["CIUDAD_DESTINO", "CIUDAD DESTINO"]),
    valorPeso: findHeaderIndex(["VALOR PESO", "VALOR_PESO"])
  };

  const AE = findHeaderIndex(["MONEDA", "MONEDA ALL IN"]);
  const AF = findHeaderIndex(["ALL IN", "ALL_IN"]);
  const AG = findHeaderIndex(["PRECIO VIAJE SENCILLO", "TARIFA VIAJE SENCILLO"]);
  const AH = findHeaderIndex(["MONEDA_BASE", "MONEDA BASE", "MONEDA VIAJE SENCILLO"]);
  const AI = findHeaderIndex(["PRECIO VIAJE REDONDO", "TARIFA VIAJE FULL", "TARIFA VIAJE REDONDO"]);
  const AJ = findHeaderIndex(["MONEDA_BASE_REDONDO", "MONEDA BASE REDONDO", "MONEDA VIAJE REDONDO"]);

  const AEIndex = AE;
  const AFIndex = AF;
  const AGIndex = AG;
  const AHIndex = AH;
  const AIIndex = AI;
  const AJIndex = AJ;

  const esMXN = value => {
    const moneda = normalizeText(value);
    return moneda === "MXN" || moneda === "PESOS";
  };

  const parseNumber = value => {
    if (typeof value === "number") return value;
    const cleaned = String(value || "").replace(/[^0-9.-]/g, "");
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : 0;
  };

  const match = (value, filtro) => !filtro || normalizeText(value) === normalizeText(filtro);

  const formatCurrency = (value, moneda) => {
    const amount = Number(value) || 0;
    if (esMXN(moneda)) {
      return `MX$ ${amount.toFixed(2)}`;
    }
    return `$ ${amount.toFixed(2)}`;
  };

  const fechaFinIndex = findHeaderIndex(["FECHA_VIGENCIA_FIN", "FECHA VIGENCIA FIN"]);
  let mejorOpcion = null;

  data.forEach(row => {
    if (
      (idx.tipoOperacion >= 0 && !match(row[idx.tipoOperacion], filtros.tipoOperacion)) ||
      (idx.tipoViaje >= 0 && !match(row[idx.tipoViaje], filtros.tipoViaje)) ||
      (idx.tipoUnidad >= 0 && !match(row[idx.tipoUnidad], filtros.tipoUnidad)) ||
      (idx.transportista >= 0 && !match(row[idx.transportista], filtros.transportista)) ||
      (idx.cliente >= 0 && !match(row[idx.cliente], filtros.cliente)) ||
      (idx.paisOrigen >= 0 && !match(row[idx.paisOrigen], filtros.paisOrigen)) ||
      (idx.estadoOrigen >= 0 && !match(row[idx.estadoOrigen], filtros.estadoOrigen)) ||
      (idx.ciudadOrigen >= 0 && !match(row[idx.ciudadOrigen], filtros.ciudadOrigen)) ||
      (idx.paisDestino >= 0 && !match(row[idx.paisDestino], filtros.paisDestino)) ||
      (idx.estadoDestino >= 0 && !match(row[idx.estadoDestino], filtros.estadoDestino)) ||
      (idx.ciudadDestino >= 0 && !match(row[idx.ciudadDestino], filtros.ciudadDestino)) ||
      (idx.valorPeso >= 0 && !match(row[idx.valorPeso], filtros.valorPeso))
    ) {
      return;
    }

    const tipoOperacion = idx.tipoOperacion >= 0 ? normalizeText(row[idx.tipoOperacion]) : "";
    const esInternacional = tipoOperacion.includes("IMPORT") || tipoOperacion.includes("EXPORT");

    let tarifaSeleccionada = 0;
    let monedaSeleccionada = "";

    if (esInternacional) {
      if (AFIndex < 0 || AEIndex < 0) {
        return;
      }
      tarifaSeleccionada = parseNumber(row[AFIndex]);
      monedaSeleccionada = row[AEIndex] || "";
    } else {
      const seleccion = seleccionarTarifaLocal(row, {
        AGIndex,
        AHIndex,
        AIIndex,
        AJIndex,
        idx,
        tipoCambio,
        normalizeText,
        esMXN,
        parseNumber
      });
      if (!seleccion) {
        return;
      }
      tarifaSeleccionada = seleccion.tarifa;
      monedaSeleccionada = seleccion.moneda;
    }

    if (tarifaSeleccionada <= 0 || !monedaSeleccionada) {
      return;
    }

    const tarifaUSD = esMXN(monedaSeleccionada)
      ? parseNumber(tarifaSeleccionada) / tipoCambio
      : parseNumber(tarifaSeleccionada);

    if (!mejorOpcion || tarifaUSD < mejorOpcion.tarifaUSD) {
      mejorOpcion = {
        row,
        tarifaSeleccionada,
        monedaSeleccionada,
        tarifaUSD
      };
    }
  });

  const startRow = 3;
  const startCol = 9; // I
  const outputRows = 10;
  const outputCols = 2;

  shReporte.getRange(startRow, startCol, outputRows, outputCols)
    .clearContent()
    .clearFormat();

  if (!mejorOpcion) {
    shReporte.getRange(startRow, startCol).setValue("SIN RESULTADOS");
    shReporte.getRange(startRow, startCol)
      .setFontWeight("bold")
      .setFontColor("#990000");
    return;
  }

  const { row, tarifaSeleccionada, monedaSeleccionada, tarifaUSD } = mejorOpcion;
  const ciudadOrigen = idx.ciudadOrigen >= 0 ? row[idx.ciudadOrigen] : "";
  const ciudadDestino = idx.ciudadDestino >= 0 ? row[idx.ciudadDestino] : "";
  const tipoOperacion = idx.tipoOperacion >= 0 ? row[idx.tipoOperacion] : "";
  const tipoUnidad = idx.tipoUnidad >= 0 ? row[idx.tipoUnidad] : "";
  const transportista = idx.transportista >= 0 ? row[idx.transportista] : "";
  const fechaFin = fechaFinIndex >= 0 ? row[fechaFinIndex] : "";

  const resumen = [
    ["COTIZACIÓN DISPONIBLE", ""],
    ["Tipo:", tipoOperacion],
    ["Ruta:", `${ciudadOrigen} \u2192 ${ciudadDestino}`],
    ["Unidad:", tipoUnidad],
    ["Mejor tarifa:", formatCurrency(tarifaUSD, "USD")],
    ["Transportista:", transportista],
    ["Vigencia:", fechaFin || "-"]
  ];

  shReporte.getRange(startRow, startCol, resumen.length, 2).setValues(resumen);

  shReporte.getRange(startRow, startCol, 1, 2)
    .merge()
    .setFontWeight("bold")
    .setFontColor("#2E7D32");

  const cardRange = shReporte.getRange(startRow, startCol, resumen.length, 2);
  cardRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  cardRange.setVerticalAlignment("middle");

  shReporte.getRange(startRow + 1, startCol, resumen.length - 1, 1)
    .setFontWeight("bold")
    .setFontColor("#6D4C41");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Bot Cotización")
    .addItem("Abrir Bot", "mostrarBotCotizacion")
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu("Cotización")
    .addItem("Enviar cotización", "openSendSidebar")
    .addToUi();
}

function mostrarBotCotizacion() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 16px;">
      <h2 style="margin-top: 0;">Bot de Cotización</h2>
      <p>Ingresa una pregunta en la celda K2 y ejecuta el bot.</p>
      <button onclick="cotizar()" style="width: 100%; padding: 10px; background: #1F4E79; color: #fff; border: 0; border-radius: 4px;">Cotizar</button>
      <pre id="resultado" style="margin-top: 16px; background: #f7f7f7; padding: 12px; border-radius: 4px; white-space: pre-wrap;"></pre>
    </div>
    <script>
      function cotizar() {
        document.getElementById('resultado').textContent = 'Procesando...';
        google.script.run
          .withSuccessHandler(res => {
            document.getElementById('resultado').textContent = res;
          })
          .withFailureHandler(err => {
            document.getElementById('resultado').textContent = 'Error: ' + err.message;
          })
          .botMejorTarifa();
      }
    </script>
  `).setTitle("Bot de Cotización");

  SpreadsheetApp.getUi().showSidebar(html);
}

function generarCotizacionBot() {
  return botMejorTarifa();
}

function escribirRespuestaBot(lineas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shReporte = ss.getSheetByName("Reporte Final");
  const startRow = 2;
  const startCol = 9; // I
  const maxRows = 9; // I2:I10

  const cleaned = lineas.filter(line => line !== undefined && line !== null);
  const output = cleaned.slice(0, maxRows).map(line => [line]);

  shReporte.getRange(startRow, startCol, maxRows, 1).clearContent().clearFormat();
  if (output.length > 0) {
    shReporte.getRange(startRow, startCol, output.length, 1).setValues(output);
  }

  const range = shReporte.getRange(startRow, startCol, maxRows, 1);
  range
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  if (output.length > 0) {
    shReporte.getRange(startRow, startCol)
      .setFontWeight("bold")
      .setFontColor("#1B5E20");
  }

  range.setBorder(true, true, true, true, true, true, "#9E9E9E", SpreadsheetApp.BorderStyle.SOLID);

  return cleaned.join("\n");
}

function botMejorTarifa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTarifario = ss.getSheetByName("Tarifario estandar");
  const shReporte = ss.getSheetByName("Reporte Final");

  if (!shTarifario || !shReporte) {
    throw new Error("No se encontró la hoja 'Tarifario estandar' o 'Reporte Final'.");
  }

  const rawTipoCambio = shReporte.getRange("C2").getValue();
  const parsedTipoCambio = Number(String(rawTipoCambio || "").replace(/[^0-9.]/g, ""));
  const tipoCambio = Number.isFinite(parsedTipoCambio) && parsedTipoCambio > 0 ? parsedTipoCambio : 1;

  const pregunta = shReporte.getRange("K2").getValue();
  const preguntaNormalizada = normalizeText(pregunta);
  const monedaPreferida = preguntaNormalizada.includes("MXN") || preguntaNormalizada.includes("PESOS")
    ? "MXN"
    : "USD";

  const filtros = {
    tipoOperacion: shReporte.getRange("C3").getValue(),
    tipoViaje: shReporte.getRange("C4").getValue(),
    tipoUnidad: shReporte.getRange("C5").getValue(),
    transportista: shReporte.getRange("C6").getValue(),
    cliente: shReporte.getRange("C7").getValue(),
    paisOrigen: shReporte.getRange("C8").getValue(),
    estadoOrigen: shReporte.getRange("C9").getValue(),
    ciudadOrigen: shReporte.getRange("C10").getValue(),
    paisDestino: shReporte.getRange("C11").getValue(),
    estadoDestino: shReporte.getRange("C12").getValue(),
    ciudadDestino: shReporte.getRange("C13").getValue(),
    valorPeso: shReporte.getRange("C14").getValue()
  };

  const lastRow = shTarifario.getLastRow();
  if (lastRow < 2) {
    return escribirRespuestaBot([
      "NO HAY TARIFA DISPONIBLE",
      "No se encontró ninguna tarifa vigente que cumpla los filtros seleccionados."
    ]);
  }

  const lastCol = shTarifario.getLastColumn();
  const headers = shTarifario.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = shTarifario.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const { idx } = getHeaderIndexMap(headers);

  const requiredHeaders = [
    ["TIPO_DE_OPERACION", idx.tipoOperacion],
    ["TIPO_DE_VIAJE", idx.tipoViaje],
    ["TIPO_UNIDAD", idx.tipoUnidad],
    ["TRANSPORTISTA", idx.transportista],
    ["FECHA_VIGENCIA_FIN", idx.fechaFin]
  ];

  const faltantes = requiredHeaders.filter(([, index]) => index < 0).map(([name]) => name);
  if (faltantes.length > 0) {
    return escribirRespuestaBot([
      "NO HAY TARIFA DISPONIBLE",
      "Encabezados faltantes en el tarifario:",
      ...faltantes.map(name => `- ${name}`)
    ]);
  }

  const fallbackIndex = (value, fallback) => (value >= 0 ? value : fallback);
  const fallbackSencillo = fallbackIndex(idx.precioSencillo, 32); // AG
  const fallbackMonedaSencillo = fallbackIndex(idx.monedaSencillo, 33); // AH
  const fallbackRedondo = fallbackIndex(idx.precioRedondo, 34); // AI
  const fallbackMonedaRedondo = fallbackIndex(idx.monedaRedondo, 35); // AJ

  const tieneSencillo = fallbackSencillo >= 0 && fallbackMonedaSencillo >= 0;
  const tieneRedondo = fallbackRedondo >= 0 && fallbackMonedaRedondo >= 0;
  const tieneAllIn = idx.allIn >= 0 && idx.monedaAllIn >= 0;

  if (!tieneSencillo && !tieneRedondo && !tieneAllIn) {
    return escribirRespuestaBot([
      "NO HAY TARIFA DISPONIBLE",
      "No se encontraron columnas de tarifa válidas para comparar."
    ]);
  }

  const match = (value, filtro) => !filtro || normalizeText(value) === normalizeText(filtro);
  const normalizeMoneda = moneda => {
    const valor = normalizeText(moneda);
    if (valor === "MXN" || valor === "PESOS" || valor === "MEX") return "MXN";
    if (["USD", "DOLARES", "DOLARES US", "US", "US DOLLARS"].includes(valor)) return "USD";
    return null;
  };

  const hoy = new Date();
  const parseFecha = value => {
    if (value instanceof Date) return value;
    const texto = String(value || "").trim();
    if (!texto) return null;
    const match = texto.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (!match) return null;
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    let year = Number(match[3]);
    if (year < 100) year += 2000;
    const fecha = new Date(year, month, day);
    return Number.isNaN(fecha.getTime()) ? null : fecha;
  };
  let mejorUSD = null;
  const mejoresEmpate = [];

  data.forEach(row => {
    if (
      (idx.tipoOperacion >= 0 && !match(row[idx.tipoOperacion], filtros.tipoOperacion)) ||
      (idx.tipoViaje >= 0 && !match(row[idx.tipoViaje], filtros.tipoViaje)) ||
      (idx.tipoUnidad >= 0 && !match(row[idx.tipoUnidad], filtros.tipoUnidad)) ||
      (idx.transportista >= 0 && !match(row[idx.transportista], filtros.transportista)) ||
      (idx.cliente >= 0 && !match(row[idx.cliente], filtros.cliente)) ||
      (idx.paisOrigen >= 0 && !match(row[idx.paisOrigen], filtros.paisOrigen)) ||
      (idx.estadoOrigen >= 0 && !match(row[idx.estadoOrigen], filtros.estadoOrigen)) ||
      (idx.ciudadOrigen >= 0 && !match(row[idx.ciudadOrigen], filtros.ciudadOrigen)) ||
      (idx.paisDestino >= 0 && !match(row[idx.paisDestino], filtros.paisDestino)) ||
      (idx.estadoDestino >= 0 && !match(row[idx.estadoDestino], filtros.estadoDestino)) ||
      (idx.ciudadDestino >= 0 && !match(row[idx.ciudadDestino], filtros.ciudadDestino)) ||
      (idx.valorPeso >= 0 && !match(row[idx.valorPeso], filtros.valorPeso))
    ) {
      return;
    }

    const fechaFin = row[idx.fechaFin];
    const fechaFinDate = parseFecha(fechaFin);
    if (!fechaFinDate || fechaFinDate < hoy) {
      return;
    }

    const tipoOperacion = normalizeText(row[idx.tipoOperacion]);
    const esInternacional = tipoOperacion.includes("IMPORT") || tipoOperacion.includes("EXPORT");

    let tarifaSeleccionada = 0;
    let monedaSeleccionada = "";

    if (esInternacional) {
      if (idx.allIn < 0 || idx.monedaAllIn < 0) {
        return;
      }
      tarifaSeleccionada = parseNumber(row[idx.allIn]);
      monedaSeleccionada = row[idx.monedaAllIn];
    } else {
    let seleccion = seleccionarTarifaLocal(row, {
      AGIndex: fallbackSencillo,
      AHIndex: fallbackMonedaSencillo,
      AIIndex: fallbackRedondo,
      AJIndex: fallbackMonedaRedondo,
      idx,
      tipoCambio,
      normalizeText,
      esMXN,
      parseNumber
    });
    if (!seleccion && idx.precioSencilloAlt >= 0) {
      const altSencillo = fallbackIndex(idx.precioSencilloAlt, fallbackSencillo);
      const altRedondo = fallbackIndex(idx.precioRedondoAlt, fallbackRedondo);
      seleccion = seleccionarTarifaLocal(row, {
        AGIndex: altSencillo,
        AHIndex: fallbackMonedaSencillo,
        AIIndex: altRedondo,
        AJIndex: fallbackMonedaRedondo,
        idx,
        tipoCambio,
        normalizeText,
        esMXN,
        parseNumber
      });
    }
    if (!seleccion) {
      return;
    }
    tarifaSeleccionada = seleccion.tarifa;
    monedaSeleccionada = seleccion.moneda;
    }

    const monedaNormalizada = normalizeMoneda(monedaSeleccionada);
    if (!monedaNormalizada || tarifaSeleccionada <= 0) {
      return;
    }

    const tarifaUSD = monedaNormalizada === "MXN"
      ? tarifaSeleccionada / tipoCambio
      : tarifaSeleccionada;

    if (!Number.isFinite(tarifaUSD) || tarifaUSD <= 0) {
      return;
    }

    if (mejorUSD === null || tarifaUSD < mejorUSD) {
      mejorUSD = tarifaUSD;
      mejoresEmpate.length = 0;
    }

    if (tarifaUSD === mejorUSD) {
      mejoresEmpate.push({
        row,
        tarifaSeleccionada,
        monedaNormalizada
      });
    }
  });

  if (mejoresEmpate.length === 0) {
    return escribirRespuestaBot([
      "NO HAY TARIFA DISPONIBLE",
      "No se encontró ninguna tarifa vigente que cumpla los filtros seleccionados."
    ]);
  }

  if (mejoresEmpate.length === 1) {
    const { row, tarifaSeleccionada, monedaNormalizada } = mejoresEmpate[0];
    const tipoViaje = row[idx.tipoViaje];
    const tipoUnidad = row[idx.tipoUnidad];
    const transportista = row[idx.transportista];
    const fechaFin = row[idx.fechaFin];
    const tarifaSalida = monedaPreferida === "MXN" && monedaNormalizada === "USD"
      ? tarifaSeleccionada * tipoCambio
      : monedaPreferida === "USD" && monedaNormalizada === "MXN"
        ? tarifaSeleccionada / tipoCambio
        : tarifaSeleccionada;

    const monedaSalida = monedaPreferida;

    return escribirRespuestaBot([
      "LA MEJOR OPCIÓN ES:",
      `Transportista: ${transportista}`,
      `Tarifa: ${formatCurrency(tarifaSalida, monedaSalida)}`,
      `Tipo de viaje: ${tipoViaje}`,
      `Unidad: ${tipoUnidad}`,
      `Vigencia: ${fechaFin}`,
      "Motivo: es la tarifa más baja vigente disponible."
    ]);
  }

  const resumen = [
    `SE ENCONTRARON ${mejoresEmpate.length} OPCIONES CON LA TARIFA MÁS BAJA:`
  ];

  mejoresEmpate.forEach((opcion, index) => {
    const transportista = opcion.row[idx.transportista];
    const tarifa = formatCurrency(opcion.tarifaSeleccionada, opcion.monedaNormalizada);
    resumen.push(`${index + 1}) ${transportista} – ${tarifa}`);
  });

  return escribirRespuestaBot(resumen);
}

function generarFormatoCotizacionDirector() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTarifario = ss.getSheetByName("Tarifario estandar");
  const shReporte = ss.getSheetByName("Reporte Final");

  if (!shTarifario || !shReporte) {
    throw new Error("No se encontró la hoja 'Tarifario estandar' o 'Reporte Final'.");
  }

  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt(
    "Formato de cotización",
    "Escribe: 1 (IMP/EXP) o 2 (LOC/FOR)",
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const opcion = normalizeText(respuesta.getResponseText());
  const esInternacional = opcion === "1";
  const esLocalForaneo = opcion === "2";

  if (!esInternacional && !esLocalForaneo) {
    ui.alert("Opción inválida. Usa: 1 (IMP/EXP) o 2 (LOC/FOR).");
    return;
  }

  const filtros = {
    tipoOperacion: shReporte.getRange("C3").getValue(),
    tipoViaje: shReporte.getRange("C4").getValue(),
    tipoUnidad: shReporte.getRange("C5").getValue(),
    transportista: shReporte.getRange("C6").getValue(),
    cliente: shReporte.getRange("C7").getValue(),
    paisOrigen: shReporte.getRange("C8").getValue(),
    estadoOrigen: shReporte.getRange("C9").getValue(),
    ciudadOrigen: shReporte.getRange("C10").getValue(),
    paisDestino: shReporte.getRange("C11").getValue(),
    estadoDestino: shReporte.getRange("C12").getValue(),
    ciudadDestino: shReporte.getRange("C13").getValue(),
    valorPeso: shReporte.getRange("C14").getValue()
  };

  const dataRange = shTarifario.getDataRange();
  const values = dataRange.getValues();
  const headers = values.shift();
  const normalizedHeaders = headers.map(header => normalizeText(header));

  const findHeaderIndex = aliases => {
    const normalizedAliases = aliases.map(alias => normalizeText(alias));
    return normalizedHeaders.findIndex(header => normalizedAliases.includes(header));
  };

  const columnIndexFromLetter = letter => {
    return letter.toUpperCase().split("").reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
  };

  const idx = {
    tipoOperacion: findHeaderIndex(["TIPO_DE_OPERACION", "TIPO DE OPERACION"]),
    tipoViaje: findHeaderIndex(["TIPO_DE_VIAJE", "TIPO DE VIAJE"]),
    tipoUnidad: findHeaderIndex(["TIPO_UNIDAD", "TIPO UNIDAD"]),
    transportista: findHeaderIndex(["TRANSPORTISTA"]),
    cliente: findHeaderIndex(["CLIENTE"]),
    paisOrigen: findHeaderIndex(["PAIS_ORIGEN", "PAIS ORIGEN"]),
    estadoOrigen: findHeaderIndex(["ESTADO_ORIGEN", "ESTADO ORIGEN"]),
    ciudadOrigen: findHeaderIndex(["CIUDAD_ORIGEN", "CIUDAD ORIGEN"]),
    paisDestino: findHeaderIndex(["PAIS_DESTINO", "PAIS DESTINO"]),
    estadoDestino: findHeaderIndex(["ESTADO_DESTINO", "ESTADO DESTINO"]),
    ciudadDestino: findHeaderIndex(["CIUDAD_DESTINO", "CIUDAD DESTINO"]),
    valorPeso: findHeaderIndex(["VALOR PESO", "VALOR_PESO"]),
    fechaIni: findHeaderIndex(["FECHA_VIGENCIA_INI", "FECHA VIGENCIA INI"]),
    fechaFin: findHeaderIndex(["FECHA_VIGENCIA_FIN", "FECHA VIGENCIA FIN"])
  };

  const columnasInternacional = [
    { label: "ORIGEN", index: findHeaderIndex(["DIRECCION DE RECOLECCION", "ORIGEN"]), fallback: columnIndexFromLetter("N") },
    { label: "DESTINO", index: findHeaderIndex(["DESTINO_DIRECCION", "DESTINO DIRECCION", "DESTINO"]), fallback: columnIndexFromLetter("T") },
    { label: "REQUERIMIENTO", index: findHeaderIndex(["REQUERIMIENTO"]), fallback: columnIndexFromLetter("AM") },
    { label: "MEXICO FREIGHT", index: findHeaderIndex(["MEXICO FREIGHT", "MEXICAN_FREIGHT"]), fallback: columnIndexFromLetter("V") },
    { label: "BORDER CROSSING", index: findHeaderIndex(["BORDER CROSSING", "CROSSING"]), fallback: columnIndexFromLetter("U") },
    { label: "USA FREIGHT", index: findHeaderIndex(["USA FREIGHT", "US FREIGHT"]), fallback: columnIndexFromLetter("AN") },
    { label: "ALL IN", index: findHeaderIndex(["ALL IN", "ALL_IN"]), fallback: columnIndexFromLetter("AF") },
    { label: "DOUBLE DRIVER (TEAM DRIVER)", index: findHeaderIndex(["DOUBLE DRIVER (TEAM DRIVER)", "DOUBLE DRIVER"]), fallback: columnIndexFromLetter("X") },
    { label: "COSTO DE CANCELACIÓN DE POSICIONAMIENTO DE UNIDAD (TRUCKING CANCEL FEE)", index: findHeaderIndex(["COSTO DE CANCELACIÓN DE POSICIONAMIENTO DE UNIDAD (TRUCKING CANCEL FEE)", "TRUCKING CANCEL FEE"]), fallback: columnIndexFromLetter("AO") },
    { label: "WAITING", index: findHeaderIndex(["WAITING"]), fallback: columnIndexFromLetter("AP") },
    { label: "COSTO DE WAITING CHARGE", index: findHeaderIndex(["COSTO DE WAITING CHARGE", "WAITING CHARGE"]), fallback: columnIndexFromLetter("AQ") },
    { label: "INSURANCE", index: findHeaderIndex(["INSURANCE", "SEGUROS"]), fallback: columnIndexFromLetter("AA") },
    { label: "FREE TIME (HOURS)", index: findHeaderIndex(["FREE TIME (HOURS)", "FREE TIME"]), fallback: columnIndexFromLetter("AR") },
    { label: "USD / HOUR", index: findHeaderIndex(["USD / HOUR", "USD/HOUR"]), fallback: columnIndexFromLetter("AE") },
    { label: "REMARK", index: findHeaderIndex(["REMARK", "OBSERVACIONES"]), fallback: columnIndexFromLetter("AL") },
    { label: "FECHA_VIGENCIA_INI", index: idx.fechaIni, fallback: -1 },
    { label: "FECHA_VIGENCIA_FIN", index: idx.fechaFin, fallback: -1 }
  ];

  const columnasLocal = [
    { label: "ORIGEN", index: findHeaderIndex(["DIRECCION DE RECOLECCION", "ORIGEN"]), fallback: columnIndexFromLetter("N") },
    { label: "DESTINO", index: findHeaderIndex(["DESTINO_DIRECCION", "DESTINO DIRECCION", "DESTINO"]), fallback: columnIndexFromLetter("T") },
    { label: "PRECIO VIAJE SENCILLO", index: findHeaderIndex(["PRECIO VIAJE SENCILLO", "TARIFA VIAJE SENCILLO"]), fallback: columnIndexFromLetter("AG") },
    { label: "PRECIO VIAJE REDONDO", index: findHeaderIndex(["PRECIO VIAJE REDONDO", "TARIFA VIAJE FULL", "TARIFA VIAJE REDONDO"]), fallback: columnIndexFromLetter("AI") },
    { label: "FECHA_VIGENCIA_INI", index: idx.fechaIni, fallback: -1 },
    { label: "FECHA_VIGENCIA_FIN", index: idx.fechaFin, fallback: -1 }
  ];

  const columnas = esInternacional ? columnasInternacional : columnasLocal;

  const match = (value, filtro) => !filtro || normalizeText(value) === normalizeText(filtro);
  const salida = [];
  const mejorValores = [];
  const parseNumero = value => {
    if (typeof value === "number") return value;
    const cleaned = String(value || "").replace(/[^0-9.-]/g, "");
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  };
  const salidaIndexAllIn = columnas.findIndex(columna => columna.label === "ALL IN");
  const salidaIndexSencillo = columnas.findIndex(columna => columna.label === "PRECIO VIAJE SENCILLO");
  const salidaIndexRedondo = columnas.findIndex(columna => columna.label === "PRECIO VIAJE REDONDO");

  values.forEach(row => {
    if (
      (idx.tipoOperacion >= 0 && !match(row[idx.tipoOperacion], filtros.tipoOperacion)) ||
      (idx.tipoViaje >= 0 && !match(row[idx.tipoViaje], filtros.tipoViaje)) ||
      (idx.tipoUnidad >= 0 && !match(row[idx.tipoUnidad], filtros.tipoUnidad)) ||
      (idx.transportista >= 0 && !match(row[idx.transportista], filtros.transportista)) ||
      (idx.cliente >= 0 && !match(row[idx.cliente], filtros.cliente)) ||
      (idx.paisOrigen >= 0 && !match(row[idx.paisOrigen], filtros.paisOrigen)) ||
      (idx.estadoOrigen >= 0 && !match(row[idx.estadoOrigen], filtros.estadoOrigen)) ||
      (idx.ciudadOrigen >= 0 && !match(row[idx.ciudadOrigen], filtros.ciudadOrigen)) ||
      (idx.paisDestino >= 0 && !match(row[idx.paisDestino], filtros.paisDestino)) ||
      (idx.estadoDestino >= 0 && !match(row[idx.estadoDestino], filtros.estadoDestino)) ||
      (idx.ciudadDestino >= 0 && !match(row[idx.ciudadDestino], filtros.ciudadDestino)) ||
      (idx.valorPeso >= 0 && !match(row[idx.valorPeso], filtros.valorPeso))
    ) {
      return;
    }

    const tipoOperacion = idx.tipoOperacion >= 0 ? normalizeText(row[idx.tipoOperacion]) : "";
    if (esInternacional && !(tipoOperacion.includes("IMPORT") || tipoOperacion.includes("EXPORT"))) {
      return;
    }
    if (esLocalForaneo && (tipoOperacion.includes("IMPORT") || tipoOperacion.includes("EXPORT"))) {
      return;
    }

    const fila = columnas.map(columna => {
      const index = columna.index >= 0 ? columna.index : columna.fallback;
      return index >= 0 ? row[index] : "";
    });
    let mejorValor = null;
    if (esInternacional && salidaIndexAllIn >= 0) {
      mejorValor = parseNumero(fila[salidaIndexAllIn]);
    } else if (!esInternacional) {
      const sencillo = salidaIndexSencillo >= 0 ? parseNumero(fila[salidaIndexSencillo]) : null;
      const redondo = salidaIndexRedondo >= 0 ? parseNumero(fila[salidaIndexRedondo]) : null;
      if (sencillo !== null && redondo !== null) {
        mejorValor = Math.min(sencillo, redondo);
      } else {
        mejorValor = sencillo !== null ? sencillo : redondo;
      }
    }
    salida.push(fila);
    mejorValores.push(mejorValor);
  });

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const nombreHoja = `FORMATO_COTIZACION_${timestamp}`;
  const shDestino = ss.insertSheet(nombreHoja);
  const headersSalida = columnas.map(columna => columna.label);

  shDestino.getRange(1, 1, 1, headersSalida.length).setValues([headersSalida]);
  if (salida.length > 0) {
    shDestino.getRange(2, 1, salida.length, headersSalida.length).setValues(salida);
    const candidatos = mejorValores
      .map((valor, index) => ({ valor, index }))
      .filter(item => item.valor !== null && item.valor > 0);
    if (candidatos.length > 0) {
      candidatos.sort((a, b) => a.valor - b.valor);
      const rowIndex = candidatos[0].index + 2;
      shDestino.getRange(rowIndex, 1, 1, headersSalida.length).setBackground("#FFF59D");
    }
  }

  shDestino.setFrozenRows(1);
  shDestino.getRange(1, 1, 1, headersSalida.length)
    .setBackground("#1F4E79")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  shDestino.autoResizeColumns(1, headersSalida.length);
}

function generarFormatoCotizacionExportacionEconomica() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTarifario = ss.getSheetByName("Tarifario estandar");
  const shReporte = ss.getSheetByName("Reporte Final");

  if (!shTarifario || !shReporte) {
    throw new Error("No se encontró la hoja 'Tarifario estandar' o 'Reporte Final'.");
  }

  const filtros = {
    tipoOperacion: shReporte.getRange("C3").getValue(),
    tipoViaje: shReporte.getRange("C4").getValue(),
    tipoUnidad: shReporte.getRange("C5").getValue(),
    transportista: shReporte.getRange("C6").getValue(),
    cliente: shReporte.getRange("C7").getValue(),
    paisOrigen: shReporte.getRange("C8").getValue(),
    estadoOrigen: shReporte.getRange("C9").getValue(),
    ciudadOrigen: shReporte.getRange("C10").getValue(),
    paisDestino: shReporte.getRange("C11").getValue(),
    estadoDestino: shReporte.getRange("C12").getValue(),
    ciudadDestino: shReporte.getRange("C13").getValue(),
    valorPeso: shReporte.getRange("C14").getValue()
  };

  const dataRange = shTarifario.getDataRange();
  const values = dataRange.getValues();
  const headers = values.shift();
  const normalizedHeaders = headers.map(header => normalizeText(header));

  const findHeaderIndex = aliases => {
    const normalizedAliases = aliases.map(alias => normalizeText(alias));
    return normalizedHeaders.findIndex(header => normalizedAliases.includes(header));
  };

  const idx = {
    tipoOperacion: findHeaderIndex(["TIPO_DE_OPERACION", "TIPO DE OPERACION"]),
    tipoViaje: findHeaderIndex(["TIPO_DE_VIAJE", "TIPO DE VIAJE"]),
    tipoUnidad: findHeaderIndex(["TIPO_UNIDAD", "TIPO UNIDAD"]),
    transportista: findHeaderIndex(["TRANSPORTISTA"]),
    cliente: findHeaderIndex(["CLIENTE"]),
    paisOrigen: findHeaderIndex(["PAIS_ORIGEN", "PAIS ORIGEN"]),
    estadoOrigen: findHeaderIndex(["ESTADO_ORIGEN", "ESTADO ORIGEN"]),
    ciudadOrigen: findHeaderIndex(["CIUDAD_ORIGEN", "CIUDAD ORIGEN"]),
    paisDestino: findHeaderIndex(["PAIS_DESTINO", "PAIS DESTINO"]),
    estadoDestino: findHeaderIndex(["ESTADO_DESTINO", "ESTADO DESTINO"]),
    ciudadDestino: findHeaderIndex(["CIUDAD_DESTINO", "CIUDAD DESTINO"]),
    requerimiento: findHeaderIndex(["REQUERIMIENTO"]),
    mexicoFreight: findHeaderIndex(["MEXICAN_FREIGHT", "MEXICO FREIGHT"]),
    crossing: findHeaderIndex(["CROSSING", "BORDER CROSSING"]),
    usFreight: findHeaderIndex(["US FREIGHT", "USA FREIGHT"]),
    allIn: findHeaderIndex(["ALL IN", "ALL_IN"]),
    monedaAllIn: findHeaderIndex(["MONEDA", "MONEDA ALL IN"]),
    fechaFin: findHeaderIndex(["FECHA_VIGENCIA_FIN", "FECHA VIGENCIA FIN"])
  };

  const match = (value, filtro) => !filtro || normalizeText(value) === normalizeText(filtro);

  let mejor = null;

  values.forEach(row => {
    if (
      (idx.tipoOperacion >= 0 && !match(row[idx.tipoOperacion], filtros.tipoOperacion)) ||
      (idx.tipoViaje >= 0 && !match(row[idx.tipoViaje], filtros.tipoViaje)) ||
      (idx.tipoUnidad >= 0 && !match(row[idx.tipoUnidad], filtros.tipoUnidad)) ||
      (idx.transportista >= 0 && !match(row[idx.transportista], filtros.transportista)) ||
      (idx.cliente >= 0 && !match(row[idx.cliente], filtros.cliente)) ||
      (idx.paisOrigen >= 0 && !match(row[idx.paisOrigen], filtros.paisOrigen)) ||
      (idx.estadoOrigen >= 0 && !match(row[idx.estadoOrigen], filtros.estadoOrigen)) ||
      (idx.ciudadOrigen >= 0 && !match(row[idx.ciudadOrigen], filtros.ciudadOrigen)) ||
      (idx.paisDestino >= 0 && !match(row[idx.paisDestino], filtros.paisDestino)) ||
      (idx.estadoDestino >= 0 && !match(row[idx.estadoDestino], filtros.estadoDestino)) ||
      (idx.ciudadDestino >= 0 && !match(row[idx.ciudadDestino], filtros.ciudadDestino)) ||
      (idx.valorPeso >= 0 && !match(row[idx.valorPeso], filtros.valorPeso))
    ) {
      return;
    }

    const tipoOperacion = idx.tipoOperacion >= 0 ? normalizeText(row[idx.tipoOperacion]) : "";
    if (!tipoOperacion.includes("EXPORT")) {
      return;
    }

    const tarifa = idx.allIn >= 0 ? parseNumber(row[idx.allIn]) : 0;
    const moneda = idx.monedaAllIn >= 0 ? row[idx.monedaAllIn] : "";
    if (!tarifa || !moneda) {
      return;
    }

    if (!mejor || tarifa < mejor.tarifa) {
      mejor = { row, tarifa, moneda };
    }
  });

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const nombreHoja = `COTIZACION_EXPORT_${timestamp}`;
  const shDestino = ss.insertSheet(nombreHoja);

  shDestino.getRange("A1").setValue("Pactra México");
  shDestino.getRange("A2").setValue("Estimado(a) [Nombre del cliente],");
  shDestino.getRange("A3").setValue("Esperando se encuentre muy bien.");
  shDestino.getRange("A5").setValue("En seguimiento a su solicitud, compartimos la cotización correspondiente al servicio de transporte solicitado:");
  shDestino.getRange("A6").setValue("Cotización de Servicio de Transporte – Pactra México (Exportación)");

  if (!mejor) {
    shDestino.getRange("A8").setValue("No se encontró tarifa vigente de exportación con los filtros seleccionados.");
    return;
  }

  const origen = idx.ciudadOrigen >= 0 ? mejor.row[idx.ciudadOrigen] : "";
  const destino = idx.ciudadDestino >= 0 ? mejor.row[idx.ciudadDestino] : "";
  const requerimiento = idx.requerimiento >= 0 ? mejor.row[idx.requerimiento] : "";
  const mexicoFreight = idx.mexicoFreight >= 0 ? mejor.row[idx.mexicoFreight] : "";
  const crossing = idx.crossing >= 0 ? mejor.row[idx.crossing] : "";
  const usFreight = idx.usFreight >= 0 ? mejor.row[idx.usFreight] : "";
  const allIn = idx.allIn >= 0 ? mejor.row[idx.allIn] : "";
  const vigencia = idx.fechaFin >= 0 ? mejor.row[idx.fechaFin] : "";

  shDestino.getRange("A8:G8").setValues([[
    "ORIGEN",
    "DESTINO",
    "REQUERIMIENTO",
    "MEXICAN FREIGHT",
    "CROSSING",
    "US FREIGHT",
    "ALL IN"
  ]]);
  shDestino.getRange("A9:G9").setValues([[
    origen,
    destino,
    requerimiento,
    mexicoFreight,
    crossing,
    usFreight,
    allIn
  ]]);

  shDestino.getRange("A11").setValue("RETORNO");
  shDestino.getRange("A12:G12").setValues([[
    "[Origen Retorno]",
    "[Destino Retorno]",
    "[Requerimiento]",
    "[Mex Freight]",
    "[Crossing]",
    "[US Freight]",
    "[All In]"
  ]]);

  shDestino.getRange("A14").setValue("Condiciones Comerciales:");
  shDestino.getRange("A15").setValue(`• Las tarifas están expresadas en ${mejor.moneda || "[USD/MXN]"} y no incluyen IVA.`);
  shDestino.getRange("A16").setValue("• La tarifa incluye el costo del transporte en unidad tipo [especificar].");
  shDestino.getRange("A17").setValue("• Tiempo libre de maniobras: 3 horas en carga y 3 horas en descarga. Después de ese tiempo aplicará cargo por demora según tarifa vigente.");
  shDestino.getRange("A18").setValue("• El seguro de mercancía no está incluido. Puede cotizarse adicional bajo solicitud expresa del cliente.");
  shDestino.getRange("A19").setValue("• Tarifas sujetas a disponibilidad de unidad y condiciones de ruta al momento de confirmar servicio.");
  shDestino.getRange("A20").setValue(`• Vigencia de la cotización: ${vigencia || "[X días]"}.`);
  shDestino.getRange("A21").setValue("• No incluye maniobras especiales, custodia, almacenaje ni costos de cruce fronterizo salvo se indique lo contrario.");
  shDestino.getRange("A22").setValue("• Los tiempos de tránsito son estimados y pueden variar por factores externos (clima, tráfico, inspecciones, etc.).");
  shDestino.getRange("A23").setValue("Quedamos atentos a sus comentarios o cualquier ajuste que requiera.");

  shDestino.getRange("A1").setFontWeight("bold");
  shDestino.getRange("A6").setFontWeight("bold");
  shDestino.getRange("A8:G8").setFontWeight("bold").setBackground("#1F4E79").setFontColor("#FFFFFF");
  shDestino.getRange("A9:G9").setBorder(true, true, true, true, true, true);
  shDestino.autoResizeColumns(1, 7);
}

function generarCotizacionExportacionCliente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTarifario = ss.getSheetByName("Tarifario estandar");
  const shReporte = ss.getSheetByName("Reporte Final");

  if (!shTarifario || !shReporte) {
    throw new Error("No se encontró la hoja 'Tarifario estandar' o 'Reporte Final'.");
  }

  const filtros = {
    tipoOperacion: shReporte.getRange("C3").getValue(),
    tipoViaje: shReporte.getRange("C4").getValue(),
    tipoUnidad: shReporte.getRange("C5").getValue(),
    transportista: shReporte.getRange("C6").getValue(),
    cliente: shReporte.getRange("C7").getValue(),
    paisOrigen: shReporte.getRange("C8").getValue(),
    estadoOrigen: shReporte.getRange("C9").getValue(),
    ciudadOrigen: shReporte.getRange("C10").getValue(),
    paisDestino: shReporte.getRange("C11").getValue(),
    estadoDestino: shReporte.getRange("C12").getValue(),
    ciudadDestino: shReporte.getRange("C13").getValue(),
    valorPeso: shReporte.getRange("C14").getValue()
  };

  const dataRange = shTarifario.getDataRange();
  const values = dataRange.getValues();
  const headers = values.shift();
  const normalizedHeaders = headers.map(header => normalizeText(header));

  const findHeaderIndex = aliases => {
    const normalizedAliases = aliases.map(alias => normalizeText(alias));
    return normalizedHeaders.findIndex(header => normalizedAliases.includes(header));
  };

  const idx = {
    tipoOperacion: findHeaderIndex(["TIPO_DE_OPERACION", "TIPO DE OPERACION"]),
    tipoViaje: findHeaderIndex(["TIPO_DE_VIAJE", "TIPO DE VIAJE"]),
    tipoUnidad: findHeaderIndex(["TIPO_UNIDAD", "TIPO UNIDAD"]),
    transportista: findHeaderIndex(["TRANSPORTISTA"]),
    cliente: findHeaderIndex(["CLIENTE"]),
    paisOrigen: findHeaderIndex(["PAIS_ORIGEN", "PAIS ORIGEN"]),
    estadoOrigen: findHeaderIndex(["ESTADO_ORIGEN", "ESTADO ORIGEN"]),
    ciudadOrigen: findHeaderIndex(["CIUDAD_ORIGEN", "CIUDAD ORIGEN"]),
    paisDestino: findHeaderIndex(["PAIS_DESTINO", "PAIS DESTINO"]),
    estadoDestino: findHeaderIndex(["ESTADO_DESTINO", "ESTADO DESTINO"]),
    ciudadDestino: findHeaderIndex(["CIUDAD_DESTINO", "CIUDAD DESTINO"]),
    requerimiento: findHeaderIndex(["REQUERIMIENTO"]),
    mexicoFreight: findHeaderIndex(["MEXICAN_FREIGHT", "MEXICO FREIGHT"]),
    crossing: findHeaderIndex(["CROSSING", "BORDER CROSSING"]),
    usFreight: findHeaderIndex(["US FREIGHT", "USA FREIGHT"]),
    allIn: findHeaderIndex(["ALL IN", "ALL_IN"]),
    monedaAllIn: findHeaderIndex(["MONEDA", "MONEDA ALL IN"]),
    fechaFin: findHeaderIndex(["FECHA_VIGENCIA_FIN", "FECHA VIGENCIA FIN"])
  };

  const match = (value, filtro) => !filtro || normalizeText(value) === normalizeText(filtro);

  let mejor = null;

  values.forEach(row => {
    if (
      (idx.tipoOperacion >= 0 && !match(row[idx.tipoOperacion], filtros.tipoOperacion)) ||
      (idx.tipoViaje >= 0 && !match(row[idx.tipoViaje], filtros.tipoViaje)) ||
      (idx.tipoUnidad >= 0 && !match(row[idx.tipoUnidad], filtros.tipoUnidad)) ||
      (idx.transportista >= 0 && !match(row[idx.transportista], filtros.transportista)) ||
      (idx.cliente >= 0 && !match(row[idx.cliente], filtros.cliente)) ||
      (idx.paisOrigen >= 0 && !match(row[idx.paisOrigen], filtros.paisOrigen)) ||
      (idx.estadoOrigen >= 0 && !match(row[idx.estadoOrigen], filtros.estadoOrigen)) ||
      (idx.ciudadOrigen >= 0 && !match(row[idx.ciudadOrigen], filtros.ciudadOrigen)) ||
      (idx.paisDestino >= 0 && !match(row[idx.paisDestino], filtros.paisDestino)) ||
      (idx.estadoDestino >= 0 && !match(row[idx.estadoDestino], filtros.estadoDestino)) ||
      (idx.ciudadDestino >= 0 && !match(row[idx.ciudadDestino], filtros.ciudadDestino)) ||
      (idx.valorPeso >= 0 && !match(row[idx.valorPeso], filtros.valorPeso))
    ) {
      return;
    }

    const tipoOperacion = idx.tipoOperacion >= 0 ? normalizeText(row[idx.tipoOperacion]) : "";
    if (!tipoOperacion.includes("EXPORT")) {
      return;
    }

    const tarifa = idx.allIn >= 0 ? parseNumber(row[idx.allIn]) : 0;
    const moneda = idx.monedaAllIn >= 0 ? row[idx.monedaAllIn] : "";
    if (!tarifa || !moneda) {
      return;
    }

    if (!mejor || tarifa < mejor.tarifa) {
      mejor = { row, tarifa, moneda };
    }
  });

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const nombreHoja = `COTIZACION_EXPORT_CLIENTE_${timestamp}`;
  const shDestino = ss.insertSheet(nombreHoja);

  shDestino.getRange("A1").setValue("Pactra México");
  shDestino.getRange("A3").setValue("Estimado(a) [Nombre del cliente],");
  shDestino.getRange("A4").setValue("Esperando se encuentre muy bien.");
  shDestino.getRange("A6").setValue("En seguimiento a su solicitud, compartimos la cotización correspondiente al servicio de transporte solicitado:");
  shDestino.getRange("A7").setValue("Cotización de Servicio de Transporte – Pactra México (Exportación)");

  if (!mejor) {
    shDestino.getRange("A9").setValue("No se encontró tarifa vigente de exportación con los filtros seleccionados.");
    return;
  }

  const origen = idx.ciudadOrigen >= 0 ? mejor.row[idx.ciudadOrigen] : "";
  const destino = idx.ciudadDestino >= 0 ? mejor.row[idx.ciudadDestino] : "";
  const requerimiento = idx.requerimiento >= 0 ? mejor.row[idx.requerimiento] : "";
  const mexicoFreight = idx.mexicoFreight >= 0 ? mejor.row[idx.mexicoFreight] : "";
  const crossing = idx.crossing >= 0 ? mejor.row[idx.crossing] : "";
  const usFreight = idx.usFreight >= 0 ? mejor.row[idx.usFreight] : "";
  const allIn = idx.allIn >= 0 ? mejor.row[idx.allIn] : "";
  const vigencia = idx.fechaFin >= 0 ? mejor.row[idx.fechaFin] : "";

  shDestino.getRange("A9:G9").setValues([[
    "ORIGEN",
    "DESTINO",
    "REQUERIMIENTO",
    "MEXICAN FREIGHT",
    "CROSSING",
    "US FREIGHT",
    "ALL IN"
  ]]);
  shDestino.getRange("A10:G10").setValues([[
    origen,
    destino,
    requerimiento,
    mexicoFreight,
    crossing,
    usFreight,
    allIn
  ]]);

  shDestino.getRange("A12").setValue("RETORNO");
  shDestino.getRange("A13:G13").setValues([[
    "[Origen Retorno]",
    "[Destino Retorno]",
    "[Requerimiento]",
    "[Mex Freight]",
    "[Crossing]",
    "[US Freight]",
    "[All In]"
  ]]);

  shDestino.getRange("A15").setValue("Condiciones Comerciales:");
  shDestino.getRange("A16").setValue(`• Las tarifas están expresadas en ${mejor.moneda || "[USD/MXN]"} y no incluyen IVA.`);
  shDestino.getRange("A17").setValue("• La tarifa incluye el costo del transporte en unidad tipo [especificar].");
  shDestino.getRange("A18").setValue("• Tiempo libre de maniobras: 3 horas en carga y 3 horas en descarga. Después de ese tiempo aplicará cargo por demora según tarifa vigente.");
  shDestino.getRange("A19").setValue("• El seguro de mercancía no está incluido. Puede cotizarse adicional bajo solicitud expresa del cliente.");
  shDestino.getRange("A20").setValue("• Tarifas sujetas a disponibilidad de unidad y condiciones de ruta al momento de confirmar servicio.");
  shDestino.getRange("A21").setValue(`• Vigencia de la cotización: ${vigencia || "[X días]"}.`);
  shDestino.getRange("A22").setValue("• No incluye maniobras especiales, custodia, almacenaje ni costos de cruce fronterizo salvo se indique lo contrario.");
  shDestino.getRange("A23").setValue("• Los tiempos de tránsito son estimados y pueden variar por factores externos (clima, tráfico, inspecciones, etc.).");
  shDestino.getRange("A24").setValue("Quedamos atentos a sus comentarios o cualquier ajuste que requiera.");

  shDestino.getRange("A1").setFontWeight("bold");
  shDestino.getRange("A7").setFontWeight("bold");
  shDestino.getRange("A9:G9").setFontWeight("bold").setBackground("#1F4E79").setFontColor("#FFFFFF");
  shDestino.getRange("A10:G10").setBorder(true, true, true, true, true, true);
  shDestino.autoResizeColumns(1, 7);
}

function setupSendTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(trigger => trigger.getHandlerFunction() === "onEditSendTrigger");
  if (!exists) {
    ScriptApp.newTrigger("onEditSendTrigger")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
}

function onEditSendTrigger(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (!sheetName.startsWith("COTIZACION_EXPORT_CLIENTE_")) return;
  if (e.range.getA1Notation() !== "J2") return;
  if (!e.range.getValue()) return;

  e.range.setValue(false);
  sendQuoteEmail();
}

// ============================
// ENVÍO DE COTIZACIÓN (CLIENTE)
// ============================
function openSendSidebar() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 16px;">
      <h2 style="margin-top: 0; color: #1F4E79;">Enviar Cotización</h2>
      <p>Se enviará exactamente la información mostrada en esta hoja.</p>
      <button onclick="enviar()" style="width: 100%; padding: 10px; background: #1F4E79; color: #fff; border: 0; border-radius: 4px;">
        Enviar cotización
      </button>
      <div id="estado" style="margin-top: 12px; font-size: 12px; color: #555;"></div>
    </div>
    <script>
      function enviar() {
        document.getElementById('estado').textContent = 'Enviando...';
        google.script.run
          .withSuccessHandler(res => {
            document.getElementById('estado').textContent = res;
          })
          .withFailureHandler(err => {
            document.getElementById('estado').textContent = 'Error: ' + err.message;
          })
          .sendQuoteEmail();
      }
    </script>
  `).setTitle("Enviar Cotización");

  SpreadsheetApp.getUi().showSidebar(html);
}

function sendQuoteEmail() {
  const data = getActiveQuoteData();
  if (!data.ok) {
    throw new Error(data.mensaje);
  }

  const htmlBody = buildEmailHTML(data.payload);
  const subject = data.payload.subject || "Cotización de Transporte – Pactra México";
  const plainBody = "Se adjunta la cotización en formato HTML.";

  MailApp.sendEmail({
    to: data.payload.to,
    cc: data.payload.cc || "",
    subject,
    htmlBody,
    body: plainBody
  });

  return "Cotización enviada con éxito.";
}

function getActiveQuoteData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (!sh) {
    return { ok: false, mensaje: "No se encontró la hoja activa." };
  }

  const to = sh.getRange("B2").getDisplayValue();
  if (!to) {
    return { ok: false, mensaje: "Falta el correo destino en B2 de la hoja activa." };
  }

  const values = sh.getDataRange().getDisplayValues();
  const normalized = values.map(row => row.map(cell => normalizeText(cell)));

  const requiredHeaders = [
    "ORIGEN",
    "DESTINO",
    "REQUERIMIENTO",
    "MEXICAN FREIGHT",
    "CROSSING",
    "US FREIGHT",
    "ALL IN",
    "MONEDA",
    "VIGENCIA"
  ];

  const headerRowIndex = normalized.findIndex(row =>
    requiredHeaders.every(header => row.includes(header))
  );

  if (headerRowIndex < 0) {
    return { ok: false, mensaje: "No se encontró una fila de encabezados válida en la hoja activa." };
  }

  const headers = values[headerRowIndex];
  const normalizedHeaders = normalized[headerRowIndex];

  const findIndex = header => normalizedHeaders.indexOf(header);

  const indexes = {
    origen: findIndex("ORIGEN"),
    destino: findIndex("DESTINO"),
    requerimiento: findIndex("REQUERIMIENTO"),
    mexicanFreight: findIndex("MEXICAN FREIGHT"),
    crossing: findIndex("CROSSING"),
    usFreight: findIndex("US FREIGHT"),
    allIn: findIndex("ALL IN"),
    moneda: findIndex("MONEDA"),
    vigencia: findIndex("VIGENCIA")
  };

  if (Object.values(indexes).some(idx => idx < 0)) {
    return { ok: false, mensaje: "Faltan columnas obligatorias en la cotización." };
  }

  const dataRows = values
    .slice(headerRowIndex + 1)
    .filter(row => row.some(cell => String(cell || "").trim() !== ""));

  if (dataRows.length === 0) {
    return { ok: false, mensaje: "No hay filas de cotización debajo del encabezado." };
  }

  const tableRows = dataRows.map(row => ({
    origen: row[indexes.origen],
    destino: row[indexes.destino],
    requerimiento: row[indexes.requerimiento],
    mexicanFreight: row[indexes.mexicanFreight],
    crossing: row[indexes.crossing],
    usFreight: row[indexes.usFreight],
    allIn: row[indexes.allIn],
    moneda: row[indexes.moneda],
    vigencia: row[indexes.vigencia]
  }));

  return {
    ok: true,
    payload: {
      to,
      cc: sh.getRange("B3").getDisplayValue(),
      subject: sh.getRange("B4").getDisplayValue(),
      headers,
      tableRows,
      clientName: sh.getRange("B5").getDisplayValue() || "[Nombre del cliente]"
    }
  };
}

function buildEmailHTML(data) {
  const tableHeaders = [
    "ORIGEN",
    "DESTINO",
    "REQUERIMIENTO",
    "MEXICAN FREIGHT",
    "CROSSING",
    "US FREIGHT",
    "ALL IN",
    "MONEDA",
    "VIGENCIA"
  ];

  const headerHtml = tableHeaders
    .map(header => `<th style="background:#1F4E79;color:#FFFFFF;padding:8px;border:1px solid #D0D0D0;">${header}</th>`)
    .join("");

  const rowsHtml = data.tableRows.map(row => `
    <tr>
      <td style="padding:8px;border:1px solid #D0D0D0;">${row.origen}</td>
      <td style="padding:8px;border:1px solid #D0D0D0;">${row.destino}</td>
      <td style="padding:8px;border:1px solid #D0D0D0;">${row.requerimiento}</td>
      <td style="padding:8px;border:1px solid #D0D0D0;text-align:right;">${row.mexicanFreight}</td>
      <td style="padding:8px;border:1px solid #D0D0D0;text-align:right;">${row.crossing}</td>
      <td style="padding:8px;border:1px solid #D0D0D0;text-align:right;">${row.usFreight}</td>
      <td style="padding:8px;border:1px solid #D0D0D0;text-align:right;">${row.allIn}</td>
      <td style="padding:8px;border:1px solid #D0D0D0;text-align:center;">${row.moneda}</td>
      <td style="padding:8px;border:1px solid #D0D0D0;text-align:center;">${row.vigencia}</td>
    </tr>
  `).join("");

  return `
    <div style="font-family:Arial,sans-serif;font-size:12px;color:#333;">
      <h2 style="margin:0 0 8px 0;color:#1F4E79;">Pactra México</h2>
      <p>Estimado(a) ${data.clientName},</p>
      <p>Esperando se encuentre muy bien.</p>
      <p>En seguimiento a su solicitud, compartimos la cotización correspondiente al servicio de transporte solicitado:</p>
      <table style="border-collapse:collapse;width:100%;margin:12px 0;">
        <tr>${headerHtml}</tr>
        ${rowsHtml}
      </table>
      <p>Quedamos atentos a cualquier comentario o ajuste que requiera.</p>
    </div>
  `;
}
