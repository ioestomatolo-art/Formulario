document.addEventListener("DOMContentLoaded", () => {
  let productosExcel = [];       // productos de la hoja activa
  let productosAgregados = [];   // claves ya usadas en filas
  let wbGlobal = null;           // workbook cargado
  let hojaActiva = null;         // nombre de la hoja activa

  // =================== Elementos del DOM ===================
  const page1 = document.getElementById("page1");
  const page2 = document.getElementById("page2");
  const page3 = document.getElementById("page3");

  const selectCategoria = document.getElementById("categoria");
  const btnSiguiente = document.getElementById("btnSiguiente");
  const btnRegresar1 = document.getElementById("btnRegresar1");
  const btnRegresar2 = document.getElementById("btnRegresar2");

  const inputExcel = document.getElementById("inputExcel");
  const btnAgregarFila = document.getElementById("btnAgregarFila");
  const tbody = document.querySelector("#tablaInsumos tbody");
  const btnEnviarInsumos = document.getElementById("btnEnviarInsumos");

  // =================== Mapa de categorías → hojas Excel ===================
  const mapaCategoriaHoja = {
    "insumos": "MEDICAMENTOS Y RAYOS X",
    "material": "MATERIAL",
    "equipo": "EQUIPO MEDICO",
    "mobiliario": "MOBILIARIO",
    "informatico": "BIENES INFORMATICOS"
  };

  // =================== Navegación ===================
  btnSiguiente.addEventListener("click", () => {
    const cat = selectCategoria.value;
    const hojaBuscada = mapaCategoriaHoja[cat];

    if (!hojaBuscada) {
      alert("Selecciona una categoría válida.");
      return;
    }

    // Verifica que la hoja exista en el Excel si se requiere
    if (wbGlobal && !wbGlobal.SheetNames.includes(hojaBuscada)) {
      alert(`No se encontró la hoja "${hojaBuscada}" en el Excel.`);
      return;
    }

    hojaActiva = hojaBuscada;

    if (cat === "insumos" || cat === "material") {
      page1.classList.remove("activo");
      page1.classList.add("oculto");
      page2.classList.add("activo");
      if (wbGlobal) parseHoja(hojaActiva);
    } else {
      page1.classList.remove("activo");
      page1.classList.add("oculto");
      page3.classList.add("activo");
      if (wbGlobal) parseHoja(hojaActiva);
    }
  });

  btnRegresar1.addEventListener("click", () => {
    page2.classList.remove("activo");
    page2.classList.add("oculto");
    page1.classList.add("activo");
  });

  btnRegresar2.addEventListener("click", () => {
    page3.classList.remove("activo");
    page3.classList.add("oculto");
    page1.classList.add("activo");
  });

  // =================== Carga de Excel ===================
  inputExcel.addEventListener("change", (evt) => {
    const file = evt.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        wbGlobal = XLSX.read(data, { type: "array" });

        // Crear selector de hojas si no existe
        ensureSheetSelector(wbGlobal);

        // Detectar hoja automáticamente
        const preferida = "MEDICAMENTOS Y RAYOS X";
        const auto = autoDetectHoja(wbGlobal);
        hojaActiva = wbGlobal.SheetNames.includes(preferida) ? preferida : auto || wbGlobal.SheetNames[0];

        document.getElementById("selectHoja").value = hojaActiva;

        parseHoja(hojaActiva);
        alert(`Archivo cargado. Hoja activa: "${hojaActiva}".`);
      } catch (err) {
        console.error(err);
        alert("Error al leer el archivo Excel.");
      }
    };
    reader.readAsArrayBuffer(file);
  });

  function ensureSheetSelector(wb) {
    let sel = document.getElementById("selectHoja");
    if (!sel) {
      sel = document.createElement("select");
      sel.id = "selectHoja";
      sel.style.marginLeft = "10px";
      sel.title = "Selecciona la hoja del Excel a usar";
      inputExcel.insertAdjacentElement("afterend", sel);

      sel.addEventListener("change", () => {
        hojaActiva = sel.value;
        parseHoja(hojaActiva);
        alert(`Hoja cambiada a: "${hojaActiva}".`);
      });
    }
    sel.innerHTML = "";
    wb.SheetNames.forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      sel.appendChild(opt);
    });
  }

  function autoDetectHoja(wb) {
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      if (!rows || rows.length === 0) continue;
      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const rowUp = rows[i].map((c) => String(c).trim().toUpperCase());
        if (rowUp.some((c) => c.includes("CLAVE")) && rowUp.some((c) => c.includes("DESCRIP"))) {
          return name;
        }
      }
    }
    return null;
  }

  // =================== Parsea hoja ===================
  function parseHoja(sheetName) {
    const sheet = wbGlobal.Sheets[sheetName];
    if (!sheet) {
      alert(`No se pudo leer la hoja "${sheetName}".`);
      return;
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    let headerIndex = -1;
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = rows[i].map((c) => String(c).trim().toUpperCase());
      if (row.some((c) => c.includes("CLAVE")) && row.some((c) => c.includes("DESCRIP"))) {
        headerIndex = i;
        break;
      }
    }
    if (headerIndex === -1) {
      alert('No se encontró fila de encabezados con "CLAVE" y "DESCRIPCIÓN" en esta hoja.');
      return;
    }

    const headers = rows[headerIndex].map((h) => String(h).trim());
    const idxClave = headers.findIndex((h) => /CLAVE/i.test(h));
    const idxDesc  = headers.findIndex((h) => /DESCRIPCION|DESCRIPCIÓN|DESCRIPC|DESC/i.test(h));
    const idxStock = headers.findIndex((h) => /STOCK|EXISTEN|EXISTENCIA/i.test(h));
    const idxMin   = headers.findIndex((h) => /MINIMO|MÍNIMO|MINIMUM/i.test(h));
    const idxCad   = headers.findIndex((h) => /CADUCIDAD|CADUC/i.test(h));

    if (idxClave === -1 || idxDesc === -1) {
      alert('No se encontraron columnas "CLAVE" o "DESCRIPCIÓN" en esta hoja.');
      return;
    }

    const dataRows = rows.slice(headerIndex + 1);
    productosExcel = dataRows
      .map((r) => ({
        clave: String(r[idxClave] ?? "").trim(),
        descripcion: String(r[idxDesc] ?? "").trim(),
        stock: r[idxStock] ?? "",
        minimo: r[idxMin] ?? "",
        caducidadRaw: r[idxCad] ?? "",
      }))
      .filter((p) => p.clave && p.clave.toUpperCase() !== "CLAVE");

    limpiarTabla();
    agregarFila(); // fila inicial vacía
  }

  // =================== Agregar fila ===================
  btnAgregarFila.addEventListener("click", () => {
    if (productosExcel.length === 0) {
      alert("Primero debes cargar un archivo Excel y seleccionar una hoja válida.");
      return;
    }
    agregarFila();
  });

  // =================== Enviar datos ===================
  btnEnviarInsumos.addEventListener("click", () => {
    const datos = [];
    for (const row of tbody.rows) {
      const select = row.cells[1].querySelector("select");
      if (!select.value) continue;
      const desc = row.cells[2].querySelector("input").value;
      const stock = row.cells[3].querySelector("input").value;
      const minimo = row.cells[4].querySelector("input").value;
      const cad = row.cells[6].querySelector("input").value;
      const diasRest = row.cells[7].querySelector("input").value;

      datos.push({
        hoja: hojaActiva || "",
        clave: select.value,
        descripcion: desc,
        stock: Number(stock),
        minimo: Number(minimo),
        caducidad: cad,
        diasRestantes: diasRest === "Caducado" ? diasRest : (diasRest === "" ? null : Number(diasRest)),
      });
    }
    console.log("Datos a enviar:", datos);
    alert("Datos listos para enviar (ver consola).");
    window.location.reload();
  });

  // =================== Funciones auxiliares ===================
  function limpiarTabla() {
    tbody.innerHTML = "";
    productosAgregados = [];
  }

  function agregarFila() {
    const rowIndex = tbody.rows.length + 1;
    const row = tbody.insertRow();

    // No.
    const cNo = row.insertCell();
    cNo.textContent = rowIndex;

    // Clave
    const cClave = row.insertCell();
    const select = document.createElement("select");
    select.innerHTML = `<option value="">Seleccione</option>`;
    cClave.appendChild(select);

    // Descripción
    const cDesc = row.insertCell();
    const inDesc = document.createElement("input");
    inDesc.type = "text";
    inDesc.readOnly = true;
    cDesc.appendChild(inDesc);

    // Stock
    const cStock = row.insertCell();
    const inStock = document.createElement("input");
    inStock.type = "number";
    inStock.min = "0";
    inStock.value = "";
    cStock.appendChild(inStock);

    // Mínimo
    const cMin = row.insertCell();
    const inMin = document.createElement("input");
    inMin.type = "number";
    inMin.min = "0";
    inMin.value = "";
    cMin.appendChild(inMin);

    // Estado
    const cEstado = row.insertCell();
    const spanEstado = document.createElement("span");
    cEstado.appendChild(spanEstado);

    // Caducidad
    const cCad = row.insertCell();
    const inCad = document.createElement("input");
    inCad.type = "date";
    cCad.appendChild(inCad);

    // Días restantes
    const cDias = row.insertCell();
    const inDias = document.createElement("input");
    inDias.type = "text";
    inDias.readOnly = true;
    cDias.appendChild(inDias);

    select.dataset.claveAnterior = "";
    actualizarSelects();

    select.addEventListener("change", function () {
      const claveAnterior = this.dataset.claveAnterior || "";
      const claveNueva = this.value || "";

      if (claveAnterior && claveAnterior !== claveNueva) {
        const idx = productosAgregados.indexOf(claveAnterior);
        if (idx !== -1) productosAgregados.splice(idx, 1);
      }

      if (!claveNueva) {
        this.dataset.claveAnterior = "";
        inDesc.value = "";
        inStock.value = "";
        inMin.value = "";
        inCad.value = "";
        inDias.value = "";
        spanEstado.textContent = "";
        row.classList.remove("expired", "warning-expiry", "valid-expiry");
        actualizarSelects();
        return;
      }

      if (productosAgregados.includes(claveNueva) && claveNueva !== claveAnterior) {
        alert("Esta clave ya fue agregada en otra fila.");
        this.value = claveAnterior || "";
        return;
      }

      if (!productosAgregados.includes(claveNueva)) productosAgregados.push(claveNueva);
      this.dataset.claveAnterior = claveNueva;

      const producto = productosExcel.find((p) => String(p.clave) === String(claveNueva));
      if (producto) {
        inDesc.value = producto.descripcion || "";
        inStock.value = producto.stock ?? "";
        inMin.value = producto.minimo ?? "";
        inCad.value = formatoFechaISO(producto.caducidadRaw) || "";
      }

      actualizarFila(row);
      actualizarSelects();
    });

    // Inputs
    inStock.addEventListener("input", () => {
      if (inStock.value < 0) inStock.value = 0;
      actualizarFila(row);
    });
    inMin.addEventListener("input", () => {
      if (inMin.value < 0) inMin.value = 0;
      actualizarFila(row);
    });
    inCad.addEventListener("change", () => actualizarFila(row));
  }

  function actualizarFila(row) {
    const inStock = row.cells[3].querySelector("input");
    const inMin = row.cells[4].querySelector("input");
    const inCad = row.cells[6].querySelector("input");
    const inDias = row.cells[7].querySelector("input");
    const estadoSpan = row.cells[5].querySelector("span");

    const stock = inStock.value !== "" ? parseFloat(inStock.value) : 0;
    const minimo = inMin.value !== "" ? parseFloat(inMin.value) : 0;

    let estadoTextStock = "";
    if (stock > minimo) estadoTextStock = "Stock suficiente";
    else if (stock < minimo) estadoTextStock = "Bajo stock";
    else estadoTextStock = "Stock justo";

    const dias = calcularDiasRestantes(inCad.value);
    inDias.value = dias === null ? "" : (dias < 0 ? "Caducado" : dias);

    let estadoTextCad = "";
    row.classList.remove("expired", "warning-expiry", "valid-expiry");
    if (dias !== null) {
      if (dias < 0) {
        row.classList.add("expired");
        estadoTextCad = "Caducado";
      } else if (dias < 180) {
        row.classList.add("expired");
        estadoTextCad = "Próximo a caducar (<6 meses)";
      } else if (dias < 365) {
        row.classList.add("warning-expiry");
        estadoTextCad = "Caduca en 6–12 meses";
      } else {
        row.classList.add("valid-expiry");
        estadoTextCad = "Vigente (>12 meses)";
      }
    }

    const partes = [];
    if (estadoTextCad) partes.push(estadoTextCad);
    if (estadoTextStock) partes.push(estadoTextStock);
    estadoSpan.textContent = partes.join(" | ");
  }

  function actualizarSelects() {
    const selects = document.querySelectorAll("#tablaInsumos select");
    selects.forEach((select) => {
      const valorActual = select.value || "";
      select.innerHTML = `<option value="">Seleccione</option>`;
      productosExcel.forEach((p) => {
        const clave = String(p.clave);
        if (!productosAgregados.includes(clave) || clave === valorActual) {
          const opt = document.createElement("option");
          opt.value = clave;
          opt.textContent = clave;
          if (clave === valorActual) opt.selected = true;
          select.appendChild(opt);
        }
      });
    });
  }

  function calcularDiasRestantes(fechaIso) {
    if (!fechaIso) return null;
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const f = new Date(fechaIso);
    if (isNaN(f)) return null;
    return Math.ceil((f - hoy) / (1000 * 60 * 60 * 24));
  }

  function formatoFechaISO(valor) {
    if (valor === null || valor === undefined || valor === "") return "";
    if (typeof valor === "number") {
      const d = XLSX.SSF.parse_date_code(valor);
      if (!d) return "";
      return new Date(d.y, d.m - 1, d.d).toISOString().split("T")[0];
    }
    const date = new Date(valor);
    return isNaN(date) ? "" : date.toISOString().split("T")[0];
  }
});
