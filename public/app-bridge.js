(function () {
  if (window.electronExcel) return;

  // Si Tauri está disponible (desktop), usar el bridge original
  const tauriCore = window.__TAURI__ && window.__TAURI__.core;
  if (tauriCore && typeof tauriCore.invoke === "function") {
    const invoke = tauriCore.invoke;
    window.electronExcel = {
      selectFile: () => invoke("excel_select_file"),
      getSelectedFile: () => invoke("excel_get_selected_file"),
      saveAlumnos: (alumnos) => invoke("excel_save_alumnos", { alumnos }),
      getUnidades: () => invoke("excel_get_unidades"),
      saveUnidades: (unidades) => invoke("excel_save_unidades", { unidades }),
      getRraaCriterios: () => invoke("excel_get_rraa_criterios"),
      saveRraaCriterios: (payloadOrRraa, criterios, ponderacionesUnidad = []) => {
        const payload = Array.isArray(payloadOrRraa)
          ? { rraa: payloadOrRraa, criterios, ponderacionesUnidad }
          : payloadOrRraa;
        return invoke("excel_save_rraa_criterios", { payload });
      },
      getNotasActividad: (payload) => invoke("excel_get_notas_actividad", { payload }),
      getNotasActividadesTipo: (payload) => invoke("excel_get_notas_actividades_tipo", { payload }),
      saveNotasActividad: (payload) => invoke("excel_save_notas_actividad", { payload }),
      saveCeNotas: (payload) => invoke("excel_save_ce_notas", { payload }),
      addActividad: (payload) => invoke("excel_add_actividad", { payload }),
      getNotasUnidad: (payload) => invoke("excel_get_notas_unidad", { payload }),
      getNotasEvaluacion: (payload) => invoke("excel_get_notas_evaluacion", { payload }),
      getNotasEvaluacionAlumno: (payload) => invoke("excel_get_notas_evaluacion_alumno", { payload }),
      getAlumnosInformes: () => invoke("excel_get_alumnos_informes"),
      setSelectedFile: (filePath) => invoke("excel_set_selected_file", { filePath }),
      verifyFileExists: (filePath) => invoke("excel_verify_file_exists", { filePath }),
      openExternal: (url) => invoke("app_open_external", { url }),
      getDiarioData: () => invoke("excel_get_diario"),
      saveDiarioEntrada: (payload) => invoke("excel_save_diario_entrada", { payload }),
      deleteDiarioEntrada: (payload) => invoke("excel_delete_diario_entrada", { payload }),
    };
    return;
  }

  // ── Bridge Android/Web ──────────────────────────────────────────────────────

  const FILE_KEY = "android_excel_file_name";
  const DATA_KEY = "android_excel_data"; // base64 del xlsx en memoria

  let _workbook = null;
  let _fileName = localStorage.getItem(FILE_KEY) || null;

  // Carga el workbook desde localStorage si existe
  function _loadFromStorage() {
    const b64 = localStorage.getItem(DATA_KEY);
    if (!b64) return false;
    try {
      // atob devuelve una cadena binaria — convertir a Uint8Array sin loop
      const bin = atob(b64);
      const buf = Uint8Array.from(bin, c => c.charCodeAt(0));
      _workbook = XLSX.read(buf, { type: "array", cellDates: true });
      return true;
    } catch { return false; }
  }

  function _saveToStorage() {
    // Deshabilitado: XLSX.write es síncrono y bloquea el hilo principal en Android WebView
    // El workbook se mantiene en _workbook durante la sesión
  }

  // Descarga el xlsx actual al dispositivo
  function _downloadWorkbook() {
    if (!_workbook || !_fileName) return;
    const wbout = XLSX.write(_workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = _fileName;
    a.click();
    URL.revokeObjectURL(url);
  }

  function _wb() {
    if (_workbook) return _workbook;
    _loadFromStorage();
    return _workbook;
  }

  function _sheet(name) {
    const wb = _wb();
    if (!wb) throw new Error("No hay archivo Excel cargado.");
    if (!wb.SheetNames.includes(name)) throw new Error(`Hoja "${name}" no encontrada.`);
    return wb.Sheets[name];
  }

  function _sheetToJson(name, opts = {}) {
    return XLSX.utils.sheet_to_json(_sheet(name), { defval: "", ...opts });
  }

  // Construye el resultado completo que esperan todos los HTML
  function _buildResult() {
    const wb = _workbook;
    const result = { filePath: _fileName, fileName: _fileName };

    // alumnos (gestor-alumnos.html)
    try {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets["DATOS"] || wb.Sheets[wb.SheetNames[0]], { defval: "" });
      result.alumnos = rows;
    } catch { result.alumnos = []; }

    // rraa + criterios (gestor-rraa-criterios.html)
    try {
      result.rraa = XLSX.utils.sheet_to_json(wb.Sheets["RRAA"] || {}, { defval: "" });
    } catch { result.rraa = []; }
    try {
      result.criterios = XLSX.utils.sheet_to_json(wb.Sheets["Criterios"] || {}, { defval: "" });
    } catch { result.criterios = []; }
    result.ponderacionesUnidad = [];

    return result;
  }

  // Abre el selector de archivos nativo de Android
  function _openFilePicker() {
    return new Promise((resolve) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".xlsx,.xls";
      input.style.cssText = "position:fixed;top:-9999px;left:-9999px;width:1px;height:1px;opacity:0;";
      document.body.appendChild(input);

      const cleanup = () => {
        if (input.parentNode) input.parentNode.removeChild(input);
      };

      let resolved = false;
      let focusTimer = null;

      const onFocus = () => {
        window.removeEventListener("focus", onFocus);
        focusTimer = setTimeout(() => {
          if (!resolved) {
            resolved = true;
            cleanup();
            resolve(null);
          }
        }, 1000);
      };

      input.addEventListener("change", async () => {
        if (resolved) return;
        resolved = true;
        window.removeEventListener("focus", onFocus);
        if (focusTimer) clearTimeout(focusTimer);
        const file = input.files && input.files[0];
        cleanup();
        if (!file) { resolve(null); return; }
        try {
          const buf = await file.arrayBuffer();
          _workbook = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
          _fileName = file.name;
          localStorage.setItem(FILE_KEY, _fileName);
          const result = _buildResult();
          resolve(result);
          setTimeout(() => _saveToStorage(), 0);
        } catch(e) {
          resolve(null);
        }
      });

      input.click();

      setTimeout(() => {
        window.addEventListener("focus", onFocus);
      }, 800);
    });
  }

  // ── Helpers para leer datos de las hojas ────────────────────────────────────

  function _getAlumnos() {
    const rows = _sheetToJson("Alumnos");
    return rows.map(r => ({
      nombre: r["Nombre"] || r["nombre"] || "",
      apellidos: r["Apellidos"] || r["apellidos"] || "",
      id: r["ID"] || r["id"] || "",
    })).filter(a => a.nombre || a.apellidos);
  }

  function _getUnidades() {
    try {
      const rows = _sheetToJson("Unidades");
      return rows.map(r => ({
        numero: r["Numero"] || r["numero"] || r["Nº"] || "",
        nombre: r["Nombre"] || r["nombre"] || "",
        evaluacion: r["Evaluacion"] || r["evaluacion"] || r["Evaluación"] || "",
      })).filter(u => u.nombre);
    } catch { return []; }
  }

  function _getRraaCriterios() {
    try {
      const rraa = _sheetToJson("RRAA").map(r => ({
        codigo: r["Codigo"] || r["código"] || r["Código"] || "",
        descripcion: r["Descripcion"] || r["descripción"] || r["Descripción"] || "",
      })).filter(r => r.codigo);
      const criterios = _sheetToJson("Criterios").map(r => ({
        codigo: r["Codigo"] || r["código"] || "",
        descripcion: r["Descripcion"] || r["descripción"] || "",
        ra: r["RA"] || r["ra"] || "",
        peso: r["Peso"] || r["peso"] || 0,
      })).filter(c => c.codigo);
      return { rraa, criterios };
    } catch { return { rraa: [], criterios: [] }; }
  }

  function _getNotasActividad({ unidad, tipo, actividad }) {
    try {
      const hoja = `U${unidad}_${tipo}`;
      const rows = _sheetToJson(hoja);
      const col = `Act${actividad}`;
      return rows.map(r => ({
        alumno: r["Alumno"] || r["alumno"] || "",
        nota: r[col] !== undefined ? r[col] : "",
      }));
    } catch { return []; }
  }

  function _saveNotasActividad({ unidad, tipo, actividad, notas }) {
    const hoja = `U${unidad}_${tipo}`;
    const ws = _sheet(hoja);
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const header = rows[0] || [];
    const colIdx = header.indexOf(`Act${actividad}`);
    if (colIdx < 0) throw new Error(`Columna Act${actividad} no encontrada en ${hoja}`);
    notas.forEach((n, i) => {
      const cell = XLSX.utils.encode_cell({ r: i + 1, c: colIdx });
      ws[cell] = { v: n.nota === "" ? "" : Number(n.nota), t: n.nota === "" ? "s" : "n" };
    });
    _saveToStorage();
    _downloadWorkbook();
    return { ok: true };
  }

  // ── API pública ─────────────────────────────────────────────────────────────

  window.electronExcel = {
    selectFile: () => _openFilePicker(),

    getSelectedFile: () => {
      if (_fileName && (_workbook || _loadFromStorage())) {
        return Promise.resolve(_buildResult());
      }
      return Promise.resolve(null);
    },

    setSelectedFile: (filePath) => {
      if (_fileName === filePath && (_workbook || _loadFromStorage())) {
        return Promise.resolve(_buildResult());
      }
      return Promise.resolve(null);
    },

    verifyFileExists: (_filePath) => {
      return Promise.resolve(!!_workbook || !!localStorage.getItem(DATA_KEY));
    },

    getUnidades: () => Promise.resolve(_getUnidades()),
    saveUnidades: (unidades) => {
      // Actualiza hoja Unidades
      const wb = _wb();
      if (!wb) throw new Error("Sin archivo");
      const ws = XLSX.utils.json_to_sheet(unidades);
      wb.Sheets["Unidades"] = ws;
      _saveToStorage();
      _downloadWorkbook();
      return Promise.resolve({ ok: true });
    },

    getRraaCriterios: () => Promise.resolve(_getRraaCriterios()),
    saveRraaCriterios: (payloadOrRraa, criterios, ponderacionesUnidad = []) => {
      const payload = Array.isArray(payloadOrRraa)
        ? { rraa: payloadOrRraa, criterios, ponderacionesUnidad }
        : payloadOrRraa;
      const wb = _wb();
      if (!wb) throw new Error("Sin archivo");
      wb.Sheets["RRAA"] = XLSX.utils.json_to_sheet(payload.rraa || []);
      wb.Sheets["Criterios"] = XLSX.utils.json_to_sheet(payload.criterios || []);
      _saveToStorage();
      _downloadWorkbook();
      return Promise.resolve({ ok: true });
    },

    saveAlumnos: (alumnos) => {
      const wb = _wb();
      if (!wb) throw new Error("Sin archivo");
      wb.Sheets["Alumnos"] = XLSX.utils.json_to_sheet(alumnos);
      _saveToStorage();
      _downloadWorkbook();
      return Promise.resolve({ ok: true });
    },

    getNotasActividad: (payload) => Promise.resolve(_getNotasActividad(payload)),
    getNotasActividadesTipo: ({ unidad, tipo }) => {
      try {
        const hoja = `U${unidad}_${tipo}`;
        return Promise.resolve(_sheetToJson(hoja));
      } catch { return Promise.resolve([]); }
    },
    saveNotasActividad: (payload) => Promise.resolve(_saveNotasActividad(payload)),

    saveCeNotas: (payload) => {
      // Guarda notas de criterios de evaluación
      const wb = _wb();
      if (!wb) throw new Error("Sin archivo");
      const hoja = `CE_U${payload.unidad}`;
      wb.Sheets[hoja] = XLSX.utils.json_to_sheet(payload.notas || []);
      if (!wb.SheetNames.includes(hoja)) wb.SheetNames.push(hoja);
      _saveToStorage();
      _downloadWorkbook();
      return Promise.resolve({ ok: true });
    },

    addActividad: (payload) => {
      const wb = _wb();
      if (!wb) throw new Error("Sin archivo");
      const hoja = `U${payload.unidad}_${payload.tipo}`;
      const ws = wb.Sheets[hoja];
      if (!ws) throw new Error(`Hoja ${hoja} no encontrada`);
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      rows[0].push(`Act${payload.numero}`);
      rows.slice(1).forEach(r => r.push(""));
      wb.Sheets[hoja] = XLSX.utils.aoa_to_sheet(rows);
      _saveToStorage();
      _downloadWorkbook();
      return Promise.resolve({ ok: true });
    },

    getNotasUnidad: (payload) => {
      try {
        return Promise.resolve(_sheetToJson(`U${payload.unidad}_resumen`));
      } catch { return Promise.resolve([]); }
    },

    getNotasEvaluacion: (payload) => {
      try {
        return Promise.resolve(_sheetToJson(`Eval${payload.evaluacion}`));
      } catch { return Promise.resolve([]); }
    },

    getNotasEvaluacionAlumno: (payload) => {
      try {
        const rows = _sheetToJson(`Eval${payload.evaluacion}`);
        return Promise.resolve(rows.find(r => r["Alumno"] === payload.alumno) || null);
      } catch { return Promise.resolve(null); }
    },

    getAlumnosInformes: () => {
      try { return Promise.resolve(_getAlumnos()); }
      catch { return Promise.resolve([]); }
    },

    openExternal: (url) => { window.open(url, "_blank"); return Promise.resolve(); },

    getDiarioData: () => {
      try { return Promise.resolve(_sheetToJson("Diario")); }
      catch { return Promise.resolve([]); }
    },

    saveDiarioEntrada: (payload) => {
      const wb = _wb();
      if (!wb) throw new Error("Sin archivo");
      let rows = [];
      try { rows = _sheetToJson("Diario"); } catch {}
      rows.push(payload);
      wb.Sheets["Diario"] = XLSX.utils.json_to_sheet(rows);
      if (!wb.SheetNames.includes("Diario")) wb.SheetNames.push("Diario");
      _saveToStorage();
      _downloadWorkbook();
      return Promise.resolve({ ok: true });
    },

    deleteDiarioEntrada: (payload) => {
      const wb = _wb();
      if (!wb) throw new Error("Sin archivo");
      let rows = [];
      try { rows = _sheetToJson("Diario"); } catch {}
      rows = rows.filter(r => !(r.fecha === payload.fecha && r.texto === payload.texto));
      wb.Sheets["Diario"] = XLSX.utils.json_to_sheet(rows);
      _saveToStorage();
      _downloadWorkbook();
      return Promise.resolve({ ok: true });
    },
  };

  // Intentar cargar desde storage al arrancar
  _loadFromStorage();

})();
