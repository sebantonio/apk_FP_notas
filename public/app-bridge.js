(function () {
  if (window.electronExcel) return;

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

  const FILE_KEY = "android_excel_file_name";
  const LEGACY_DATA_KEY = "android_excel_data";
  const DB_NAME = "apk_fp_notas";
  const DB_STORE = "excel";
  const DB_KEY = "selected_file";

  let _workbook = null;
  let _fileName = localStorage.getItem(FILE_KEY) || null;
  let _loadPromise = null;

  function _openDb() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) {
        reject(new Error("IndexedDB no disponible"));
        return;
      }
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => request.result.createObjectStore(DB_STORE);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("No se pudo abrir IndexedDB"));
    });
  }

  async function _dbGet() {
    const db = await _openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readonly");
      const request = tx.objectStore(DB_STORE).get(DB_KEY);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error || new Error("No se pudo leer IndexedDB"));
      tx.oncomplete = () => db.close();
      tx.onerror = () => db.close();
    });
  }

  async function _dbSet(record) {
    const db = await _openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readwrite");
      tx.objectStore(DB_STORE).put(record, DB_KEY);
      tx.oncomplete = () => {
        db.close();
        resolve();
      };
      tx.onerror = () => {
        db.close();
        reject(tx.error || new Error("No se pudo guardar IndexedDB"));
      };
    });
  }

  function _loadLegacyStorage() {
    const b64 = localStorage.getItem(LEGACY_DATA_KEY);
    if (!b64) return false;
    try {
      const bin = atob(b64);
      const buf = Uint8Array.from(bin, c => c.charCodeAt(0));
      _workbook = XLSX.read(buf, { type: "array", cellDates: true });
      localStorage.removeItem(LEGACY_DATA_KEY);
      return true;
    } catch {
      localStorage.removeItem(LEGACY_DATA_KEY);
      return false;
    }
  }

  async function _loadFromDb() {
    if (_workbook) return true;
    try {
      const record = await _dbGet();
      if (!record || !record.buffer) return false;
      _fileName = record.fileName || _fileName;
      if (_fileName) localStorage.setItem(FILE_KEY, _fileName);
      _workbook = XLSX.read(new Uint8Array(record.buffer), { type: "array", cellDates: true });
      return true;
    } catch (err) {
      console.warn("No se pudo cargar el Excel guardado.", err);
      return false;
    }
  }

  async function _ensureWorkbook() {
    if (_workbook) return _workbook;
    if (_loadLegacyStorage()) return _workbook;
    if (!_loadPromise) {
      _loadPromise = _loadFromDb().finally(() => {
        _loadPromise = null;
      });
    }
    await _loadPromise;
    return _workbook;
  }

  function _wb() {
    if (!_workbook) throw new Error("No hay archivo Excel cargado.");
    return _workbook;
  }

  function _sheet(name) {
    const wb = _wb();
    if (!wb.SheetNames.includes(name)) throw new Error(`Hoja "${name}" no encontrada.`);
    return wb.Sheets[name];
  }

  function _sheetToJson(name, opts = {}) {
    return XLSX.utils.sheet_to_json(_sheet(name), { defval: "", ...opts });
  }

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

  function _buildResult() {
    const wb = _wb();
    const result = { filePath: _fileName, fileName: _fileName };

    try {
      result.alumnos = XLSX.utils.sheet_to_json(wb.Sheets["DATOS"] || wb.Sheets[wb.SheetNames[0]], { defval: "" });
    } catch { result.alumnos = []; }

    try {
      result.rraa = XLSX.utils.sheet_to_json(wb.Sheets["RRAA"] || {}, { defval: "" });
    } catch { result.rraa = []; }

    try {
      result.criterios = XLSX.utils.sheet_to_json(wb.Sheets["Criterios"] || {}, { defval: "" });
    } catch { result.criterios = []; }

    result.ponderacionesUnidad = [];
    return result;
  }

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
      const cancelTimer = setTimeout(() => {
        if (!resolved) {
          resolved = true;
          cleanup();
          resolve(null);
        }
      }, 60000);

      input.addEventListener("change", async () => {
        if (resolved) return;
        resolved = true;
        clearTimeout(cancelTimer);
        const file = input.files && input.files[0];
        cleanup();
        if (!file) {
          resolve(null);
          return;
        }
        try {
          const buffer = await file.arrayBuffer();
          _workbook = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true });
          _fileName = file.name;
          localStorage.setItem(FILE_KEY, _fileName);
          localStorage.removeItem(LEGACY_DATA_KEY);
          await _dbSet({ fileName: _fileName, buffer });
          resolve(_buildResult());
        } catch (err) {
          console.error("No se pudo leer el Excel seleccionado.", err);
          resolve(null);
        }
      });

      input.addEventListener("cancel", () => {
        if (resolved) return;
        resolved = true;
        clearTimeout(cancelTimer);
        cleanup();
        resolve(null);
      });

      input.click();
    });
  }

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
    _downloadWorkbook();
    return { ok: true };
  }

  window.electronExcel = {
    selectFile: () => _openFilePicker(),
    getSelectedFile: async () => (_fileName && await _ensureWorkbook()) ? _buildResult() : null,
    setSelectedFile: async (filePath) => (_fileName === filePath && await _ensureWorkbook()) ? _buildResult() : null,
    verifyFileExists: async () => !!_workbook || !!(await _dbGet()),

    getUnidades: async () => {
      await _ensureWorkbook();
      return _getUnidades();
    },
    saveUnidades: async (unidades) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      wb.Sheets["Unidades"] = XLSX.utils.json_to_sheet(unidades);
      _downloadWorkbook();
      return { ok: true };
    },

    getRraaCriterios: async () => {
      await _ensureWorkbook();
      return _getRraaCriterios();
    },
    saveRraaCriterios: async (payloadOrRraa, criterios, ponderacionesUnidad = []) => {
      const payload = Array.isArray(payloadOrRraa)
        ? { rraa: payloadOrRraa, criterios, ponderacionesUnidad }
        : payloadOrRraa;
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      wb.Sheets["RRAA"] = XLSX.utils.json_to_sheet(payload.rraa || []);
      wb.Sheets["Criterios"] = XLSX.utils.json_to_sheet(payload.criterios || []);
      _downloadWorkbook();
      return { ok: true };
    },

    saveAlumnos: async (alumnos) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      wb.Sheets["Alumnos"] = XLSX.utils.json_to_sheet(alumnos);
      _downloadWorkbook();
      return { ok: true };
    },

    getNotasActividad: async (payload) => {
      await _ensureWorkbook();
      return _getNotasActividad(payload);
    },
    getNotasActividadesTipo: async ({ unidad, tipo }) => {
      await _ensureWorkbook();
      try { return _sheetToJson(`U${unidad}_${tipo}`); }
      catch { return []; }
    },
    saveNotasActividad: async (payload) => {
      await _ensureWorkbook();
      return _saveNotasActividad(payload);
    },

    saveCeNotas: async (payload) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      const hoja = `CE_U${payload.unidad}`;
      wb.Sheets[hoja] = XLSX.utils.json_to_sheet(payload.notas || []);
      if (!wb.SheetNames.includes(hoja)) wb.SheetNames.push(hoja);
      _downloadWorkbook();
      return { ok: true };
    },

    addActividad: async (payload) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      const hoja = `U${payload.unidad}_${payload.tipo}`;
      const ws = wb.Sheets[hoja];
      if (!ws) throw new Error(`Hoja ${hoja} no encontrada`);
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      rows[0].push(`Act${payload.numero}`);
      rows.slice(1).forEach(r => r.push(""));
      wb.Sheets[hoja] = XLSX.utils.aoa_to_sheet(rows);
      _downloadWorkbook();
      return { ok: true };
    },

    getNotasUnidad: async (payload) => {
      await _ensureWorkbook();
      try { return _sheetToJson(`U${payload.unidad}_resumen`); }
      catch { return []; }
    },
    getNotasEvaluacion: async (payload) => {
      await _ensureWorkbook();
      try { return _sheetToJson(`Eval${payload.evaluacion}`); }
      catch { return []; }
    },
    getNotasEvaluacionAlumno: async (payload) => {
      await _ensureWorkbook();
      try {
        const rows = _sheetToJson(`Eval${payload.evaluacion}`);
        return rows.find(r => r["Alumno"] === payload.alumno) || null;
      } catch { return null; }
    },

    getAlumnosInformes: async () => {
      await _ensureWorkbook();
      try { return _getAlumnos(); }
      catch { return []; }
    },
    openExternal: (url) => { window.open(url, "_blank"); return Promise.resolve(); },

    getDiarioData: async () => {
      await _ensureWorkbook();
      try { return _sheetToJson("Diario"); }
      catch { return []; }
    },
    saveDiarioEntrada: async (payload) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      let rows = [];
      try { rows = _sheetToJson("Diario"); } catch {}
      rows.push(payload);
      wb.Sheets["Diario"] = XLSX.utils.json_to_sheet(rows);
      if (!wb.SheetNames.includes("Diario")) wb.SheetNames.push("Diario");
      _downloadWorkbook();
      return { ok: true };
    },
    deleteDiarioEntrada: async (payload) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      let rows = [];
      try { rows = _sheetToJson("Diario"); } catch {}
      rows = rows.filter(r => !(r.fecha === payload.fecha && r.texto === payload.texto));
      wb.Sheets["Diario"] = XLSX.utils.json_to_sheet(rows);
      _downloadWorkbook();
      return { ok: true };
    },
  };

})();
