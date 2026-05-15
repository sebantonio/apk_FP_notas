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
  let _sourceBuffer = null;
  let _fileName = localStorage.getItem(FILE_KEY) || null;
  let _loadPromise = null;
  let _rowsCache = new Map();

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
      _sourceBuffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
      _clearRowsCache();
      localStorage.removeItem(LEGACY_DATA_KEY);
      return true;
    } catch {
      localStorage.removeItem(LEGACY_DATA_KEY);
      return false;
    }
  }

  async function _loadFromDb() {
    if (_sourceBuffer || _workbook) return true;
    try {
      const record = await _dbGet();
      if (!record || !record.buffer) return false;
      _fileName = record.fileName || _fileName;
      if (_fileName) localStorage.setItem(FILE_KEY, _fileName);
      _sourceBuffer = record.buffer;
      _clearRowsCache();
      return true;
    } catch (err) {
      console.warn("No se pudo cargar el Excel guardado.", err);
      return false;
    }
  }

  async function _ensureWorkbook() {
    if (_workbook) return _workbook;
    if (!_sourceBuffer && _loadLegacyStorage()) {
      // legacy data is now available as a source buffer
    }
    if (!_sourceBuffer) {
      if (!_loadPromise) {
        _loadPromise = _loadFromDb().finally(() => {
          _loadPromise = null;
        });
      }
      await _loadPromise;
    }
    if (!_workbook && _sourceBuffer) {
      _workbook = XLSX.read(new Uint8Array(_sourceBuffer), { type: "array", cellDates: true });
      _clearRowsCache();
    }
    return _workbook;
  }

  async function _ensureSourceBuffer() {
    if (_sourceBuffer) return _sourceBuffer;
    if (_loadLegacyStorage()) return _sourceBuffer;
    if (!_loadPromise) {
      _loadPromise = _loadFromDb().finally(() => {
        _loadPromise = null;
      });
    }
    await _loadPromise;
    return _sourceBuffer;
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

  function _readPartialWorkbook(sheetNames) {
    if (!_sourceBuffer) return _workbook;
    const sheets = Array.isArray(sheetNames) ? sheetNames : [sheetNames];
    return XLSX.read(new Uint8Array(_sourceBuffer), {
      type: "array",
      cellDates: true,
      sheets,
    });
  }

  function _sheetToJson(name, opts = {}) {
    return XLSX.utils.sheet_to_json(_sheet(name), { defval: "", ...opts });
  }

  function _clearRowsCache(sheetName) {
    if (!sheetName) {
      _rowsCache = new Map();
      return;
    }
    _rowsCache.delete(sheetName);
  }

  function _rows(name) {
    if (!_rowsCache.has(name)) {
      const wb = _workbook || _readPartialWorkbook(name);
      if (!wb || !wb.Sheets || !wb.Sheets[name]) throw new Error(`Hoja "${name}" no encontrada.`);
      _rowsCache.set(name, XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "" }));
    }
    return _rowsCache.get(name);
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
      result.alumnos = _getAlumnos();
    } catch { result.alumnos = []; }

    try {
      const rraaCriterios = _getRraaCriterios();
      result.rraa = rraaCriterios.rraa;
      result.criterios = rraaCriterios.criterios;
      result.ponderacionesUnidad = rraaCriterios.ponderacionesUnidad || [];
    } catch {
      result.rraa = [];
      result.criterios = [];
      result.ponderacionesUnidad = [];
    }
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
          _sourceBuffer = buffer;
          _workbook = null;
          _clearRowsCache();
          _fileName = file.name;
          localStorage.setItem(FILE_KEY, _fileName);
          localStorage.removeItem(LEGACY_DATA_KEY);
          await _dbSet({ fileName: _fileName, buffer });
          resolve({ fileName: _fileName, filePath: _fileName });
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
    const rows = _rows("DATOS");
    const header = rows.findIndex((row) =>
      row[1] && String(row[1]).toUpperCase().includes("ALUMNADO")
    );
    if (header < 0) return [];

    const alumnos = [];
    for (let rowIdx = header + 1; rowIdx < rows.length; rowIdx += 1) {
      const row = rows[rowIdx] || [];
      const nombre = row[1];
      if (!nombre || String(nombre).trim() === "") break;
      alumnos.push({
        numero: row[0] || alumnos.length + 1,
        nombre,
        fechaNac: row[2] || "",
        rowIdx,
        excelRowIdx: rowIdx,
      });
    }
    return alumnos;
  }

  function _getUnidades() {
    const rows = _rows("DATOS");
    const start = _findMainUnitsStart(rows);
    if (start < 0) return [];

    const unidades = [];
    for (let idx = 0; idx < 16; idx += 1) {
      const row = rows[start + idx] || [];
      const codigo = String(row[8] || `U${idx + 1}`);
      const nombre = String(row[9] || "");
      const evaluacion = String(row[10] || "");
      const horas = row[11] === undefined || row[11] === null ? "" : String(row[11]);
      if (codigo || nombre || evaluacion || horas) {
        unidades.push({
          codigo,
          numero: codigo,
          nombre,
          evaluacion,
          horas,
          label: nombre ? `${codigo} - ${nombre}` : codigo,
        });
      }
    }
    return unidades;
  }

  function _getRraaCriterios() {
    const datosRows = _rows("DATOS");
    const pesosRows = _rows("PESOS");
    const rraa = _readRraa(datosRows);
    const textos = _extractCriteriaTexts(datosRows);
    const criterios = _readCriterios(pesosRows, rraa, textos);
    const ponderacionesUnidad = _readPonderacionesUnidad(pesosRows, criterios);
    return { fileName: _fileName, filePath: _fileName, rraa, criterios, ponderacionesUnidad };
  }

  function _findMainUnitsStart(rows) {
    const header = rows.findIndex((row) =>
      row[8] &&
      row[9] &&
      String(row[9]).toUpperCase().includes("UNIDADES")
    );
    return header === -1 ? -1 : header + 1;
  }

  function _findEvaluationUnitsStart(rows, evaluacion) {
    const numero = String(evaluacion).replace("ª", "");
    const header = rows.findIndex((row) =>
      row[10] &&
      String(row[10]).toUpperCase().includes("UNIDADES") &&
      String(row[10]).includes(numero)
    );
    return header === -1 ? -1 : header + 1;
  }

  function _saveUnidadesToDatos(unidades) {
    const sheet = _sheet("DATOS");
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const start = _findMainUnitsStart(rows);
    if (start < 0) throw new Error("No se encontró la tabla principal de unidades en DATOS.");

    const normalizadas = (unidades || []).slice(0, 16);
    for (let idx = 0; idx < 16; idx += 1) {
      const rowIdx = start + idx;
      const unidad = normalizadas[idx] || { codigo: `U${idx + 1}`, nombre: "", evaluacion: "", horas: "" };
      rows[rowIdx] = rows[rowIdx] || [];
      rows[rowIdx][8] = unidad.codigo || `U${idx + 1}`;
      rows[rowIdx][9] = unidad.nombre || "";
      rows[rowIdx][10] = unidad.evaluacion || "";
      rows[rowIdx][11] = unidad.horas || "";
    }

    ["1ª", "2ª", "3ª"].forEach((evaluacion) => {
      const blockStart = _findEvaluationUnitsStart(rows, evaluacion);
      if (blockStart < 0) return;

      const filtradas = normalizadas.filter((unidad) => unidad.evaluacion === evaluacion);
      for (let idx = 0; idx < 16; idx += 1) {
        const rowIdx = blockStart + idx;
        const unidad = filtradas[idx];
        rows[rowIdx] = rows[rowIdx] || [];
        rows[rowIdx][9] = `U${idx + 1}`;
        rows[rowIdx][10] = unidad ? unidad.nombre : "";
        rows[rowIdx][11] = evaluacion;
      }
    });

    _wb().Sheets["DATOS"] = _replaceSheetKeepingMeta(sheet, rows);
    _clearRowsCache("DATOS");
  }

  function _readRraaFromColumns(rows, headerCol, numberCol, descriptionCol) {
    const header = rows.findIndex((row) => {
      const value = row && row[headerCol];
      const text = String(value || "").toUpperCase();
      return text.includes("RRAA") || text.includes("RESULTADOS DE APRENDIZAJE");
    });
    if (header < 0) return [];

    const rraa = [];
    for (let rowIdx = header + 1; rowIdx < rows.length; rowIdx += 1) {
      const row = rows[rowIdx] || [];
      const numero = row[numberCol];
      const descripcion = row[descriptionCol];
      if (!numero || !descripcion || String(descripcion).trim() === "") break;
      rraa.push({ numero, descripcion });
    }
    return rraa;
  }

  function _readRraa(rows) {
    const candidates = [
      _readRraaFromColumns(rows, 1, 0, 1),
      _readRraaFromColumns(rows, 6, 5, 6),
    ];
    return candidates.find((items) => items.length > 0) || [];
  }

  function _isCriteriaCode(value) {
    return Boolean(value && /^\d+\.[a-z]\)/i.test(String(value).trim()));
  }

  function _raNumberFromCriteria(value) {
    const match = String(value || "").match(/^(\d+)/);
    return match ? Number(match[1]) : 0;
  }

  function _normalizeCriteriaCode(value) {
    return String(value || "").trim().replace(/\)$/u, "").toLowerCase();
  }

  function _extractCriteriaTexts(rows) {
    const textos = {};
    rows.forEach((row) => {
      const codigo = row && row[21];
      const texto = row && row[22];
      if (codigo && texto) {
        textos[_normalizeCriteriaCode(codigo)] = String(texto).trim();
      }
    });
    return textos;
  }

  function _readCriterios(pesosRows, rraa, textos) {
    const criterios = [];
    const headerCriterios = pesosRows[3] || [];
    const ponderaciones = pesosRows[21] || [];

    for (let colIdx = 0; colIdx < headerCriterios.length; colIdx += 1) {
      const codigo = headerCriterios[colIdx];
      if (!_isCriteriaCode(codigo)) continue;

      const raNumero = _raNumberFromCriteria(codigo);
      const ra = rraa.find(item => Number(item.numero) === raNumero);
      criterios.push({
        numero: criterios.length + 1,
        codigo: String(codigo),
        nombre: String(codigo),
        originalCodigo: String(codigo),
        raNumero,
        raDescripcion: ra ? ra.descripcion : "",
        ponderacion: ponderaciones[colIdx] || 0,
        ponderacionInstituto: ponderaciones[colIdx + 1] || 0,
        ponderacionEmpresa: ponderaciones[colIdx + 2] || 0,
        texto: textos[_normalizeCriteriaCode(codigo)] || "",
        colIdx,
      });
    }

    return criterios;
  }

  function _readPonderacionesUnidad(pesosRows, criterios) {
    const unidades = [];
    for (let rowIdx = 5; rowIdx < 21; rowIdx += 1) {
      const row = pesosRows[rowIdx] || [];
      const ponderaciones = {};
      criterios.forEach((criterio) => {
        ponderaciones[criterio.colIdx] = {
          ponderacion: row[criterio.colIdx] || 0,
          ponderacionInstituto: row[criterio.colIdx + 1] || 0,
          ponderacionEmpresa: row[criterio.colIdx + 2] || 0,
        };
      });
      unidades.push({
        numero: rowIdx - 4,
        rowIdx,
        nombre: row[0] && String(row[0]) !== "0" ? String(row[0]) : "",
        ponderaciones,
      });
    }
    return unidades;
  }

  function _replaceSheetKeepingMeta(oldSheet, rows) {
    const nextSheet = XLSX.utils.aoa_to_sheet(rows);
    Object.keys(oldSheet || {}).forEach((key) => {
      if (key.startsWith("!")) nextSheet[key] = oldSheet[key];
    });
    return nextSheet;
  }

  function _saveRraaToDatos(rraa) {
    const sheet = _sheet("DATOS");
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const header = rows.findIndex((row) =>
      row[1] && String(row[1]).toUpperCase().includes("RRAA")
    );
    if (header < 0) throw new Error('No se encontró la sección RRAA en la hoja DATOS.');

    const start = header + 1;
    for (let idx = 0; idx < 40; idx += 1) {
      const rowIdx = start + idx;
      rows[rowIdx] = rows[rowIdx] || [];
      rows[rowIdx][0] = "";
      rows[rowIdx][1] = "";
    }

    (rraa || []).forEach((item, idx) => {
      const rowIdx = start + idx;
      rows[rowIdx] = rows[rowIdx] || [];
      rows[rowIdx][0] = item.numero || idx + 1;
      rows[rowIdx][1] = item.descripcion || "";
    });

    _wb().Sheets["DATOS"] = _replaceSheetKeepingMeta(sheet, rows);
    _clearRowsCache("DATOS");
  }

  function _saveCriteriosToPesos(criterios, ponderacionesUnidad) {
    const sheet = _sheet("PESOS");
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    (criterios || []).forEach((crit) => {
      const colIdx = Number.isInteger(crit.colIdx) ? crit.colIdx : null;
      if (colIdx === null || colIdx < 0) return;
      rows[3] = rows[3] || [];
      rows[21] = rows[21] || [];
      rows[3][colIdx] = crit.codigo || crit.nombre || crit.originalCodigo || "";
      rows[21][colIdx] = crit.ponderacion || 0;
      rows[21][colIdx + 1] = crit.ponderacionInstituto || 0;
      rows[21][colIdx + 2] = crit.ponderacionEmpresa || 0;
    });

    (ponderacionesUnidad || []).forEach((unidad) => {
      const rowIdx = Number(unidad.rowIdx);
      if (!Number.isInteger(rowIdx) || rowIdx < 0) return;
      rows[rowIdx] = rows[rowIdx] || [];
      Object.entries(unidad.ponderaciones || {}).forEach(([key, values]) => {
        const colIdx = Number(key);
        if (!Number.isInteger(colIdx)) return;
        rows[rowIdx][colIdx] = values.ponderacion || 0;
        rows[rowIdx][colIdx + 1] = values.ponderacionInstituto || 0;
        rows[rowIdx][colIdx + 2] = values.ponderacionEmpresa || 0;
      });
    });

    _wb().Sheets["PESOS"] = _replaceSheetKeepingMeta(sheet, rows);
    _clearRowsCache("PESOS");
  }

  function _getNotasActividad({ unidad, tipo, actividad }) {
    const unidades = _getUnidades();
    const tipos = _getTiposActividad(unidad);
    const rraaData = _getRraaCriterios();
    const hoja = `U${unidad}_${tipo}`;
    const col = `Act${actividad}`;
    let rows = [];
    let notas = [];

    try {
      rows = _sheetToJson(hoja);
      notas = rows.map((r, idx) => ({
        numero: r["Nº"] || r["Numero"] || r["numero"] || idx + 1,
        nombre: r["Alumno"] || r["ALUMNO"] || r["alumno"] || r["Nombre"] || r["nombre"] || "",
        nota: r[col] !== undefined ? r[col] : "",
        rowIdx: idx + 1,
      })).filter(item => item.nombre);
    } catch {
      notas = _getAlumnos().map((alumno, idx) => ({
        numero: alumno.numero || idx + 1,
        nombre: alumno.nombre,
        nota: "",
        rowIdx: alumno.rowIdx,
      }));
    }

    const type = tipos.find(item => item.key === tipo) || tipos[0] || { actividades: [] };
    const block = (type.actividades || []).find(item => Number(item.numero) === Number(actividad)) || {
      numero: actividad,
      nombre: "",
      incluida: true,
    };

    return {
      fileName: _fileName,
      filePath: _fileName,
      unidad,
      tipo,
      actividad,
      unidades,
      tipos,
      actividades: type.actividades || [],
      block,
      notas,
      rraa: rraaData.rraa,
      criterios: rraaData.criterios,
      todasCriterios: rraaData.criterios,
      ponderacionesUnidad: rraaData.ponderacionesUnidad || [],
    };
  }

  function _getTiposActividad(unidad) {
    const definitions = [
      { key: "practicas", label: "Practicas" },
      { key: "memorias", label: "Memorias" },
      { key: "otros", label: "Otras actividades" },
      { key: "controles", label: "Controles" },
    ];

    return definitions.map((definition) => {
      const hoja = `U${unidad}_${definition.key}`;
      let actividades = [];
      try {
        const header = _rows(hoja)[0] || [];
        actividades = header
          .map((value) => {
            const match = String(value || "").match(/^Act(\d+)/i);
            return match ? { numero: Number(match[1]), nombre: "", incluida: true } : null;
          })
          .filter(Boolean);
      } catch {
        actividades = [];
      }
      if (!actividades.length) actividades = [{ numero: 1, nombre: "", incluida: true }];
      return {
        ...definition,
        total: actividades.length,
        incluidas: actividades.filter(item => item.incluida).length,
        actividades,
      };
    });
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
    _clearRowsCache(hoja);
    _downloadWorkbook();
    return { ok: true };
  }

  window.electronExcel = {
    selectFile: () => _openFilePicker(),
    getSelectedFile: async () => (_fileName && await _ensureSourceBuffer())
      ? { fileName: _fileName, filePath: _fileName }
      : null,
    setSelectedFile: async (filePath) => (_fileName === filePath && await _ensureSourceBuffer())
      ? { fileName: _fileName, filePath: _fileName }
      : null,
    verifyFileExists: async () => !!_workbook || !!(await _dbGet()),

    getUnidades: async () => {
      await _ensureSourceBuffer();
      return { fileName: _fileName, filePath: _fileName, unidades: _getUnidades() };
    },
    getAlumnos: async () => {
      await _ensureSourceBuffer();
      return { fileName: _fileName, filePath: _fileName, alumnos: _getAlumnos() };
    },
    saveUnidades: async (unidades) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      _saveUnidadesToDatos(unidades);
      _downloadWorkbook();
      return { ok: true, fileName: _fileName, unidades: _getUnidades() };
    },

    getRraaCriterios: async () => {
      await _ensureSourceBuffer();
      return _getRraaCriterios();
    },
    saveRraaCriterios: async (payloadOrRraa, criterios, ponderacionesUnidad = []) => {
      const payload = Array.isArray(payloadOrRraa)
        ? { rraa: payloadOrRraa, criterios, ponderacionesUnidad }
        : payloadOrRraa;
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      _saveRraaToDatos(payload.rraa || []);
      _saveCriteriosToPesos(payload.criterios || [], payload.ponderacionesUnidad || []);
      _downloadWorkbook();
      return { ok: true, ..._getRraaCriterios() };
    },

    saveAlumnos: async (alumnos) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      wb.Sheets["Alumnos"] = XLSX.utils.json_to_sheet(alumnos);
      _clearRowsCache("Alumnos");
      _downloadWorkbook();
      return { ok: true, fileName: _fileName, alumnos };
    },

    getNotasActividad: async (payload) => {
      await _ensureWorkbook();
      return _getNotasActividad(payload);
    },
    getNotasActividadesTipo: async ({ unidad, tipo }) => {
      await _ensureWorkbook();
      const tipos = _getTiposActividad(unidad);
      const type = tipos.find(item => item.key === tipo) || tipos[0];
      const actividades = type ? type.actividades : [];
      const notas = [];
      for (const actividad of actividades) {
        notas.push({
          actividad: actividad.numero,
          nombre: actividad.nombre || "",
          notas: _getNotasActividad({ unidad, tipo, actividad: actividad.numero }).notas,
        });
      }
      return { fileName: _fileName, unidad, tipo, tipos, actividades, notas };
    },
    saveNotasActividad: async (payload) => {
      await _ensureWorkbook();
      _saveNotasActividad(payload);
      return _getNotasActividad(payload);
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
