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
  const DB_PATCHES_KEY = "cell_patches";

  let _workbook = null;
  let _sourceBuffer = null;
  let _fileName = localStorage.getItem(FILE_KEY) || null;
  let _loadPromise = null;
  let _rowsCache = new Map();
  let _pendingPatches = [];

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
      if (_isAndroid()) {
        const patches = await _dbGetPatches();
        if (patches && patches.length) _applyPatchesToWorkbook(patches);
      }
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

  function _unitSheetName(unidad) {
    const value = String(unidad || "1").trim();
    return /^U\d+/i.test(value) ? value.toUpperCase() : `U${value}`;
  }

  function _activitySheetName(unidad, tipo) {
    return `${_unitSheetName(unidad)}_${tipo}`;
  }

  function _activityDefinitions() {
    return [
      { key: "practicas", label: "Practicas", match: ["PRACTICA", "PRÁCTICA", "PRÁCTICAS", "PRACTICAS"], notaOffset: 4 },
      { key: "memorias", label: "Memorias", match: ["MEMORIA", "MEMORIAS"], notaOffset: 3 },
      { key: "otros", label: "Otras actividades", match: ["OTROS INSTRUMENTOS", "O.INSTRU", "OTRAS", "OTROS", "O.I"], notaOffset: 3 },
      { key: "controles", label: "Controles", match: ["CONTROL", "CONTROLES", "PRUEBA"], notaOffset: 3 },
    ];
  }

  function _activityDefinition(tipo) {
    return _activityDefinitions().find((item) => item.key === tipo) || _activityDefinitions()[0];
  }

  function _matchesType(text, definition) {
    const upper = text.toUpperCase();
    return definition.match.some((term) => upper.includes(term.toUpperCase()));
  }

  // Detecta la estructura de bloques de actividades en la hoja unidad.
  // Devuelve { sheetName, tipoColMap, primeraFilaBloque, alturaBloque }
  // tipoColMap: { practicas: colIdx, memorias: colIdx, otros: colIdx, controles: colIdx }
  function _detectActivityLayout(unidad) {
    const unitSheet = _unitSheetName(unidad);
    const rows = _rows(unitSheet);
    const defs = _activityDefinitions();

    // Buscar la fila donde col A contiene "practicas" Y la fila siguiente tiene "N°" en col A.
    // Esto distingue los bloques reales de las filas de resumen.
    for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
      const row = rows[rowIdx] || [];
      const aCell = String(row[0] || "");
      if (!_matchesType(aCell, defs[0])) continue;

      // Verificar que la fila siguiente tiene "N°" o "N" en col A (cabecera del bloque)
      const nextRow = rows[rowIdx + 1] || [];
      const nextA = String(nextRow[0] || "").trim().toUpperCase();
      if (!nextA.startsWith("N")) continue; // No es un bloque real, es resumen

      // Esta fila es el inicio del primer bloque. Buscar otros tipos en la misma fila.
      const tipoColMap = { practicas: 0 };
      for (let colIdx = 1; colIdx < row.length; colIdx++) {
        const cellText = String(row[colIdx] || "");
        if (!cellText.trim()) continue;
        for (const def of defs.slice(1)) {
          if (_matchesType(cellText, def) && tipoColMap[def.key] === undefined) {
            tipoColMap[def.key] = colIdx;
          }
        }
      }

      // Calcular altura del bloque: buscar siguiente bloque del mismo tipo practicas
      let alturaBloque = 44; // valor por defecto
      for (let r2 = rowIdx + 1; r2 < Math.min(rows.length, rowIdx + 100); r2++) {
        const r2cell = String((rows[r2] || [])[0] || "");
        if (_matchesType(r2cell, defs[0])) {
          alturaBloque = r2 - rowIdx;
          break;
        }
      }

      return { sheetName: unitSheet, rows, tipoColMap, primeraFilaBloque: rowIdx, alturaBloque };
    }

    return { sheetName: unitSheet, rows, tipoColMap: { practicas: 0 }, primeraFilaBloque: 0, alturaBloque: 44 };
  }

  // Devuelve el colIdx de inicio de un tipo en su layout
  function _tipoColIdx(tipoColMap, tipo) {
    if (tipoColMap[tipo] !== undefined) return tipoColMap[tipo];
    // Si el tipo no se encontró, devolver 0 (practicas)
    return 0;
  }

  // Devuelve todos los bloques de un tipo: [{ numero, nombre, incluida, filaInicio }]
  function _getActivityBlocks(rows, tipoColIdx, primeraFilaBloque, alturaBloque) {
    const blocks = [];
    for (let filaInicio = primeraFilaBloque; filaInicio < rows.length; filaInicio += alturaBloque) {
      const titleRow = rows[filaInicio] || [];
      const titleCell = String(titleRow[tipoColIdx] || "");
      if (!titleCell.trim()) break;

      const numRow = rows[filaInicio + 1] || [];
      const numero = Number(numRow[tipoColIdx + 1]) || (blocks.length + 1);
      const nombre = String(numRow[tipoColIdx + 3] || ""); // col: N° | num | NOMBRE | <nombre real>

      const inclRow = rows[filaInicio + 2] || [];
      const inclCell = String(inclRow[tipoColIdx + 1] || "").toUpperCase().trim();
      const incluida = inclCell === "X" || inclCell === "x" || inclCell === "SI" || inclCell === "SÍ" || inclCell === "S";

      blocks.push({ numero, nombre, incluida, filaInicio });
    }
    return blocks;
  }

  function _activitySheetRows(unidad, tipo) {
    // Compatibilidad: intentar hoja separada primero
    const splitSheet = _activitySheetName(unidad, tipo);
    try {
      const rows = _rows(splitSheet);
      const defs = _activityDefinitions();
      const def = _activityDefinition(tipo);
      let headerRowIdx = -1;
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i] || [];
        if (row.some((cell) => /^Act\d+/i.test(String(cell || "").trim()))) {
          headerRowIdx = i;
          break;
        }
      }
      return { sheetName: splitSheet, rows, headerRowIdx };
    } catch {}

    const layout = _detectActivityLayout(unidad);
    return { sheetName: layout.sheetName, rows: layout.rows, headerRowIdx: -1, layout };
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
    const sheet = _workbook ? _sheet(name) : (_readPartialWorkbook(name).Sheets || {})[name];
    if (!sheet) throw new Error(`Hoja "${name}" no encontrada.`);
    return XLSX.utils.sheet_to_json(sheet, { defval: "", ...opts });
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

  function _isAndroid() {
    return /Android/i.test(navigator.userAgent);
  }

  async function _dbGetPatches() {
    try {
      const db = await _openDb();
      return new Promise((resolve) => {
        const tx = db.transaction(DB_STORE, "readonly");
        const req = tx.objectStore(DB_STORE).get(DB_PATCHES_KEY);
        req.onsuccess = () => { db.close(); resolve(req.result || []); };
        req.onerror = () => { db.close(); resolve([]); };
      });
    } catch { return []; }
  }

  async function _dbSavePatches(patches) {
    try {
      const db = await _openDb();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_STORE, "readwrite");
        tx.objectStore(DB_STORE).put(patches, DB_PATCHES_KEY);
        tx.oncomplete = () => { db.close(); resolve(); };
        tx.onerror = () => { db.close(); reject(tx.error); };
      });
    } catch { }
  }

  function _applyPatchesToWorkbook(patches) {
    if (!_workbook || !patches || !patches.length) return;
    for (const p of patches) {
      const ws = _workbook.Sheets[p.sheet];
      if (!ws) continue;
      const cellRef = XLSX.utils.encode_cell({ r: p.r, c: p.c });
      if (p.v === "" || p.v === null || p.v === undefined) {
        ws[cellRef] = { v: "", t: "s" };
      } else if (typeof p.v === "number") {
        ws[cellRef] = { v: p.v, t: "n" };
      } else {
        ws[cellRef] = { v: String(p.v), t: "s" };
      }
    }
  }

  function _recordPatch(sheet, r, c, v) {
    _pendingPatches.push({ sheet, r, c, v, ts: Date.now() });
  }

  async function _flushPatchesAndroid() {
    if (!_pendingPatches.length) return;
    const existing = await _dbGetPatches();
    const merged = [...existing, ..._pendingPatches];
    await _dbSavePatches(merged);
    _pendingPatches = [];
  }

  async function _downloadWorkbook() {
    if (!_workbook || !_fileName) return;
    if (_isAndroid()) {
      // En Android: guardar patches en IndexedDB, sin reserializar el workbook completo.
      await _flushPatchesAndroid();
      return;
    }
    const wbout = XLSX.write(_workbook, { bookType: "xlsx", type: "array" });
    const buffer = new Uint8Array(wbout).buffer;
    _sourceBuffer = buffer;
    const blob = new Blob([new Uint8Array(buffer)], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = _fileName;
    a.click();
    URL.revokeObjectURL(url);
    await _dbSet({ fileName: _fileName, buffer });
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
          _pendingPatches = [];
          _fileName = file.name;
          localStorage.setItem(FILE_KEY, _fileName);
          localStorage.removeItem(LEGACY_DATA_KEY);
          await _dbSet({ fileName: _fileName, buffer });
          await _dbSavePatches([]);
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

  function _getNotasActividad({ unidad, tipo, actividad, includeRraa = true }) {
    const unidades = _getUnidades();
    const tipos = _getTiposActividad(unidad);
    const rraaData = includeRraa ? _getRraaCriterios() : { rraa: [], criterios: [], ponderacionesUnidad: [] };
    let notas = [];

    try {
      const layout = _detectActivityLayout(unidad);
      const colIdx = _tipoColIdx(layout.tipoColMap, tipo);
      const def = _activityDefinition(tipo);
      const notaOffset = def.notaOffset !== undefined ? def.notaOffset : 4;
      const blocks = _getActivityBlocks(layout.rows, colIdx, layout.primeraFilaBloque, layout.alturaBloque);
      const block = blocks.find((b) => Number(b.numero) === Number(actividad));
      if (block) {
        const alumnosStartRow = block.filaInicio + 4;
        const notaColIdx = colIdx + notaOffset;
        for (let r = alumnosStartRow; r < block.filaInicio + layout.alturaBloque; r++) {
          const row = layout.rows[r] || [];
          const nombre = String(row[colIdx] || "").trim();
          if (!nombre) continue;
          notas.push({
            numero: r - alumnosStartRow + 1,
            nombre,
            nota: row[notaColIdx] !== undefined && row[notaColIdx] !== "" ? row[notaColIdx] : "",
            rowIdx: r,
          });
        }
      }
    } catch {
      notas = _getAlumnos().map((alumno, idx) => ({
        numero: alumno.numero || idx + 1,
        nombre: alumno.nombre,
        nota: "",
        rowIdx: alumno.rowIdx,
      }));
    }

    const type = tipos.find((item) => item.key === tipo) || tipos[0] || { actividades: [] };
    const block = (type.actividades || []).find((item) => Number(item.numero) === Number(actividad)) || {
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
    let layout = null;
    try { layout = _detectActivityLayout(unidad); } catch { layout = null; }

    return _activityDefinitions().map((definition) => {
      let actividades = [];
      try {
        if (layout && layout.tipoColMap[definition.key] !== undefined) {
          const colIdx = layout.tipoColMap[definition.key];
          const blocks = _getActivityBlocks(layout.rows, colIdx, layout.primeraFilaBloque, layout.alturaBloque);
          actividades = blocks.map((b) => ({ numero: b.numero, nombre: b.nombre, incluida: b.incluida }));
        } else {
          // Intentar hoja separada
          const splitSheet = _activitySheetName(unidad, definition.key);
          const splitRows = _rows(splitSheet);
          const headerIdx = splitRows.findIndex((row) =>
            (row || []).some((cell) => /^Act\d+/i.test(String(cell || "").trim()))
          );
          const header = splitRows[headerIdx >= 0 ? headerIdx : 0] || [];
          actividades = header
            .map((value) => {
              const match = String(value || "").match(/^Act(\d+)/i);
              return match ? { numero: Number(match[1]), nombre: "", incluida: true } : null;
            })
            .filter(Boolean);
        }
      } catch {
        actividades = [];
      }
      if (!actividades.length) actividades = [{ numero: 1, nombre: "", incluida: true }];
      return {
        ...definition,
        total: actividades.length,
        incluidas: actividades.filter((item) => item.incluida).length,
        actividades,
      };
    });
  }

  async function _saveNotasActividad({ unidad, tipo, actividad, notas, nombreActividad, incluida }) {
    const layout = _detectActivityLayout(unidad);
    const hoja = layout.sheetName;
    const ws = _sheet(hoja);
    const colIdx = _tipoColIdx(layout.tipoColMap, tipo);
    const def = _activityDefinition(tipo);
    const notaOffset = def.notaOffset !== undefined ? def.notaOffset : 4;
    const notaColIdx = colIdx + notaOffset;
    const blocks = _getActivityBlocks(layout.rows, colIdx, layout.primeraFilaBloque, layout.alturaBloque);
    const block = blocks.find((b) => Number(b.numero) === Number(actividad));
    if (!block) throw new Error(`Actividad ${actividad} no encontrada en ${hoja}`);

    function _writeAndRecord(r, c, v) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      if (v === "" || v === null || v === undefined) {
        ws[cellRef] = { v: "", t: "s" };
      } else if (typeof v === "number") {
        ws[cellRef] = { v, t: "n" };
      } else {
        ws[cellRef] = { v: String(v), t: "s" };
      }
      if (_isAndroid()) _recordPatch(hoja, r, c, v);
    }

    // Guardar nombre en fila N° col <nombre real> (colIdx+3: N° | num | NOMBRE | <nombre real>)
    if (nombreActividad !== undefined) {
      _writeAndRecord(block.filaInicio + 1, colIdx + 3, String(nombreActividad));
    }
    // Guardar incluida en fila INCLUIDO col colIdx+1
    if (incluida !== undefined) {
      _writeAndRecord(block.filaInicio + 2, colIdx + 1, incluida ? "x" : "");
    }
    // Guardar notas de alumnos
    (notas || []).forEach((n) => {
      if (n.rowIdx !== undefined) {
        const v = n.nota === "" ? "" : Number(n.nota);
        _writeAndRecord(n.rowIdx, notaColIdx, v);
      }
    });
    _clearRowsCache(hoja);
    await _downloadWorkbook();
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
      await _downloadWorkbook();
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
      await _downloadWorkbook();
      return { ok: true, ..._getRraaCriterios() };
    },

    saveAlumnos: async (alumnos) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      wb.Sheets["Alumnos"] = XLSX.utils.json_to_sheet(alumnos);
      _clearRowsCache("Alumnos");
      await _downloadWorkbook();
      return { ok: true, fileName: _fileName, alumnos };
    },

    getNotasActividad: async (payload) => {
      if (payload && payload.includeRraa === false) await _ensureSourceBuffer();
      else await _ensureWorkbook();
      return _getNotasActividad(payload);
    },
    getNotasActividadesTipo: async ({ unidad, tipo }) => {
      await _ensureSourceBuffer();
      const tipos = _getTiposActividad(unidad);
      const type = tipos.find(item => item.key === tipo) || tipos[0];
      const actividades = type ? type.actividades : [];
      const notas = [];
      for (const actividad of actividades) {
        notas.push({
          actividad: actividad.numero,
          nombre: actividad.nombre || "",
          notas: _getNotasActividad({ unidad, tipo, actividad: actividad.numero, includeRraa: false }).notas,
        });
      }
      return { fileName: _fileName, unidad, tipo, tipos, actividades, notas };
    },
    saveNotasActividad: async (payload) => {
      await _ensureWorkbook();
      await _saveNotasActividad(payload);
      return _getNotasActividad(payload);
    },

    saveCeNotas: async (payload) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      const hoja = `CE_${_unitSheetName(payload.unidad)}`;
      wb.Sheets[hoja] = XLSX.utils.json_to_sheet(payload.notas || []);
      if (!wb.SheetNames.includes(hoja)) wb.SheetNames.push(hoja);
      await _downloadWorkbook();
      return { ok: true };
    },

    addActividad: async (payload) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      const layout = _detectActivityLayout(payload.unidad);
      const hoja = layout.sheetName;
      const ws = wb.Sheets[hoja];
      if (!ws) throw new Error(`Hoja ${hoja} no encontrada`);
      const colIdx = _tipoColIdx(layout.tipoColMap, payload.tipo);
      const blocks = _getActivityBlocks(layout.rows, colIdx, layout.primeraFilaBloque, layout.alturaBloque);
      const numero = Number(payload.numero) || (Math.max(0, ...blocks.map((b) => Number(b.numero) || 0)) + 1);
      const def = _activityDefinition(payload.tipo);

      // Calcular dónde insertar el nuevo bloque (al final de los bloques existentes)
      let nuevaFilaInicio;
      if (blocks.length > 0) {
        const ultimoBloque = blocks[blocks.length - 1];
        nuevaFilaInicio = ultimoBloque.filaInicio + layout.alturaBloque;
      } else {
        nuevaFilaInicio = layout.primeraFilaBloque;
      }

      // Copiar estructura del último bloque o el primero como plantilla
      const plantilla = blocks.length > 0 ? blocks[blocks.length - 1] : blocks[0];
      if (!plantilla) throw new Error(`No hay bloques existentes para copiar en ${hoja}`);

      // Copiar bloque plantilla al nuevo lugar escribiendo celda a celda (no reescribir toda la hoja)
      const notaOffset = def.notaOffset !== undefined ? def.notaOffset : 4;
      const endCol = colIdx + Math.max(notaOffset + 1, 20);

      function _writeCell(ws, r, c, v) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        if (v === "" || v === undefined || v === null) {
          ws[cellRef] = { v: "", t: "s" };
        } else if (typeof v === "number") {
          ws[cellRef] = { v, t: "n" };
        } else {
          ws[cellRef] = { v: String(v), t: "s" };
        }
      }

      // Copiar estructura del bloque plantilla al nuevo lugar
      for (let offset = 0; offset < layout.alturaBloque; offset++) {
        const srcRow = layout.rows[plantilla.filaInicio + offset] || [];
        const dstRowIdx = nuevaFilaInicio + offset;
        for (let c = colIdx; c < Math.min(srcRow.length, endCol); c++) {
          _writeCell(ws, dstRowIdx, c, srcRow[c]);
        }
      }

      // Actualizar N° con el nuevo número
      _writeCell(ws, nuevaFilaInicio + 1, colIdx + 1, numero);
      // Actualizar INCLUIDO
      _writeCell(ws, nuevaFilaInicio + 2, colIdx + 1, payload.incluida !== false ? "x" : "");
      // Limpiar notas de alumnos
      for (let offset = 4; offset < layout.alturaBloque; offset++) {
        _writeCell(ws, nuevaFilaInicio + offset, colIdx + notaOffset, "");
      }

      // Actualizar !ref si es necesario
      const ref = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
      const newMaxRow = nuevaFilaInicio + layout.alturaBloque - 1;
      const newMaxCol = endCol - 1;
      if (newMaxRow > ref.e.r) ref.e.r = newMaxRow;
      if (newMaxCol > ref.e.c) ref.e.c = newMaxCol;
      ws["!ref"] = XLSX.utils.encode_range(ref);

      _clearRowsCache(hoja);
      await _downloadWorkbook();
      return _getNotasActividad({ unidad: payload.unidad, tipo: payload.tipo, actividad: numero, includeRraa: false });
    },

    getNotasUnidad: async (payload) => {
      await _ensureWorkbook();
      try { return _sheetToJson(`${_unitSheetName(payload.unidad)}_resumen`); }
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
      await _downloadWorkbook();
      return { ok: true };
    },
    deleteDiarioEntrada: async (payload) => {
      const wb = await _ensureWorkbook();
      if (!wb) throw new Error("Sin archivo");
      let rows = [];
      try { rows = _sheetToJson("Diario"); } catch {}
      rows = rows.filter(r => !(r.fecha === payload.fecha && r.texto === payload.texto));
      wb.Sheets["Diario"] = XLSX.utils.json_to_sheet(rows);
      await _downloadWorkbook();
      return { ok: true };
    },
  };

})();
