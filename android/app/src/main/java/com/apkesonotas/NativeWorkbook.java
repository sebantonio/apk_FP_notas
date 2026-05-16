package com.apkesonotas;

import android.content.Context;
import android.net.Uri;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import android.util.Log;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

class NativeWorkbook {
  private final String fileKey;
  private final NativeNoteStore store;
  private final Map<String, String> sheetEntries;
  private final Map<String, List<List<String>>> rowsCache = new HashMap<>();
  private final List<String> sharedStrings;
  private final byte[] xlsxBytes;

  static NativeWorkbook open(Context context, Uri uri, NativeNoteStore store) throws Exception {
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    try (InputStream in = context.getContentResolver().openInputStream(uri)) {
      if (in == null) throw new IllegalStateException("No se pudo abrir el archivo");
      byte[] buffer = new byte[65536];
      int read;
      while ((read = in.read(buffer)) != -1) out.write(buffer, 0, read);
    }
    byte[] bytes = out.toByteArray();
    Map<String, byte[]> entries = unzip(bytes);
    List<String> shared = parseSharedStrings(entries.get("xl/sharedStrings.xml"));
    Map<String, String> sheets = parseSheets(entries);
    String fileKey = uri.getLastPathSegment();
    if (fileKey == null || fileKey.isEmpty()) fileKey = uri.toString();
    Log.d("NativeWorkbook", "fileKey=" + fileKey);
    return new NativeWorkbook(fileKey, store, sheets, shared, bytes);
  }

  private NativeWorkbook(String fileKey, NativeNoteStore store, Map<String, String> sheetEntries, List<String> sharedStrings, byte[] xlsxBytes) {
    this.fileKey = fileKey;
    this.store = store;
    this.sheetEntries = sheetEntries;
    this.sharedStrings = sharedStrings;
    this.xlsxBytes = xlsxBytes;
  }

  List<String> getUnitSheetNames() {
    List<String> units = new ArrayList<>();
    for (String name : sheetEntries.keySet()) {
      if (name.toUpperCase(Locale.ROOT).matches("U\\d+")) units.add(name);
    }
    if (units.isEmpty()) units.addAll(sheetEntries.keySet());
    return units;
  }

  List<ActivityBlock> getActivities(String unit, ActivityType type) {
    Layout layout = detectLayout(unit);
    int col = layout.typeCol.containsKey(type.key) ? layout.typeCol.get(type.key) : 0;
    List<ActivityBlock> blocks = getBlocks(unit, layout, type, col);
    if (blocks.isEmpty()) blocks.add(new ActivityBlock(1, "", true, layout.firstRow));
    return blocks;
  }

  List<StudentNote> getNotes(String unit, ActivityType type, ActivityBlock block) {
    Layout layout = detectLayout(unit);
    int col = layout.typeCol.containsKey(type.key) ? layout.typeCol.get(type.key) : 0;
    int noteCol = col + type.noteOffset;
    List<List<String>> rows = rows(unit);
    List<StudentNote> result = new ArrayList<>();
    for (int r = block.startRow + 4; r < Math.min(rows.size(), block.startRow + layout.blockHeight); r++) {
      String name = cell(unit, rows, r, col);
      if (name.trim().isEmpty()) continue;
      String grade = cell(unit, rows, r, noteCol);
      result.add(new StudentNote(result.size() + 1, name, grade, r, noteCol));
    }
    return result;
  }

  void saveActivity(String unit, ActivityType type, ActivityBlock block, List<StudentNote> notes) {
    Layout layout = detectLayout(unit);
    int col = layout.typeCol.containsKey(type.key) ? layout.typeCol.get(type.key) : 0;
    int nameCol = col + nameOffset(unit, block.startRow, col);
    int includedCol = col + 1;
    List<String[]> entries = new ArrayList<>();
    entries.add(new String[]{unit, String.valueOf(block.startRow + 1), String.valueOf(nameCol), block.name});
    entries.add(new String[]{unit, String.valueOf(block.startRow + 2), String.valueOf(includedCol), block.included ? "x" : ""});
    for (StudentNote note : notes) {
      Log.d("NativeWorkbook", "saving note row=" + note.row + " col=" + note.col + " grade=" + note.grade);
      entries.add(new String[]{unit, String.valueOf(note.row), String.valueOf(note.col), note.grade});
    }
    store.putBatch(fileKey, entries);
    rowsCache.remove(unit);
  }

  private List<ActivityBlock> getBlocks(String unit, Layout layout, ActivityType type, int col) {
    List<ActivityBlock> blocks = new ArrayList<>();
    List<List<String>> rows = rows(unit);
    for (int start = layout.firstRow; start < rows.size(); start += layout.blockHeight) {
      String title = cell(unit, rows, start, col);
      if (title.trim().isEmpty()) break;
      int number = parseInt(cell(unit, rows, start + 1, col + 1), blocks.size() + 1);
      String name = cell(unit, rows, start + 1, col + nameOffset(unit, start, col));
      String inc = cell(unit, rows, start + 2, col + 1).trim().toUpperCase(Locale.ROOT);
      blocks.add(new ActivityBlock(number, name, inc.equals("X") || inc.equals("SI") || inc.equals("S"), start));
    }
    return blocks;
  }

  private Layout detectLayout(String unit) {
    List<List<String>> rows = rows(unit);
    Map<String, Integer> typeCol = new HashMap<>();
    for (int r = 0; r < rows.size(); r++) {
      String first = raw(rows, r, 0);
      if (!matches(first, defaultTypes().get(0))) continue;
      String next = raw(rows, r + 1, 0).trim().toUpperCase(Locale.ROOT);
      if (!next.startsWith("N")) continue;
      typeCol.put("practicas", 0);
      for (int c = 1; c < rows.get(r).size(); c++) {
        String text = raw(rows, r, c);
        for (ActivityType type : defaultTypes()) {
          if (!typeCol.containsKey(type.key) && matches(text, type)) typeCol.put(type.key, c);
        }
      }
      int height = 44;
      for (int r2 = r + 1; r2 < Math.min(rows.size(), r + 100); r2++) {
        if (matches(raw(rows, r2, 0), defaultTypes().get(0))) {
          height = r2 - r;
          break;
        }
      }
      return new Layout(r, height, typeCol);
    }
    typeCol.put("practicas", 0);
    return new Layout(0, 44, typeCol);
  }

  private int nameOffset(String unit, int row, int col) {
    List<List<String>> rows = rows(unit);
    int offset = 3;
    for (int c = col; c < col + 10; c++) {
      String text = raw(rows, row + 1, c).trim().toUpperCase(Locale.ROOT);
      if (text.equals("NOMBRE") || text.equals("NOMBRE ACTIVIDAD") || text.equals("ACT.") || text.equals("ACTIVIDAD")) {
        offset = c - col + 1;
        break;
      }
    }
    return offset;
  }

  private String cell(String sheet, List<List<String>> rows, int row, int col) {
    return store.get(fileKey, sheet, row, col, raw(rows, row, col));
  }

  private String raw(List<List<String>> rows, int row, int col) {
    if (row < 0 || row >= rows.size()) return "";
    List<String> values = rows.get(row);
    if (col < 0 || col >= values.size()) return "";
    return values.get(col) == null ? "" : values.get(col);
  }

  private List<List<String>> rows(String sheetName) {
    if (rowsCache.containsKey(sheetName)) return rowsCache.get(sheetName);
    try {
      Map<String, byte[]> entries = unzip(xlsxBytes);
      List<List<String>> rows = parseSheet(entries.get(sheetEntries.get(sheetName)), sharedStrings);
      rowsCache.put(sheetName, rows);
      return rows;
    } catch (Exception ex) {
      return Collections.emptyList();
    }
  }

  private boolean matches(String text, ActivityType type) {
    String upper = text == null ? "" : text.toUpperCase(Locale.ROOT);
    for (String term : type.matches) if (upper.contains(term)) return true;
    return false;
  }

  private int parseInt(String value, int fallback) {
    try { return (int) Double.parseDouble(value.replace(",", ".")); }
    catch (Exception ex) { return fallback; }
  }

  static List<ActivityType> defaultTypes() {
    List<ActivityType> list = new ArrayList<>();
    list.add(new ActivityType("practicas", "Practicas", 4, "PRACTICA", "PRÁCTICA", "PRACTICAS", "PRÁCTICAS"));
    list.add(new ActivityType("memorias", "Memorias", 3, "MEMORIA", "MEMORIAS"));
    list.add(new ActivityType("otros", "Otras actividades", 3, "OTROS", "O.I", "O.INSTRU"));
    list.add(new ActivityType("controles", "Controles", 3, "CONTROL", "CONTROLES", "PRUEBA"));
    return list;
  }

  static List<String> typeLabels(List<ActivityType> types) {
    List<String> labels = new ArrayList<>();
    for (ActivityType type : types) labels.add(type.label);
    return labels;
  }

  private static Map<String, byte[]> unzip(byte[] bytes) throws Exception {
    Map<String, byte[]> map = new HashMap<>();
    try (ZipInputStream zip = new ZipInputStream(new ByteArrayInputStream(bytes))) {
      ZipEntry entry;
      byte[] buffer = new byte[65536];
      while ((entry = zip.getNextEntry()) != null) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        int read;
        while ((read = zip.read(buffer)) != -1) out.write(buffer, 0, read);
        map.put(entry.getName(), out.toByteArray());
      }
    }
    return map;
  }

  private static List<String> parseSharedStrings(byte[] xml) throws Exception {
    List<String> strings = new ArrayList<>();
    if (xml == null) return strings;
    Document doc = parse(xml);
    NodeList sis = doc.getElementsByTagName("si");
    for (int i = 0; i < sis.getLength(); i++) strings.add(sis.item(i).getTextContent());
    return strings;
  }

  private static Map<String, String> parseSheets(Map<String, byte[]> entries) throws Exception {
    Map<String, String> rels = new HashMap<>();
    Document relDoc = parse(entries.get("xl/_rels/workbook.xml.rels"));
    NodeList relNodes = relDoc.getElementsByTagName("Relationship");
    for (int i = 0; i < relNodes.getLength(); i++) {
      Element e = (Element) relNodes.item(i);
      rels.put(e.getAttribute("Id"), normalizeTarget(e.getAttribute("Target")));
    }
    Map<String, String> sheets = new LinkedHashMap<>();
    Document wb = parse(entries.get("xl/workbook.xml"));
    NodeList sheetNodes = wb.getElementsByTagName("sheet");
    for (int i = 0; i < sheetNodes.getLength(); i++) {
      Element e = (Element) sheetNodes.item(i);
      String id = e.getAttribute("r:id");
      if (id.isEmpty()) id = e.getAttribute("id");
      String target = rels.get(id);
      if (target != null) sheets.put(e.getAttribute("name"), target);
    }
    return sheets;
  }

  private static String normalizeTarget(String target) {
    if (target.startsWith("/")) return target.substring(1);
    if (target.startsWith("xl/")) return target;
    return "xl/" + target;
  }

  private static List<List<String>> parseSheet(byte[] xml, List<String> shared) throws Exception {
    List<List<String>> rows = new ArrayList<>();
    if (xml == null) return rows;
    Document doc = parse(xml);
    NodeList cells = doc.getElementsByTagName("c");
    Pattern refPattern = Pattern.compile("([A-Z]+)(\\d+)");
    for (int i = 0; i < cells.getLength(); i++) {
      Element c = (Element) cells.item(i);
      Matcher m = refPattern.matcher(c.getAttribute("r"));
      if (!m.matches()) continue;
      int col = colIndex(m.group(1));
      int row = Integer.parseInt(m.group(2)) - 1;
      while (rows.size() <= row) rows.add(new ArrayList<>());
      List<String> values = rows.get(row);
      while (values.size() <= col) values.add("");
      String type = c.getAttribute("t");
      String value = firstText(c, "v");
      if ("s".equals(type)) value = sharedValue(shared, value);
      else if ("inlineStr".equals(type)) value = c.getTextContent();
      values.set(col, value == null ? "" : value);
    }
    return rows;
  }

  private static Document parse(byte[] xml) throws Exception {
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    factory.setNamespaceAware(false);
    return factory.newDocumentBuilder().parse(new ByteArrayInputStream(xml));
  }

  private static String firstText(Element e, String tag) {
    NodeList nodes = e.getElementsByTagName(tag);
    if (nodes.getLength() == 0) return "";
    Node node = nodes.item(0);
    return node == null ? "" : node.getTextContent();
  }

  private static String sharedValue(List<String> shared, String index) {
    try {
      int i = Integer.parseInt(index);
      return i >= 0 && i < shared.size() ? shared.get(i) : "";
    } catch (Exception ex) {
      return "";
    }
  }

  private static int colIndex(String letters) {
    int result = 0;
    for (int i = 0; i < letters.length(); i++) result = result * 26 + (letters.charAt(i) - 'A' + 1);
    return result - 1;
  }

  static class ActivityType {
    final String key;
    final String label;
    final int noteOffset;
    final String[] matches;
    ActivityType(String key, String label, int noteOffset, String... matches) {
      this.key = key;
      this.label = label;
      this.noteOffset = noteOffset;
      this.matches = matches;
    }
  }

  static class ActivityBlock {
    final int number;
    String name;
    boolean included;
    final int startRow;
    ActivityBlock(int number, String name, boolean included, int startRow) {
      this.number = number;
      this.name = name == null ? "" : name;
      this.included = included;
      this.startRow = startRow;
    }
  }

  static class StudentNote {
    final int number;
    final String name;
    String grade;
    final int row;
    final int col;
    StudentNote(int number, String name, String grade, int row, int col) {
      this.number = number;
      this.name = name;
      this.grade = grade == null ? "" : grade;
      this.row = row;
      this.col = col;
    }
  }

  private static class Layout {
    final int firstRow;
    final int blockHeight;
    final Map<String, Integer> typeCol;
    Layout(int firstRow, int blockHeight, Map<String, Integer> typeCol) {
      this.firstRow = firstRow;
      this.blockHeight = blockHeight;
      this.typeCol = typeCol;
    }
  }
}
