package com.apkesonotas;

import android.content.SharedPreferences;
import org.json.JSONArray;
import org.json.JSONObject;

class NativeNoteStore {
  private static final String PATCHES = "native_cell_patches";
  private final SharedPreferences prefs;

  NativeNoteStore(SharedPreferences prefs) {
    this.prefs = prefs;
  }

  String get(String fileKey, String sheet, int row, int col, String fallback) {
    try {
      JSONObject all = new JSONObject(prefs.getString(PATCHES, "{}"));
      String key = key(fileKey, sheet, row, col);
      return all.has(key) ? all.getString(key) : fallback;
    } catch (Exception ex) {
      return fallback;
    }
  }

  void put(String fileKey, String sheet, int row, int col, String value) {
    try {
      JSONObject all = new JSONObject(prefs.getString(PATCHES, "{}"));
      all.put(key(fileKey, sheet, row, col), value == null ? "" : value);
      prefs.edit().putString(PATCHES, all.toString()).apply();
    } catch (Exception ignored) {}
  }

  void putMany(String fileKey, JSONArray patches) {
    try {
      JSONObject all = new JSONObject(prefs.getString(PATCHES, "{}"));
      for (int i = 0; i < patches.length(); i++) {
        JSONObject p = patches.getJSONObject(i);
        all.put(key(fileKey, p.getString("sheet"), p.getInt("row"), p.getInt("col")), p.optString("value", ""));
      }
      prefs.edit().putString(PATCHES, all.toString()).apply();
    } catch (Exception ignored) {}
  }

  private String key(String fileKey, String sheet, int row, int col) {
    return fileKey + "|" + sheet + "|" + row + "|" + col;
  }
}
