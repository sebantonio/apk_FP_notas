package com.apkesonotas;

import android.content.SharedPreferences;
import android.util.Log;
import org.json.JSONObject;

class NativeNoteStore {
  private static final String TAG = "NativeNoteStore";
  private static final String PATCHES = "native_cell_patches";
  private final SharedPreferences prefs;

  NativeNoteStore(SharedPreferences prefs) {
    this.prefs = prefs;
  }

  String get(String fileKey, String sheet, int row, int col, String fallback) {
    try {
      JSONObject all = new JSONObject(prefs.getString(PATCHES, "{}"));
      String key = key(fileKey, sheet, row, col);
      String val = all.has(key) ? all.getString(key) : fallback;
      Log.d(TAG, "GET " + key + " = " + val);
      return val;
    } catch (Exception ex) {
      Log.e(TAG, "GET error: " + ex.getMessage());
      return fallback;
    }
  }

  void put(String fileKey, String sheet, int row, int col, String value) {
    try {
      JSONObject all = new JSONObject(prefs.getString(PATCHES, "{}"));
      String key = key(fileKey, sheet, row, col);
      all.put(key, value == null ? "" : value);
      boolean ok = prefs.edit().putString(PATCHES, all.toString()).commit();
      Log.d(TAG, "PUT " + key + " = " + value + " -> " + ok);
    } catch (Exception ex) {
      Log.e(TAG, "PUT error: " + ex.getMessage());
    }
  }

  void putBatch(String fileKey, java.util.List<String[]> entries) {
    try {
      JSONObject all = new JSONObject(prefs.getString(PATCHES, "{}"));
      for (String[] e : entries) {
        String key = key(fileKey, e[0], Integer.parseInt(e[1]), Integer.parseInt(e[2]));
        all.put(key, e[3] == null ? "" : e[3]);
        Log.d(TAG, "BATCH PUT " + key + " = " + e[3]);
      }
      boolean ok = prefs.edit().putString(PATCHES, all.toString()).commit();
      Log.d(TAG, "BATCH commit -> " + ok + " total keys: " + all.length());
    } catch (Exception ex) {
      Log.e(TAG, "BATCH error: " + ex.getMessage());
    }
  }

  private String key(String fileKey, String sheet, int row, int col) {
    return fileKey + "|" + sheet + "|" + row + "|" + col;
  }
}
