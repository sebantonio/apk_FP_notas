package com.apkesonotas;

import android.app.Activity;
import android.content.ActivityNotFoundException;
import android.content.ContentResolver;
import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.database.Cursor;
import android.net.Uri;
import android.provider.OpenableColumns;
import android.util.Base64;
import androidx.activity.result.ActivityResult;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.ActivityCallback;
import com.getcapacitor.annotation.CapacitorPlugin;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

@CapacitorPlugin(name = "ExcelFile")
public class ExcelFilePlugin extends Plugin {
  private static final String PREFS_NAME = "excel_file";
  private static final String KEY_URI = "uri";
  private static final String KEY_FILE_NAME = "file_name";
  private OutputStream writeStream = null;

  @PluginMethod
  public void selectFile(PluginCall call) {
    Intent intent = buildFileIntent(Intent.ACTION_OPEN_DOCUMENT);
    try {
      startActivityForResult(call, intent, "selectFileResult");
    } catch (ActivityNotFoundException ex) {
      Intent fallback = buildFileIntent(Intent.ACTION_GET_CONTENT);
      try {
        startActivityForResult(call, fallback, "selectFileResult");
      } catch (ActivityNotFoundException fallbackEx) {
        call.reject("BlueStacks/Android no tiene un selector de archivos disponible.", fallbackEx);
      }
    }
  }

  private Intent buildFileIntent(String action) {
    Intent intent = new Intent(action);
    intent.addCategory(Intent.CATEGORY_OPENABLE);
    intent.setType("*/*");
    intent.putExtra(Intent.EXTRA_MIME_TYPES, new String[] {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
      "application/octet-stream"
    });
    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
    intent.addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
    intent.addFlags(Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION);
    return intent;
  }

  @ActivityCallback
  private void selectFileResult(PluginCall call, ActivityResult result) {
    if (call == null) return;
    if (result.getResultCode() != Activity.RESULT_OK || result.getData() == null || result.getData().getData() == null) {
      JSObject ret = new JSObject();
      ret.put("cancelled", true);
      call.resolve(ret);
      return;
    }

    Uri uri = result.getData().getData();
    int flags = result.getData().getFlags() & (Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
    try {
      getContext().getContentResolver().takePersistableUriPermission(uri, flags);
    } catch (Exception ignored) {
      // Some providers grant temporary access only. Reading still works in this session.
    }

    String fileName = getDisplayName(uri);
    getPrefs().edit().putString(KEY_URI, uri.toString()).putString(KEY_FILE_NAME, fileName).apply();
    resolveWithInfo(call, uri, fileName);
  }

  @PluginMethod
  public void readSelectedFile(PluginCall call) {
    String uriValue = getPrefs().getString(KEY_URI, null);
    if (uriValue == null || uriValue.isEmpty()) {
      call.resolve(new JSObject());
      return;
    }
    Uri uri = Uri.parse(uriValue);
    String fileName = getPrefs().getString(KEY_FILE_NAME, getDisplayName(uri));
    resolveWithInfo(call, uri, fileName);
  }

  @PluginMethod
  public void readFileChunk(PluginCall call) {
    String uriValue = call.getString("uri", getPrefs().getString(KEY_URI, null));
    int offset = call.getInt("offset", 0);
    int length = call.getInt("length", 262144);
    if (uriValue == null || uriValue.isEmpty()) {
      call.reject("No hay archivo Excel seleccionado.");
      return;
    }

    Uri uri = Uri.parse(uriValue);
    try (InputStream input = getContext().getContentResolver().openInputStream(uri)) {
      if (input == null) {
        call.reject("No se pudo abrir el Excel para lectura.");
        return;
      }
      long skipped = 0;
      while (skipped < offset) {
        long step = input.skip(offset - skipped);
        if (step <= 0) break;
        skipped += step;
      }
      byte[] buffer = new byte[length];
      int read = input.read(buffer);
      JSObject ret = new JSObject();
      if (read <= 0) {
        ret.put("base64", "");
        ret.put("bytesRead", 0);
        ret.put("eof", true);
      } else {
        ret.put("base64", Base64.encodeToString(buffer, 0, read, Base64.NO_WRAP));
        ret.put("bytesRead", read);
        ret.put("eof", read < length);
      }
      call.resolve(ret);
    } catch (Exception ex) {
      call.reject("No se pudo leer el Excel: " + ex.getMessage(), ex);
    }
  }

  @PluginMethod
  public void beginWrite(PluginCall call) {
    String uriValue = call.getString("uri", getPrefs().getString(KEY_URI, null));
    if (uriValue == null || uriValue.isEmpty()) {
      call.reject("No hay archivo Excel seleccionado.");
      return;
    }

    closeWriteStream();
    try {
      writeStream = getContext().getContentResolver().openOutputStream(Uri.parse(uriValue), "wt");
      if (writeStream == null) {
        call.reject("No se pudo abrir el Excel para escritura.");
        return;
      }
      JSObject ret = new JSObject();
      ret.put("ok", true);
      call.resolve(ret);
    } catch (Exception ex) {
      call.reject("No se pudo iniciar la escritura del Excel: " + ex.getMessage(), ex);
    }
  }

  @PluginMethod
  public void writeChunk(PluginCall call) {
    String base64 = call.getString("base64");
    if (writeStream == null) {
      call.reject("No hay una escritura iniciada.");
      return;
    }
    if (base64 == null || base64.isEmpty()) {
      call.reject("Bloque de datos vacío.");
      return;
    }

    try {
      writeStream.write(Base64.decode(base64, Base64.DEFAULT));
      JSObject ret = new JSObject();
      ret.put("ok", true);
      call.resolve(ret);
    } catch (Exception ex) {
      closeWriteStream();
      call.reject("No se pudo escribir el Excel: " + ex.getMessage(), ex);
    }
  }

  @PluginMethod
  public void finishWrite(PluginCall call) {
    try {
      if (writeStream != null) writeStream.flush();
      closeWriteStream();
      JSObject ret = new JSObject();
      ret.put("ok", true);
      call.resolve(ret);
    } catch (Exception ex) {
      closeWriteStream();
      call.reject("No se pudo cerrar el Excel: " + ex.getMessage(), ex);
    }
  }

  private void resolveWithInfo(PluginCall call, Uri uri, String fileName) {
    try {
      JSObject ret = new JSObject();
      ret.put("uri", uri.toString());
      ret.put("fileName", fileName);
      ret.put("size", getSize(uri));
      call.resolve(ret);
    } catch (Exception ex) {
      call.reject("No se pudo preparar el Excel: " + ex.getMessage(), ex);
    }
  }

  private String getDisplayName(Uri uri) {
    ContentResolver resolver = getContext().getContentResolver();
    try (Cursor cursor = resolver.query(uri, null, null, null, null)) {
      if (cursor != null && cursor.moveToFirst()) {
        int index = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME);
        if (index >= 0) {
          String name = cursor.getString(index);
          if (name != null && !name.isEmpty()) return name;
        }
      }
    } catch (Exception ignored) {}
    String fallback = uri.getLastPathSegment();
    return fallback == null || fallback.isEmpty() ? "archivo.xlsx" : fallback;
  }

  private SharedPreferences getPrefs() {
    return getContext().getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE);
  }

  private long getSize(Uri uri) {
    ContentResolver resolver = getContext().getContentResolver();
    try (Cursor cursor = resolver.query(uri, null, null, null, null)) {
      if (cursor != null && cursor.moveToFirst()) {
        int index = cursor.getColumnIndex(OpenableColumns.SIZE);
        if (index >= 0) return cursor.getLong(index);
      }
    } catch (Exception ignored) {}
    return -1;
  }

  private void closeWriteStream() {
    if (writeStream == null) return;
    try {
      writeStream.close();
    } catch (Exception ignored) {}
    writeStream = null;
  }
}
