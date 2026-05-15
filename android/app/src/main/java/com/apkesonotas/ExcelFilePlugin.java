package com.apkesonotas;

import android.app.Activity;
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

  @PluginMethod
  public void selectFile(PluginCall call) {
    Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
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
    startActivityForResult(call, intent, "selectFileResult");
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
    resolveWithFile(call, uri, fileName);
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
    resolveWithFile(call, uri, fileName);
  }

  @PluginMethod
  public void saveFile(PluginCall call) {
    String base64 = call.getString("base64");
    String uriValue = call.getString("uri", getPrefs().getString(KEY_URI, null));
    if (base64 == null || base64.isEmpty()) {
      call.reject("No hay datos para guardar.");
      return;
    }
    if (uriValue == null || uriValue.isEmpty()) {
      call.reject("No hay archivo Excel seleccionado.");
      return;
    }

    Uri uri = Uri.parse(uriValue);
    try (OutputStream output = getContext().getContentResolver().openOutputStream(uri, "wt")) {
      if (output == null) {
        call.reject("No se pudo abrir el Excel para escritura.");
        return;
      }
      output.write(Base64.decode(base64, Base64.DEFAULT));
      output.flush();
      JSObject ret = new JSObject();
      ret.put("ok", true);
      ret.put("uri", uri.toString());
      call.resolve(ret);
    } catch (Exception ex) {
      call.reject("No se pudo guardar el Excel: " + ex.getMessage(), ex);
    }
  }

  private void resolveWithFile(PluginCall call, Uri uri, String fileName) {
    try {
      JSObject ret = new JSObject();
      ret.put("uri", uri.toString());
      ret.put("fileName", fileName);
      ret.put("base64", readBase64(uri));
      call.resolve(ret);
    } catch (Exception ex) {
      call.reject("No se pudo leer el Excel: " + ex.getMessage(), ex);
    }
  }

  private String readBase64(Uri uri) throws Exception {
    try (InputStream input = getContext().getContentResolver().openInputStream(uri);
         ByteArrayOutputStream output = new ByteArrayOutputStream()) {
      if (input == null) throw new IllegalStateException("Entrada de archivo no disponible");
      byte[] buffer = new byte[1024 * 64];
      int read;
      while ((read = input.read(buffer)) != -1) {
        output.write(buffer, 0, read);
      }
      return Base64.encodeToString(output.toByteArray(), Base64.NO_WRAP);
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
}
