package com.apkesonotas;

import android.app.Activity;
import android.content.Intent;
import android.content.SharedPreferences;
import android.net.Uri;
import android.os.Bundle;
import android.view.Gravity;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.Spinner;
import android.widget.TextView;
import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AppCompatActivity;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity {
  private static final String PREFS = "native_fp_notas";
  private static final String KEY_URI = "excel_uri";
  private static final String KEY_NAME = "excel_name";

  private SharedPreferences prefs;
  private NativeWorkbook workbook;
  private NativeNoteStore noteStore;
  private Uri excelUri;

  private TextView status;
  private TextView fileName;
  private Spinner unitSpinner;
  private Spinner typeSpinner;
  private Spinner activitySpinner;
  private EditText activityName;
  private CheckBox included;
  private LinearLayout notesList;

  private final List<NativeWorkbook.ActivityType> types = NativeWorkbook.defaultTypes();
  private List<String> units = new ArrayList<>();
  private List<NativeWorkbook.ActivityBlock> activities = new ArrayList<>();
  private final List<NativeWorkbook.StudentNote> notes = new ArrayList<>();

  private final ActivityResultLauncher<Intent> picker = registerForActivityResult(
    new ActivityResultContracts.StartActivityForResult(),
    result -> {
      if (result.getResultCode() != Activity.RESULT_OK || result.getData() == null || result.getData().getData() == null) return;
      Uri uri = result.getData().getData();
      int flags = result.getData().getFlags() & (Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
      try { getContentResolver().takePersistableUriPermission(uri, flags); } catch (Exception ignored) {}
      prefs.edit().putString(KEY_URI, uri.toString()).putString(KEY_NAME, uri.getLastPathSegment()).apply();
      loadWorkbook(uri);
    }
  );

  @Override
  protected void onCreate(Bundle savedInstanceState) {
    super.onCreate(savedInstanceState);
    prefs = getSharedPreferences(PREFS, MODE_PRIVATE);
    noteStore = new NativeNoteStore(prefs);
    buildUi();
    String savedUri = prefs.getString(KEY_URI, null);
    if (savedUri != null) loadWorkbook(Uri.parse(savedUri));
    else setStatus("Selecciona un Excel para empezar.");
  }

  private void buildUi() {
    ScrollView rootScroll = new ScrollView(this);
    LinearLayout root = new LinearLayout(this);
    root.setOrientation(LinearLayout.VERTICAL);
    root.setPadding(28, 28, 28, 28);
    rootScroll.addView(root);

    TextView title = new TextView(this);
    title.setText("FP Notas nativo");
    title.setTextSize(26);
    title.setGravity(Gravity.CENTER_VERTICAL);
    root.addView(title);

    fileName = new TextView(this);
    fileName.setText("Sin archivo");
    fileName.setPadding(0, 14, 0, 14);
    root.addView(fileName);

    Button select = new Button(this);
    select.setText("Seleccionar Excel");
    select.setOnClickListener(v -> openPicker());
    root.addView(select);

    unitSpinner = addSpinner(root, "Unidad");
    typeSpinner = addSpinner(root, "Tipo");
    activitySpinner = addSpinner(root, "Actividad");

    activityName = new EditText(this);
    activityName.setHint("Nombre de actividad");
    root.addView(activityName);

    included = new CheckBox(this);
    included.setText("Incluida");
    root.addView(included);

    Button save = new Button(this);
    save.setText("Guardar notas y nombre");
    save.setOnClickListener(v -> saveCurrent());
    root.addView(save);

    status = new TextView(this);
    status.setPadding(0, 14, 0, 14);
    root.addView(status);

    notesList = new LinearLayout(this);
    notesList.setOrientation(LinearLayout.VERTICAL);
    root.addView(notesList);

    AdapterView.OnItemSelectedListener reloadListener = new AdapterView.OnItemSelectedListener() {
      @Override public void onItemSelected(AdapterView<?> parent, View view, int position, long id) { reloadFromSelection(); }
      @Override public void onNothingSelected(AdapterView<?> parent) {}
    };
    unitSpinner.setOnItemSelectedListener(reloadListener);
    typeSpinner.setOnItemSelectedListener(reloadListener);
    activitySpinner.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
      @Override public void onItemSelected(AdapterView<?> parent, View view, int position, long id) { loadSelectedActivity(); }
      @Override public void onNothingSelected(AdapterView<?> parent) {}
    });

    setContentView(rootScroll);
  }

  private Spinner addSpinner(LinearLayout root, String label) {
    TextView tv = new TextView(this);
    tv.setText(label);
    tv.setPadding(0, 16, 0, 4);
    root.addView(tv);
    Spinner spinner = new Spinner(this);
    root.addView(spinner);
    return spinner;
  }

  private void openPicker() {
    Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
    intent.addCategory(Intent.CATEGORY_OPENABLE);
    intent.setType("*/*");
    intent.putExtra(Intent.EXTRA_MIME_TYPES, new String[] {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
      "application/octet-stream"
    });
    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION | Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION);
    picker.launch(intent);
  }

  private void loadWorkbook(Uri uri) {
    try {
      excelUri = uri;
      workbook = NativeWorkbook.open(this, uri, noteStore);
      units = workbook.getUnitSheetNames();
      fileName.setText(prefs.getString(KEY_NAME, uri.toString()));
      setAdapter(unitSpinner, units);
      setAdapter(typeSpinner, NativeWorkbook.typeLabels(types));
      reloadFromSelection();
      setStatus("Excel cargado. Cambios guardados en almacenamiento nativo.");
    } catch (Exception ex) {
      setStatus("No se pudo leer el Excel: " + ex.getMessage());
    }
  }

  private void reloadFromSelection() {
    if (workbook == null || units.isEmpty() || typeSpinner.getSelectedItemPosition() < 0) return;
    String unit = selectedUnit();
    NativeWorkbook.ActivityType type = selectedType();
    activities = workbook.getActivities(unit, type);
    List<String> labels = new ArrayList<>();
    for (NativeWorkbook.ActivityBlock b : activities) labels.add(b.number + (b.name.isEmpty() ? "" : " - " + b.name));
    if (labels.isEmpty()) labels.add("1");
    setAdapter(activitySpinner, labels);
    loadSelectedActivity();
  }

  private void loadSelectedActivity() {
    if (workbook == null || activities.isEmpty()) return;
    NativeWorkbook.ActivityBlock block = selectedBlock();
    activityName.setText(block.name);
    included.setChecked(block.included);
    notes.clear();
    notes.addAll(workbook.getNotes(selectedUnit(), selectedType(), block));
    renderNotes();
  }

  private void renderNotes() {
    notesList.removeAllViews();
    for (NativeWorkbook.StudentNote note : notes) {
      LinearLayout row = new LinearLayout(this);
      row.setOrientation(LinearLayout.HORIZONTAL);
      row.setGravity(Gravity.CENTER_VERTICAL);
      TextView name = new TextView(this);
      name.setText(note.number + ". " + note.name);
      row.addView(name, new LinearLayout.LayoutParams(0, LinearLayout.LayoutParams.WRAP_CONTENT, 1));
      EditText grade = new EditText(this);
      grade.setSingleLine(true);
      grade.setText(note.grade);
      grade.setInputType(android.text.InputType.TYPE_CLASS_NUMBER | android.text.InputType.TYPE_NUMBER_FLAG_DECIMAL);
      grade.setOnFocusChangeListener((v, hasFocus) -> { if (!hasFocus) note.grade = ((EditText) v).getText().toString(); });
      row.addView(grade, new LinearLayout.LayoutParams(220, LinearLayout.LayoutParams.WRAP_CONTENT));
      notesList.addView(row);
    }
  }

  private void saveCurrent() {
    if (workbook == null || excelUri == null || activities.isEmpty()) return;
    for (int i = 0; i < notesList.getChildCount(); i++) {
      LinearLayout row = (LinearLayout) notesList.getChildAt(i);
      EditText grade = (EditText) row.getChildAt(1);
      notes.get(i).grade = grade.getText().toString();
    }
    NativeWorkbook.ActivityBlock block = selectedBlock();
    block.name = activityName.getText().toString();
    block.included = included.isChecked();
    workbook.saveActivity(selectedUnit(), selectedType(), block, notes);
    setStatus("Guardado nativo: " + block.name + " (" + notes.size() + " notas).");
    reloadFromSelection();
  }

  private String selectedUnit() {
    return units.get(Math.max(0, unitSpinner.getSelectedItemPosition()));
  }

  private NativeWorkbook.ActivityType selectedType() {
    return types.get(Math.max(0, typeSpinner.getSelectedItemPosition()));
  }

  private NativeWorkbook.ActivityBlock selectedBlock() {
    int idx = Math.max(0, Math.min(activitySpinner.getSelectedItemPosition(), activities.size() - 1));
    return activities.get(idx);
  }

  private void setAdapter(Spinner spinner, List<String> items) {
    ArrayAdapter<String> adapter = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, items);
    adapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
    spinner.setAdapter(adapter);
  }

  private void setStatus(String text) {
    status.setText(text);
  }
}
