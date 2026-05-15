package com.apkesonotas;

import android.os.Bundle;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {
  @Override
  protected void onCreate(Bundle savedInstanceState) {
    registerPlugin(ExcelFilePlugin.class);
    super.onCreate(savedInstanceState);
  }
}
