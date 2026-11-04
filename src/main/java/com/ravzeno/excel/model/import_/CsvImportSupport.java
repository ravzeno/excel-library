package com.ravzeno.excel.model.import_;

import java.io.*;
import java.nio.charset.Charset;
import java.util.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Internal CSV decoder shared by the annotated import service. */
public final class CsvImportSupport {
  private CsvImportSupport() {}

  public static byte[] toXlsx(InputStream input, Charset charset) throws IOException {
    InputStream callerOwnedInput =
        new FilterInputStream(input) {
          @Override
          public void close() {
            // The caller owns the original input stream.
          }
        };
    try (Reader reader = new BufferedReader(new InputStreamReader(callerOwnedInput, charset));
        XSSFWorkbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream output = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet("CSV");
      PushbackReader csv = new PushbackReader(reader, 1);
      List<String> row = new ArrayList<>();
      StringBuilder cell = new StringBuilder();
      boolean quoted = false;
      int rowNo = 0, ch;
      while ((ch = csv.read()) != -1) {
        char c = (char) ch;
        if (rowNo == 0 && row.isEmpty() && cell.isEmpty() && c == '\uFEFF') continue;
        if (quoted) {
          if (c == '"') {
            int next = csv.read();
            if (next == '"') cell.append('"');
            else {
              quoted = false;
              if (next != -1) csv.unread(next);
            }
          } else cell.append(c);
        } else if (c == '"' && cell.length() == 0) quoted = true;
        else if (c == ',') {
          row.add(cell.toString());
          cell.setLength(0);
        } else if (c == '\n' || c == '\r') {
          if (c == '\r') {
            int next = csv.read();
            if (next != '\n' && next != -1) csv.unread(next);
          }
          row.add(cell.toString());
          cell.setLength(0);
          writeRow(sheet, rowNo++, row);
          row = new ArrayList<>();
        } else cell.append(c);
      }
      if (quoted) throw new IOException("Unterminated quoted CSV field");
      if (cell.length() > 0 || !row.isEmpty()) {
        row.add(cell.toString());
        writeRow(sheet, rowNo, row);
      }
      workbook.write(output);
      return output.toByteArray();
    }
  }

  private static void writeRow(
      org.apache.poi.ss.usermodel.Sheet sheet, int rowNo, List<String> values) {
    var row = sheet.createRow(rowNo);
    for (int i = 0; i < values.size(); i++) row.createCell(i).setCellValue(values.get(i));
  }
}
