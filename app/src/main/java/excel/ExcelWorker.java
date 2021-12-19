package excel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWorker {
  private static String[] months = new String[] {"Januar", "Februar", "März", "April", "Mai",
      "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"};

  Map<String, List<Integer>> map = new HashMap<>();
  Map<String, Beruf> idToBeruf = new HashMap<>();
  List<Beruf> berufe = new ArrayList<>();
  List<Integer> totals = new ArrayList<>();
  List<Integer> keineAngabe = new ArrayList<>();

  public void start() throws IOException {
    FileInputStream fis =
        new FileInputStream(new File("excels/berufe-heft-kldb2010-dwol-0-202111-xlsx.xlsx"));
    try (Workbook workbook = new XSSFWorkbook(fis)) {

      Sheet sheet = workbook.getSheetAt(4);
      Row dateRow = sheet.getRow(12);
      Cell dateCell = dateRow.getCell(3);
      Date date = dateCell.getDateCellValue();
      Calendar cal = Calendar.getInstance();
      cal.setTime(date);
      int year = cal.get(Calendar.YEAR);
      String month = nToMonth(cal.get(Calendar.MONTH));
      Row totalRow = sheet.getRow(15);
      Cell totalCell = totalRow.getCell(3);
      int total = (int) totalCell.getNumericCellValue();
      for (Row row : sheet) {
        if (row.getRowNum() < 21) {
          continue;
        }
        if (row.getRowNum() >= 750)
          break;
        Cell cell = row.getCell(0);
        if (cell == null) {
          continue;
        }
        String stringValue = cell.getStringCellValue();
        if (stringValue.isEmpty()) {
          continue;
        }
        if (stringValue.length() < 3) {
          String bezeichung = row.getCell(1).getStringCellValue();
          Beruf beruf = new Beruf(stringValue, bezeichung);
          berufe.add(beruf);
          map.put(stringValue, new ArrayList<>());
        }
        int value = Integer.parseInt(stringValue);
        if (value == 0) {
          continue;
        }
      }
    }
    for (int year = 2020, month = 11; year < 2021 || month <= 11; month++) {
      if (month > 12) {
        month = 1;
        year++;
      }
      figuresForYearMonth(year + (month < 10 ? "0" : "") + month);
    }
    List<StringBuilder> rows = new ArrayList<>();
    for (int i = 0; i < 13; i++) {
      rows.add(new StringBuilder());
    }
    // add totals in first column
    rows.get(0).append("Total").append(";");
    for (int i = 0; i < 12; i++) {
      rows.get(i + 1).append(totals.get(i)).append(";");
    }
    // add keineAngabe in second column
    rows.get(0).append("Keine Angabe").append(";");
    for (int i = 0; i < 12; i++) {
      rows.get(i + 1).append(keineAngabe.get(i)).append(";");
    }
    for (Beruf beruf : berufe) {
      rows.get(0).append(beruf.bezeichnung());
      rows.get(0).append(";");
      for (int i = 0; i < 12; i++) {
        rows.get(i + 1).append(map.get(beruf.id()).get(i));
        rows.get(i + 1).append(";");
      }
    }
    for (StringBuilder sb : rows) {
      // remove last char
      if (sb.length() > 0)
        sb.deleteCharAt(sb.length() - 1);
    }
    StringBuilder sb = new StringBuilder();
    for (StringBuilder row : rows) {
      sb.append(row);
      sb.append("\n");
    }
    try (BufferedWriter writer = new BufferedWriter(
        new OutputStreamWriter(new FileOutputStream("data.csv"), StandardCharsets.UTF_8))) {
      writer.write('\uFEFF');
      writer.write(sb.toString());
    }
  }

  private void figuresForYearMonth(String yearMonth) throws IOException {
    FileInputStream fis = new FileInputStream(
        new File("excels/berufe-heft-kldb2010-dwol-0-" + yearMonth + "-xlsx.xlsx"));
    try (Workbook workbook = new XSSFWorkbook(fis)) {
      Sheet sheet = workbook.getSheetAt(4);
      int curBeruf = 0;
      Row totalRow = sheet.getRow(15);
      totals.add((int) totalRow.getCell(3).getNumericCellValue());
      Row kaRow = sheet.getRow(20);
      keineAngabe.add((int) kaRow.getCell(3).getNumericCellValue());
      // TODO: get this months total
      for (Row row : sheet) {
        // skip header
        if (row.getRowNum() < 21) {
          continue;
        }
        // skip footer
        if (row.getRowNum() >= 750)
          break;
        // skip empty rows
        Cell cell = row.getCell(0);
        if (cell == null) {
          continue;
        }
        String stringValue = cell.getStringCellValue();
        // check if stringValue equals the current beruf id
        if (stringValue.isEmpty()) {
          continue;
        }
        if (stringValue.equals(berufe.get(curBeruf).id())) {
          // add the value to the current beruf
          map.get(berufe.get(curBeruf).id()).add((int) row.getCell(3).getNumericCellValue());
          curBeruf++;
        }
        // if there is no more berufs to add, break
        if (curBeruf >= berufe.size()) {
          break;
        }
      }
    }
  }

  public String padLeftZeros(String inputString, int length) {
    if (inputString.length() >= length) {
      return inputString;
    }
    StringBuilder sb = new StringBuilder();
    while (sb.length() < length - inputString.length()) {
      sb.append('0');
    }
    sb.append(inputString);

    return sb.toString();
  }

  private String nToMonth(int n) {
    return months[n];
  }
}
