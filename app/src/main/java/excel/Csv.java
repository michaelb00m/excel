package excel;

import java.util.ArrayList;
import java.util.List;

public class Csv {
  private List<String> rows = new ArrayList<>();

  public void addRow(String row) {
    rows.add(row);
  }

  @Override
  public String toString() {
    StringBuilder sb = new StringBuilder();
    for (String row : rows) {
      sb.append(row);
      sb.append("\n");
    }
    return sb.toString();
  }
}
