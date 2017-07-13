package synchronoss.com.FrontEndTool.TestLink;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jdom2.CDATA;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;

public class TransferExcelToXML {
  private final static String tableHeaderSubArea = "SubArea";
  private final static String tableHeaderPrecondition = "Precondition";
  private final static String tableHeaderArea = "Area";
  private final static String tableHeaderSummary = "Summary";
  private final static String tableHeaderSteps = "TestSteps";
  private final static String tableHeaderExpectedReseult = "ExpectedResult";

  public static void main(String[] args) {
    String excelFile = args[0];
    String sheetNumberFromUser = args[1];
    // String sheetNumberFromUser = "1";
    // File caseFile = new File("C:\\Users\\roadmapqa\\Downloads\\WIT_Consolidated_testcases.xlsx");
    File caseFile = new File(excelFile);
    Workbook wb = null;
    int sheetNumber;
    int loopStartNumber = 2;
    try {

      // create file stream
      FileInputStream is = new FileInputStream(caseFile);

      Document document = new Document();
      document.setRootElement(new Element("testsuite"));
      wb = WorkbookFactory.create(is);

      if (sheetNumberFromUser.equals("*")) {
        sheetNumber = wb.getNumberOfSheets();
      } else {
        sheetNumber = Integer.parseInt(sheetNumberFromUser);
        loopStartNumber = sheetNumber;
      }
      
      // get the first sheet
      for (int i = loopStartNumber - 1; i < sheetNumber; i++) {
        File xmlFile2 = new File(wb.getSheetName(i) + ".xml");
        List<String> columnHeaders = new ArrayList<String>();
        System.out.println(wb.getSheetName(i));
        Sheet sheetNum = wb.getSheetAt(i);

        for (Iterator<Row> rowsIT = sheetNum.rowIterator(); rowsIT.hasNext();) {
          Row row = rowsIT.next();

          boolean emptyRow = true;
          int index = 0;
          Element firstParentElement = new Element("testsuite");
          Element secondParentElement = new Element("testcase");
          Element testStepsElement = new Element("steps");
          Element testStepElement = new Element("step");

          // Iterate through the cells.
          for (Iterator<Cell> cellsIT = row.cellIterator(); cellsIT.hasNext();) {
            emptyRow = false;
            Element childElement = null;
            Cell cell = cellsIT.next();
            cell.setCellType(Cell.CELL_TYPE_STRING);

            if (row.getRowNum() == 0) {
              columnHeaders.add(cell.getStringCellValue());

            } else {

              // remove the white space
              String elementTagName = columnHeaders.get(index).replaceAll("\\s+", "");

              // check with element tag
              switch (elementTagName) {
                case tableHeaderSubArea:
                  elementTagName = "summary";
                  break;
                case tableHeaderPrecondition:
                  elementTagName = "preconditions";
                  break;
                case tableHeaderArea: {
                  firstParentElement.setAttribute("name", cell.getStringCellValue());
                  document.getRootElement().addContent(firstParentElement);
                  index++;
                  continue;
                }
                case tableHeaderSummary: {
                  secondParentElement.setAttribute("name", cell.getStringCellValue());
                  firstParentElement.addContent(secondParentElement);
                  index++;
                  continue;
                }
                case tableHeaderSteps: {
                  Element stepNumberElement = new Element("step_number");
                  stepNumberElement.addContent(new CDATA("1"));
                  Element stepActionElement = new Element("actions");

                  stepActionElement.addContent(new CDATA(cell.getStringCellValue()));
                  testStepElement.addContent(stepNumberElement);
                  testStepElement.addContent(stepActionElement);
                  testStepsElement.addContent(testStepElement);
                  secondParentElement.addContent(testStepsElement);
                  index++;
                  continue;
                }
                case tableHeaderExpectedReseult: {
                  Element expectedResult = new Element("expectedresults");
                  expectedResult.addContent(new CDATA(cell.getStringCellValue()));
                  testStepElement.addContent(expectedResult);
                  index++;
                  continue;
                }
              }

              // if the cell content is empty, escape current loop
              if (cell.getStringCellValue().equals("")) {
                index++;
                continue;
              }

              // add other element to second parent
              childElement = new Element(elementTagName);
              childElement.addContent(new CDATA(cell.getStringCellValue()));
              secondParentElement.addContent(childElement);

            }
            index++;
          }
        }

        // display with pretty format, and set encoding as UTF-8
        Format format = Format.getPrettyFormat();
        format.setEncoding("UTF-8");

        // if use the writeXML class, it will not convert to UTF-8
        XMLOutputter xmlOutput = new XMLOutputter(format);
        xmlOutput.output(document, new FileOutputStream(xmlFile2));
      }
    } catch (FileNotFoundException e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    } catch (InvalidFormatException e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    } catch (IOException e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }
}

