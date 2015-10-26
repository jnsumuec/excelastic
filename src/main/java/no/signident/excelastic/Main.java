package no.signident.excelastic;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.elasticsearch.action.index.IndexResponse;
import org.elasticsearch.client.Client;
import org.elasticsearch.client.transport.TransportClient;
import org.elasticsearch.common.transport.InetSocketTransportAddress;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.InetAddress;
import java.text.SimpleDateFormat;
import java.util.*;

public class Main {
    public static void main(String[] args) throws IOException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
        sdf.setTimeZone(TimeZone.getTimeZone("GMT"));
        FileInputStream fileInputStream = new FileInputStream("superstore.xlsx");

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.rowIterator();

        List<String> headerNames = new ArrayList<>();

        Client client = TransportClient.builder().build()
                .addTransportAddress(new InetSocketTransportAddress(InetAddress.getByName("localhost"), 9300));


        boolean firstRow = true;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            Iterator<Cell> cellIterator = row.cellIterator();

            int columnIndex = 0;
            Map<String, Object> document = new HashMap<>();


            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (firstRow) {
                    headerNames.add(cell.getStringCellValue());
                } else {
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC: {
                            document.put(headerNames.get(columnIndex), cell.getNumericCellValue());
                            break;
                        }
                        case Cell.CELL_TYPE_STRING: {
                            document.put(headerNames.get(columnIndex), cell.getStringCellValue());
                            break;
                        }
                        case Cell.CELL_TYPE_FORMULA: {
                            document.put(headerNames.get(columnIndex), sdf.format(cell.getDateCellValue()));
                            break;
                        }
                        case Cell.CELL_TYPE_BOOLEAN: {
                            document.put(headerNames.get(columnIndex), cell.getBooleanCellValue());
                            break;
                        }
                        default: {

                        }
                    }
                }
                columnIndex++;
            }

            IndexResponse response = client.prepareIndex("twitter", "tweet")
                    .setSource(document)
                    .get();
            firstRow = false;
        }
    }
}