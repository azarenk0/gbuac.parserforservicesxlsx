package ru.gbuac.parserforservicesxlsx.service.impl;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import ru.gbuac.parserforservicesxlsx.service.ParserService;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Azarenko Vladimir
 * @since 02.04.2020 11:06
 */
@Slf4j
@Service
public class ParserServiceImpl implements ParserService {

    @Value("${parser.file-path}")
    public String filePath;

    @Value("${parser.output-path}")
    public String outputPath;

    @Value("${parser.sheet-index}")
    public Integer sheetIndex;

    @Value("${parser.column-for-parse}")
    public Integer columnForParse;

    @Value("${parser.in-delimiter}")
    public String inDelimiter;

    @Value("${parser.out-delimiter}")
    public String outDelimiter;

    @Override
    public void startProcess() throws IOException {
        XSSFWorkbook workbook = openDocument();
        XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
        List<String> exportList = new ArrayList<>();

        for (Row row : sheet) {
            if (row.getCell(0) != null) {
                if (row.getCell(0).toString().equals("RegNumber")) continue;
                String regNumber = row.getCell(0).toString().trim();

                if (row.getCell(columnForParse) != null && !row.getCell(columnForParse).toString().trim().isEmpty()) {
                    if (row.getCell(columnForParse).toString().contains(inDelimiter)) {
                        String[] tempValues = row.getCell(columnForParse).toString().split(inDelimiter);
                        for (String value : tempValues) {
                            exportList.add(String.format("%s%s%s", regNumber, outDelimiter, parseValue(value)));
                        }
                    } else {
                        String value = row.getCell(columnForParse).toString();
                        exportList.add(String.format("%s%s%s", regNumber, outDelimiter, parseValue(value)));
                    }
                }
            }
        }
        Files.write(Paths.get(outputPath), exportList, Charset.defaultCharset());
    }

    private XSSFWorkbook openDocument() throws IOException {
        return new XSSFWorkbook(new FileInputStream(new File(filePath)));
    }

    private Integer parseValue(String value) {
        return (int) (Double.parseDouble(value));
    }

}