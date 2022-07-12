package com.example.encrypt2md5;

import org.apache.commons.codec.digest.DigestUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@SpringBootApplication(exclude = {DataSourceAutoConfiguration.class})
public class Encrypt2Md5Application {

    public static void main(String[] args) throws IOException {

        SpringApplication.run(Encrypt2Md5Application.class, args);

        Path filePath = Paths.get("C:\\Users\\wpy7634\\Desktop\\passwords.xlsx");
        InputStream is = Files.newInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        int rowCount = sheet.getLastRowNum();
        for (int i = 0; i <= rowCount; i++) {
            Row row = sheet.getRow(i);
            row.getCell(0).setCellType(CellType.STRING);
            String original = row.getCell(0).getStringCellValue();
            String md5 = DigestUtils.md5Hex(original);
            row.createCell(1).setCellValue(md5);
            System.out.println(original + " -> " + md5);
        }

        workbook.write(Files.newOutputStream(filePath));
        workbook.close();
    }

}
