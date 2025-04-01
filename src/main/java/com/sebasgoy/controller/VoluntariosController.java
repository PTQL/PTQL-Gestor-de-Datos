package com.sebasgoy.controller;

import com.sebasgoy.dto.Voluntario;
import com.sebasgoy.service.VoluntarioService;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;

import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

@RestController
@RequestMapping("/api/voluntarios")
@AllArgsConstructor
public class VoluntariosController {

    private final VoluntarioService voluntarioService;

    @GetMapping("/export")
    public ResponseEntity<Resource> exportVoluntarios(@RequestParam String ruta) {
        try {
            List<Voluntario> voluntarios = voluntarioService.getAll();

            // Generar archivo Excel
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Voluntarios");

            // Crear estilo para las cabeceras
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // Crear encabezado
            Row headerRow = sheet.createRow(0);
            String[] headers = {"ID", "Nombre", "DNI", "Celular", "Edad"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }
            // Rellenar datos
            int rowNum = 1;
            for (Voluntario v : voluntarios) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(v.getId());
                row.createCell(1).setCellValue(v.getNombre());
                row.createCell(2).setCellValue(v.getDni());
                row.createCell(3).setCellValue(v.getTelefono());
                row.createCell(4).setCellValue(v.getEdad());
            }

            workbook.write(outputStream);
            workbook.close();

            // Guardar en la ruta especificada
            Path filePath = Paths.get(ruta, "voluntarios.xlsx");
            Files.write(filePath, outputStream.toByteArray());

            // Devolver el archivo como descarga
            ByteArrayResource resource = new ByteArrayResource(outputStream.toByteArray());

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=voluntarios.xlsx")
                    .body(resource);

        } catch (IOException e) {
            return ResponseEntity.internalServerError().body(null);
        }
    }
}

