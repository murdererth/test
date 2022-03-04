package com.lib.excel.controller;

import com.lib.excel.service.ReportService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/benz/report")
public class ReportController {


    private ReportService reportService;

    public ReportController(ReportService reportService) {
        this.reportService = reportService;
    }

    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generateXlsxReport() {

        return createResponseEntity(reportService.generateXlsxReport(), "report.xlsx");
    }

    @PostMapping("/xls")
    public ResponseEntity<byte[]> generateXlsReport() {
        return createResponseEntity(reportService.generateXlsReport(), "report.xls");
    }

    private ResponseEntity createResponseEntity(byte[] report,String fileName) {

        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename="+fileName)
                .body(report);
    }
}
