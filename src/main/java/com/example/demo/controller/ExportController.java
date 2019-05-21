package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.ClientService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;
    @Autowired
    private FactureService factureService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        writer.println("Id" + ";" + "Nom" + ";" + "Prenom" + ";" + "Date de Naissance");

        for (Client client : allClients) {
            writer.println(client.getId() + ";"
                    + client.getNom() + ";"
                    + client.getPrenom() + ";"
                    + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("Id");

        Cell cellPrenom = headerRow.createCell(1);
        cellPrenom.setCellValue("Prénom");

        Cell cellNom = headerRow.createCell(2);
        cellNom.setCellValue("Nom");

        int iRow = 1;
        for (Client client : allClients) {
            Row row = sheet.createRow(iRow);

            Cell id = row.createCell(0);
            id.setCellValue(client.getId());

            Cell prenom = row.createCell(1);
            prenom.setCellValue(client.getPrenom());

            Cell nom = row.createCell(2);
            nom.setCellValue(client.getNom());

            iRow = iRow + 1;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/clients/{id}/factures/xlsx")
    public void factureXLSXByClient(@PathVariable("id") Long clientId, HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures-client-" + clientId + ".xlsx\"");
        List<Facture> factures = factureService.findFacturesClient(clientId);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Facture");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("Id");

        Cell cellTotal = headerRow.createCell(1);
        cellTotal.setCellValue("Prix Total");

        int iRow = 1;
        for (Facture facture : factures) {
            Row row = sheet.createRow(iRow);

            Cell id = row.createCell(0);
            id.setCellValue(facture.getId());

            Cell prenom = row.createCell(1);
            prenom.setCellValue(facture.getTotal());

            iRow = iRow + 1;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }
    
    @GetMapping("/factures/xlsx")
    public void facturesXLSXB(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");
        List<Client> clients = clientService.findAllClients();
        Workbook workbook = new XSSFWorkbook();
        
        for(Client client: clients) {
            
            Sheet sheetCli = workbook.createSheet(client.getNom() + " " + client.getPrenom());
            
            Row headerRowCli = sheetCli.createRow(0);
    
            Cell cellNom = headerRowCli.createCell(0);
            cellNom.setCellValue("Nom");
            Cell cellPrenom = headerRowCli.createCell(1);
            cellPrenom.setCellValue("Prenom");
            
            Row rowClient = sheetCli.createRow(1);
    
            Cell cellNomCli = rowClient.createCell(0);
            cellNomCli.setCellValue(client.getNom());
            Cell cellPrenomCli = rowClient.createCell(1);
            cellPrenomCli.setCellValue(client.getPrenom());
            
            List<Facture> factures = factureService.findFacturesClient(client.getId());
            
            for(Facture facture : factures) {
                
                Sheet sheetFact = workbook.createSheet("Facture " + facture.getId());
                
                Row headerRow = sheetFact.createRow(0);
    
                Cell cellDesignation = headerRow.createCell(0);
                cellDesignation.setCellValue("désignation");
                Cell cellQuantite = headerRow.createCell(1);
                cellQuantite.setCellValue("quantité");
                Cell cellPrixUnitaire = headerRow.createCell(2);
                cellPrixUnitaire.setCellValue("prixUnitaire");
                Cell cellPrixLigne = headerRow.createCell(3);
                cellPrixLigne.setCellValue("prixLigne");
    
                int iRow = 1;
                for (LigneFacture ligne : facture.getLigneFactures()) {
                    Row rowFact = sheetFact.createRow(iRow);
    
                    Cell cellDesignationFact = rowFact.createCell(0);
                    cellDesignationFact.setCellValue(ligne.getArticle().getLibelle());
                    Cell cellQuantiteFact = rowFact.createCell(1);
                    cellQuantiteFact.setCellValue(ligne.getQuantite());
                    Cell cellPrixUnitaireFact = rowFact.createCell(2);
                    cellPrixUnitaireFact.setCellValue(ligne.getArticle().getPrix());
                    Cell cellPrixLigneFact = rowFact.createCell(3);
                    cellPrixLigneFact.setCellValue(ligne.getSousTotal());
        
                    iRow = iRow + 1;
                }
            }
        }
        
        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
