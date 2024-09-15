package org.timaeva;

import org.apache.poi.ss.usermodel.*;
import java.io.IOException;
import java.util.List;
import java.util.ArrayList;
import java.io.FileInputStream;

public class Main {
    public static void main(String[] args) throws IOException {
        String filePath = "D:\\Projects\\ProjectsJava\\JavaProjectITMOFirstTask\\src\\main\\resources\\Employee.xlsx";

        List<Individual> individuals = new ArrayList<>();
        List<Company> companies = new ArrayList<>();
        List<BankAccount> bankAccounts = new ArrayList<>();
        List<Employee> employies = new ArrayList<>();

        DataFormatter dataFormatter = new DataFormatter();
        Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Sheet sheet = workbook.getSheetAt(0);

        for (int i = 2; i < 12; i++) {
            Row row = sheet.getRow(i);

            // Получаем first name
            Cell cellInSearchColumn = row.getCell(5);
            String cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
            if (cellValue != null && !cellValue.isEmpty()) {
                var individual = new Individual();
                individual.FirstName = cellValue;

                cellInSearchColumn = row.getCell(6);
                cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
                individual.LastName = cellValue;

                cellInSearchColumn = row.getCell(7);
                cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
                individual.HasChildren = Boolean.parseBoolean(cellValue);

                cellInSearchColumn = row.getCell(8);
                cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
                individual.Age = Integer.parseInt(cellValue);

                GetEmployeeFields(row, dataFormatter, formulaEvaluator, individual);

                var bankAccount = getBankAccount(row, dataFormatter, formulaEvaluator);
                bankAccounts.add(bankAccount);
                individual.BankAccount = bankAccount;

                individuals.add(individual);
                employies.add(individual);
            }

            cellInSearchColumn = row.getCell(10);
            cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
            if (cellValue != null && !cellValue.isEmpty()) {
                var company = new Company();
                company.CompanyName = cellValue;

                cellInSearchColumn = row.getCell(11);
                cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
                company.Type = CompanyType.valueOf(cellValue);

                GetEmployeeFields(row, dataFormatter, formulaEvaluator, company);

                var bankAccount = getBankAccount(row, dataFormatter, formulaEvaluator);
                bankAccounts.add(bankAccount);
                company.BankAccount = bankAccount;

                companies.add(company);
                employies.add(company);
            }
        }

        System.out.println("Количество физических лиц среди сотрудников: " + individuals.size());
        System.out.println("Количество компаний среди сотрудников: " + companies.size());
        System.out.println("Имя и фамилия сотрудников, которым меньше 20 лет: ");

        for (Individual individual:individuals) {
            if (individual.Age < 20) {
                System.out.print(individual.FirstName + " " + individual.LastName + "\n");
            }
        }

        workbook.close();
    }


    // Заполнение полей BankAccount
    private static BankAccount getBankAccount(Row row, DataFormatter dataFormatter, FormulaEvaluator formulaEvaluator) {
        String cellValue;
        Cell cellInSearchColumn;
        var bankAccount = new BankAccount();
        cellInSearchColumn = row.getCell(13);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        bankAccount.Iban = cellValue;

        cellInSearchColumn = row.getCell(14);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        bankAccount.Bic = cellValue;

        cellInSearchColumn = row.getCell(14);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        bankAccount.AccountHolder = cellValue;
        return bankAccount;
    }

    // Заполнение полей Employee
    private static void GetEmployeeFields(Row row, DataFormatter dataFormatter, FormulaEvaluator formulaEvaluator, Employee individual) {
        String cellValue;
        Cell cellInSearchColumn;
        cellInSearchColumn = row.getCell(0);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        individual.Id = Long.valueOf(cellValue);

        cellInSearchColumn = row.getCell(1);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        individual.Email = cellValue;

        cellInSearchColumn = row.getCell(2);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        individual.Phone = cellValue;

        cellInSearchColumn = row.getCell(3);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        individual.Address = cellValue;
    }
}
