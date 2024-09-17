package org.timaeva;

import org.apache.poi.ss.usermodel.*;
import java.io.IOException;
import java.util.List;
import java.util.ArrayList;
import java.io.FileInputStream;

public class Main {
    public static void main(String[] args) throws IOException {
        if (args.length == 0) {
            System.out.println("Пожалуйста, укажите путь к файлу в аргументах");
            return;
        }

        String filePath = args[0];

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
            var employee = new Employee();
            // Получаем first name
            Cell cellInSearchColumn = row.getCell(5);
            String cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
            if (cellValue != null && !cellValue.isEmpty()) {
                var individual = new Individual();
                individual.setFirstName(cellValue);

                cellInSearchColumn = row.getCell(6);
                cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
                individual.setLastName(cellValue);

                cellInSearchColumn = row.getCell(7);
                cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
                individual.setHasChildren(Boolean.parseBoolean(cellValue));

                cellInSearchColumn = row.getCell(8);
                cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
                individual.setAge(Integer.parseInt(cellValue));

                GetEmployeeFields(row, dataFormatter, formulaEvaluator, individual);

                individuals.add(individual);
                employee = individual;
            }

            cellInSearchColumn = row.getCell(10);
            cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
            if (cellValue != null && !cellValue.isEmpty()) {
                var company = new Company();
                company.setCompanyName(cellValue);

                cellInSearchColumn = row.getCell(11);
                cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
                company.setType(CompanyType.valueOf(cellValue));

                GetEmployeeFields(row, dataFormatter, formulaEvaluator, company);

                companies.add(company);
                employee = company;
            }

            var bankAccount = getBankAccount(row, dataFormatter, formulaEvaluator);
            bankAccounts.add(bankAccount);
            employee.setBankAccount(bankAccount);
            employies.add(employee);
        }

        System.out.println("Количество физических лиц среди сотрудников: " + individuals.size());
        System.out.println("Количество компаний среди сотрудников: " + companies.size());
        System.out.println("Имя и фамилия сотрудников, которым меньше 20 лет: ");

        for (Individual individual:individuals) {
            if (individual.getAge() < 20) {
                System.out.print(individual.getFirstName() + " " + individual.getLastName() + "\n");
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
        bankAccount.setIban(cellValue);

        cellInSearchColumn = row.getCell(14);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        bankAccount.setBic(cellValue);

        cellInSearchColumn = row.getCell(14);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        bankAccount.setAccountHolder(cellValue);
        return bankAccount;
    }

    // Заполнение полей Employee
    private static void GetEmployeeFields(Row row, DataFormatter dataFormatter, FormulaEvaluator formulaEvaluator, Employee individual) {
        String cellValue;
        Cell cellInSearchColumn;
        cellInSearchColumn = row.getCell(0);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        individual.setId(Long.valueOf(cellValue));

        cellInSearchColumn = row.getCell(1);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        individual.setEmail(cellValue);

        cellInSearchColumn = row.getCell(2);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        individual.setPhone(cellValue);

        cellInSearchColumn = row.getCell(3);
        cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator);
        individual.setAddress(cellValue);
    }
}
