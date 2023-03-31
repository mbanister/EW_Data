package org.example;
import java.awt.*;
import java.io.*;
import java.lang.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Vector;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.ini4j.Wini;

import java.awt.event.*;
import java.io.File;
import javax.swing.*;
import javax.swing.border.EtchedBorder;
import javax.swing.border.TitledBorder;
import java.awt.BorderLayout;
import javax.swing.BorderFactory;

import javax.swing.border.Border;

//todo:validate copy,show filename,beautify,select sheet, change copier to string/var when no sum

public class Main {
    //static String destinationPath;
    public static void main(String[] args){
        new MyFrame();
    }
    public static class MyFrame extends JFrame implements ActionListener{
        JLabel ds,dd,d1,d2,d3,d4,d5,d6,textError,sourceError,destError;
        JButton sourceButton, destButton;
        JTextField tf1,tf2,tf3,tf4,function;
        JComboBox menu;
        JButton confirm,add;
        JTextArea console;
        JScrollPane wheel;
        public String source,destination, ssnSource, ssnDest;
        //public String[] ssnSource,ssnDest;
        //public List<String> ssnDest = new ArrayList<>();
        public List<List<String>> mvSource = new ArrayList<List<String>>();
        public List<List<String>> mvDest = new ArrayList<List<String>>();
        //public List<List<String>> mvSource,mvDest = new ArrayList<List<String>>();
        public List<Integer> type= new ArrayList<>();
        MyFrame(){
            this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            this.setLayout(new FlowLayout());
            ds = new JLabel("Select Source .xlsx File");
            ds.setBounds(50,40,150,30);
            sourceButton = new JButton("Choose File...");
            sourceButton.setBounds(60,70,120,20);
            dd = new JLabel("Select Destination .xlsx File");
            dd.setBounds(210,40,160,30);
            destButton = new JButton("Choose File...");
            destButton.setBounds(230,70,120,20);
            d1 = new JLabel("SSN source column");
            d1.setBounds(65,100,150,20);
            tf1=new JTextField();
            tf1.setBounds(95,120,50,20);
            d2 = new JLabel("SSN destination column");
            d2.setBounds(220,100,150,20);
            tf2=new JTextField();
            tf2.setBounds(265,120,50,20);
            String[] type = {"Copy", "Sum"};
            menu = new JComboBox<>(type);
            menu.setBounds(170, 150,70,20);
            d3 = new JLabel("<html>Copy target column(s)<br/>  &emsp;&nbsp;(Separate by , )</html>");
            d3.setBounds(65,180,150,40);
            tf3=new JTextField();
            tf3.setBounds(95,220,50,20);
            d4 = new JLabel("<html>Paste to target column(s)<br/>  &emsp;&ensp;(Separate by , )</html>");
            d4.setBounds(225,180,150,40);
            tf4=new JTextField();
            tf4.setBounds(265,220,50,20);
            d5 = new JLabel("<html>Add target column(s)<br/>  &emsp;&nbsp;(Separate by , )</html>");
            d5.setBounds(63,180,150,40);
            d6 = new JLabel("<html>Paste sum to target column<br/>  &emsp;(only paste to one)</html>");
            d6.setBounds(225,175,160,40);
            d5.setVisible(false);d6.setVisible(false);
            add = new JButton("add");
            add.setBounds(150,250,100,20);
            textError = new JLabel("<html><font color='red'>Text field not entered</font></html>");
            textError.setBounds(135,20,200,20);
            textError.setVisible(false);
            sourceError = new JLabel("<html><font color='red'>Source file not entered or .xlsx</font></html>");
            sourceError.setBounds(100,220,200,20);
            sourceError.setVisible(false);
            destError = new JLabel("<html><font color='red'>Destination file not entered or .xlsx</font></html>");
            destError.setBounds(100,20,200,20);
            destError.setVisible(false);
            function = new JTextField();
            function.setBounds(20,270,270,20);
            confirm = new JButton("Run");
            confirm.setBounds(310,270,70,20);
            console = new JTextArea();
            console.setLineWrap(true);
            console.setEditable(false);
            wheel = new JScrollPane(console);
            wheel.setBounds(20,350,370,350);
            wheel.setVisible(true);
            wheel.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
            wheel.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            sourceButton.addActionListener(this);
            destButton.addActionListener(this);
            menu.addActionListener(this);
            confirm.addActionListener(this);
            add.addActionListener(this);
            this.add(ds);this.add(dd);
            this.add(sourceButton);this.add(destButton);
            this.add(d1);this.add(d2);this.add(d3);this.add(d4);this.add(d5);this.add(d6);
            this.add(tf1);this.add(tf2);this.add(tf3);this.add(tf4);
            this.add(menu);
            this.add(confirm);
            this.add(textError);this.add(sourceError);this.add(destError);
            this.add(wheel);
            this.add(function);this.add(add);
            this.setSize(420,750);
            this.setLayout(null);
            this.setVisible(true);
            this.setResizable(false);
            //String SourcePath = getSource();
            //String destinationPath = getDestination();
        }

        @Override
        public void actionPerformed(ActionEvent e){
            if (e.getSource() == sourceButton || e.getSource() == destButton) {
                String path = "C:\\Users";
                //System.setProperty("log4j.configurationFile","./path_to_the_log4j2_config_file/log4j2.xml");
                JFileChooser fileChooser = new JFileChooser();
                try{
                    Wini ini = new Wini(new File("configure.ini"));
                    if (e.getSource() == sourceButton)
                        path = ini.get("Directory_Paths", "Source_Path");
                    else
                        path = ini.get("Directory_Paths", "Destination_Path");
                    // To catch basically any error related to writing to the file
                    // (The system cannot find the file specified)
                }catch(Exception d){
                    System.err.println(d.getMessage());
                }
                fileChooser.setCurrentDirectory(new File(path)); //sets current directory
                int response = fileChooser.showOpenDialog(null); //select file to open
                //int response = fileChooser.showSaveDialog(null); //select file to save
                if (response == JFileChooser.APPROVE_OPTION) {
                    File file = new File(fileChooser.getSelectedFile().getAbsolutePath());
                    if (file.isFile()) {
                        if (e.getSource() == sourceButton) {
                            source = file.getPath();
                            sourceError.setVisible(false);
                        }
                        else {
                            destination = file.getPath();
                            destError.setVisible(false);
                        }
                        try{
                            Wini ini = new Wini(new File("configure.ini"));
                            if (e.getSource() == sourceButton)
                                ini.put("Directory_Paths", "Source_Path", file.getParent());
                            else
                                ini.put("Directory_Paths", "Destination_Path", file.getParent());
                            ini.store();
                            // To catch basically any error related to writing to the file
                            // (The system cannot find the file specified)
                        }catch(Exception d){
                            System.err.println(d.getMessage());
                        }
                    }
                    else {
                        if (e.getSource() == sourceButton)
                            sourceError.setVisible(true);
                        else
                            destError.setVisible(true);
                    }
                }
            }
            if (Objects.requireNonNull(menu.getSelectedItem()).toString().equalsIgnoreCase("Copy")){
                d3.setVisible(true);d4.setVisible(true);
                d5.setVisible(false);d6.setVisible(false);
            }
            else{
                d3.setVisible(false);d4.setVisible(false);
                d5.setVisible(true);d6.setVisible(true);
            }
            if(e.getSource() == add){
                String ssnSourceColumn = tf1.getText().toUpperCase();
                String SC = tf3.getText().toUpperCase();
                String ssnDestColumn = tf2.getText().toUpperCase();
                String DC = tf4.getText().toUpperCase();
                tf1.setEditable(false);
                tf2.setEditable(false);
                if (function.getText().isEmpty())
                    function.setText(function.getText() + ssnSourceColumn + "," + ssnDestColumn + " | ");
                else
                    function.setText(function.getText() + "| ");
                //function.setText(function.getText() +ssnSourceColumn+"& "+SC);
                if (menu.getSelectedItem().toString().equalsIgnoreCase("Copy"))
                    function.setText(function.getText() + SC + " = ");
                else
                    function.setText(function.getText() + SC + " + ");
                function.setText(function.getText() +DC);
            }
                /*while (!func.isEmpty()){

                }*/
            if (e.getSource() == confirm) {
                /*String ssnSourceColumn = tf1.getText().toUpperCase();
                String SC = tf3.getText().toUpperCase();
                String ssnDestColumn = tf2.getText().toUpperCase();
                String DC = tf4.getText().toUpperCase();*/
                boolean source;
                String str = function.getText().toUpperCase();
                str = str.replaceAll("\\s", "");
                String[] func = str.split("\\|");
               // List<List<String>> mvSource = new ArrayList<List<String>>();
                //List<List<String>> mvDest = new ArrayList<List<String>>();
                for (int j = 0; j < func.length; j++) {
                    String s = func[j];
                    String builder;
                    //List<String> split= new ArrayList<>();
                    String[] split = s.split("((?=[=+,])|(?<=[=+,]))");
                    //mvSource.add(new ArrayList<String>());
                    //mvDest.add(new ArrayList<String>());
                    ArrayList<String> S = new ArrayList<String>();
                    ArrayList<String> D = new ArrayList<String>();
                    source = true;
                    for (int i = 0; i < split.length; i++) {
                        if (split[i].matches("^[A-Z0-9]+$")) {
                            if (j == 0) {
                                if (source) {
                                    ssnSource = split[i];
                                    source = false;
                                }
                                else
                                    ssnDest = split[i];
                            }
                            else {
                                if (split[i].matches("^[A-Z0-9]+$")) {
                                    if (source) {
                                        //mvSource.get(j).add(j, split[i]);

                                        S.add(split[i]);
                                    }
                                    else
                                        //mvDest.get(j).add(j, split[i]);
                                        D.add(split[i]);
                                }
                            }
                        }
                        if (split[i].matches("[+]")) {
                            type.add(1);
                            source = false;
                        }
                        if (split[i].matches("\\=")) {
                            type.add(0);
                            source = false;
                        }
                    }
                    if(!S.isEmpty())
                        mvSource.add(S);
                    if(!D.isEmpty())
                        mvDest.add(D);
                }

                FileInputStream inputStream = null;
                try {
                    inputStream = new FileInputStream(getSource());
                } catch (FileNotFoundException ex) {
                    throw new RuntimeException(ex);
                }
                FileInputStream outputStream = null;
                try {
                    outputStream = new FileInputStream(getDestination());
                } catch (FileNotFoundException ex) {
                    throw new RuntimeException(ex);
                }
                Workbook sourceWorkbook = null;
                try {
                    sourceWorkbook = new XSSFWorkbook(inputStream);
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
                Workbook destinationWorkbook = null;
                try {
                    destinationWorkbook = WorkbookFactory.create(outputStream);
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
                Sheet sourceSheet = sourceWorkbook.getSheetAt(0);//todo select sheet
                Sheet destinationSheet = destinationWorkbook.getSheetAt(0);//todo select sheet
                Pattern regex = Pattern.compile("^(?!666|000|9\\d{2})\\d{3}-(?!00)\\d{2}-(?!0{4})\\d{4}$");
                //String mvSourceColumn;
                //String mvDestColumn;
                //SC = SC.replaceAll("\\s","");
                //DC = DC.replaceAll("\\s","");
                //String[] mvSourceColumn = SC.split(",");
                //String[] mvDestColumn = DC.split(",");
                Vector<Double> sum = new Vector<>();
                //int copies = SC.replaceAll("[^,]","").length();
                for (int i = 0; i < type.size(); i++) {
                    int rowNumber = 0;
                    for (int j = 0; j < mvSource.get(0).size(); j++) {
                        for (Row sourceRow : sourceSheet) {
                            int celltype = cellType(sourceRow,mvSource.get(i).get(j));
                            rowNumber++;
                            //window Mapping:
                            String ssn = GetString(sourceRow, ssnSource); //todo
                            if (!regex.matcher(ssn).matches()) continue; // skip it if it's not an SSN
                            double intMarketValue = 404;
                            String strMarketValue = "error";
                            if (celltype == 1) {
                                intMarketValue = GetDecimal(sourceRow, mvSource.get(i).get(j));
                            }
                            else if (celltype == 2) {
                                strMarketValue = GetString(sourceRow, mvSource.get(i).get(j));
                            }
                            //Report source data
                            if (celltype == 1)
                                console.append(String.format("{%d}: SSN {%s}, Column %s: {%f}%n", rowNumber, ssn, mvSource.get(i).get(j), intMarketValue));
                            if (celltype == 0)
                                console.append(String.format("{%d}: SSN {%s}, Column %s: {%s}%n", rowNumber, ssn, mvSource.get(i).get(j), strMarketValue));
                            Row destinationRow = FindColumnContainingTerm(destinationSheet, ssn, ssnDest);
                            if (destinationRow == null) {
                                console.append(String.format("SSN {%s} was not found in destination worksheet%n", ssn));
                                destinationRow = createMissingRow(destinationSheet, regex, ssnDest);
                                if (destinationRow == null) {
                                    continue;
                                }
                                console.append(String.format("Setting SSN {%s} on row{%d} of the destination sheet%n", ssn, destinationRow.getRowNum()));

                            }
                            if (i == 0)
                                sum.add(0.0);
                            //Destination mapping:
                            if (type.get(i) == 0) {
                                if (celltype == 1)
                                    CellUtil.getCell(destinationRow, ExcelColumnNameToNumber(mvDest.get(i).get(j))).setCellValue(intMarketValue);
                                if (celltype == 0)
                                    CellUtil.getCell(destinationRow, ExcelColumnNameToNumber(mvDest.get(i).get(j))).setCellValue(intMarketValue);

                            }
                            if (type.get(i) == 1) {
                                //todo error if string in sum
                                sum.set(rowNumber - 1, sum.get(rowNumber - 1) + intMarketValue);
                                //if(i+1 == mvSourceColumn.length){
                                CellUtil.getCell(destinationRow, ExcelColumnNameToNumber(mvDest.get(i).get(j))).setCellValue(sum.elementAt(rowNumber - 1));
                                //}
                            }
                        }
                    }
                    FileOutputStream os = null;
                    try {
                        os = new FileOutputStream(getDestination());
                    } catch (FileNotFoundException ex) {
                        throw new RuntimeException(ex);
                    }
                    try {
                        destinationWorkbook.write(os);
                    } catch (IOException ex) {
                        throw new RuntimeException(ex);
                    }
                    try {
                        destinationWorkbook.close();
                    } catch (IOException ex) {
                        throw new RuntimeException(ex);
                    }
                    try {
                        outputStream.close();
                    } catch (IOException ex) {
                        throw new RuntimeException(ex);
                    }
                }
                tf1.setEditable(true);
                tf3.setEditable(true);
            }
        }
        public String getSource() {
            return this.source;
        }
        public String getDestination() {
            return this.destination;
        }
    }
    static int cellType(Row sourceRow, String columnName){
        Cell cell = CellUtil.getCell(sourceRow, ExcelColumnNameToNumber(columnName));
        if (cell.getCellType() == CellType.FORMULA) {
            if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                return 1;
            }
            return 2;
        }
        else if (cell.getCellType() == CellType.NUMERIC) {
            return 1;
        } else {
            return 2;
        }
    }
    static String GetString(Row sourceRow, String columnName)
    {
        int colNumber = ExcelColumnNameToNumber(columnName);
        DataFormatter formatter = new DataFormatter();
        //return sourceRow.getCell(colNumber).getStringCellValue();
        return formatter.formatCellValue(sourceRow.getCell(colNumber));
    }
    static double GetDecimal(Row sourceRow, String columnName)
    {
        var colNumber = ExcelColumnNameToNumber(columnName);
        DataFormatter formatter = new DataFormatter();
        if (formatter.formatCellValue(sourceRow.getCell(colNumber)).length() == 0){
            return 0;
        }
        else{
            String str = formatter.formatCellValue(sourceRow.getCell(colNumber));
            return Double.parseDouble(str);
            //return sourceRow.getCell(colNumber).GetValue<decimal>();
        }
        //return formatter.formatCellValue(sourceRow.getCell(colNumber)) ==  ? 0.0 : sourceRow.Cell(colNumber).GetValue<decimal>();
    }
    static int ExcelColumnNameToNumber(String columnName)
    {
        columnName = columnName.toUpperCase();

        var sum = 0;

        for (var i = 0; i < columnName.length(); i++)
        {
            sum *= 26;
            sum += columnName.charAt(i) - 'A';
        }

        return sum;
    }
    static Row createMissingRow(Sheet worksheet, Pattern regex, String column) {
            for (int lastRowIndex = worksheet.getLastRowNum(); lastRowIndex >= 0; lastRowIndex--) {
                Row row = worksheet.getRow(lastRowIndex);
                if (regex.matcher(GetString(row, column)).matches()) {
                    worksheet.shiftRows(lastRowIndex+1, worksheet.getLastRowNum(),1,true,false);
                    return worksheet.getRow(lastRowIndex);
                }
            }
            return null;
        }
    static Row FindColumnContainingTerm(Sheet worksheet, String term, String columnName) {
        //var colNumber = ExcelColumnNameToNumber(columnName);
        for (int j = worksheet.getFirstRowNum(); j <= worksheet.getLastRowNum(); j++) {
            Row row = worksheet.getRow(j);
            //Cell cell = row.getCell(colNumber);
            if(term.equals(GetString(row,columnName))) {
                return row;
            }
        }
        return null;
          /*Cell cells = worksheet.Search(term);

        if (cells.Count() != 1) {
            return null;
        }

        var row = cells.First().WorksheetRow();
        return row;*/
    }

}
        /*
        using var sourceWorkbook = new XLWorkbook(args[0]);
        var sourceWorkSheet = sourceWorkbook.Worksheets.Worksheet(args[1]);

        using var destinationWorkbook = new XLWorkbook(args[2]);
        var destinationWorkSheet = destinationWorkbook.Worksheets.Worksheet(args[3]);

        var validateSSNRegex = new Regex("^(?!666|000|9\\d{2})\\d{3}-(?!00)\\d{2}-(?!0{4})\\d{4}$");

        var rowNumber = 0;

        foreach (var sourceRow in sourceWorkSheet.Rows())
        {
            rowNumber++;

            //Source Mapping:
            var ssn = GetString(sourceRow, "C");
            if (validateSSNRegex.IsMatch(ssn) == false) continue; // skip it if it's not an SSN

            var marketValue = GetDecimal(sourceRow, "F");

            //Report source data
            Console.WriteLine($"{rowNumber}: SSN {ssn}, market value {marketValue}");
            var destinationRow = FindRowContainingTerm(destinationWorkSheet, ssn);
            if (destinationRow == null)
            {
                Console.WriteLine($"SSN {ssn} was not found in destination worksheet");
                continue; //skip it if SSN is not found -- TODO: Report it!
            }

            //Destination mapping:
            SetValue(destinationRow, "AR", marketValue);
        }

        destinationWorkbook.Save();

        return;


// Utility Functions
        IXLRow FindRowContainingTerm(IXLWorksheet worksheet, string term)
        {
            var cells = worksheet.Search(term);

            if (cells.Count() != 1)
            {
                return null;
            }

            var row = cells.First().WorksheetRow();
            return row;
        }

        decimal GetDecimal(IXLRow sourceRow, string columnName)
        {
            var colNumber = ExcelColumnNameToNumber(columnName);
            if (sourceRow.Cell(colNumber).GetString() == string.Empty)
            then{
              return 0.0;
            }
            else{
              return sourceRow.Cell(colNumber).GetValue<decimal>();
            }
            return sourceRow.Cell(colNumber).GetString() == string.Empty ? 0m : sourceRow.Cell(colNumber).GetValue<decimal>();
        }

        string GetString(IXLRow sourceRow, string columnName)
        {
            var colNumber = ExcelColumnNameToNumber(columnName);
            return sourceRow.Cell(colNumber).GetString();
        }

        void SetValue<T>(IXLRow destinationRow, string columnName, T value)
        {
            var colNumber = ExcelColumnNameToNumber(columnName);
            destinationRow.Cell(colNumber).SetValue(value);
        }

        int ExcelColumnNameToNumber(string columnName)
        {
            columnName = columnName.ToUpperInvariant();

            var sum = 0;

            for (var i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
    }
}*/