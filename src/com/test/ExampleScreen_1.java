package com.test;

import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Savepoint;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;
import javax.swing.DefaultCellEditor;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.WindowConstants;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;


public class ExampleScreen_1 implements ActionListener,MouseListener,WindowListener
{
    String s,insTabRow;
    //connection paramaters acquired from front end
    String acqUsername,acqPwd,acqSid,acqHost,acqPort,selectedDB,driverName,url,tableNamesQuery,columnNamesQuery,
            insertStmtQuery,selectedTable,selectedSheet,expForSheetData,
            logFileDirPath,badFileDirPath,dataFileDirPath;
    boolean cmbCheck=true;
    boolean cmbStateCheck=true;
    boolean connectionCheck=false;
    int sheetCount,sheetRows,sheetColumns,pSheetColumns;
    Connection connection=null;
    Statement statement;
    ResultSet tableNamesRset,columnNamesRset;
    JLabel openLabel,connectionLabel,usernameLabel,pwdLabel,sidLabel,hostLabel,portLabel,mapLabel,titleLabel,dbTypeLabel;
    JTextField openTxtFld,usernameTxtFld,sidTxtFld,hostTxtFld,portTxtFld;
    JButton browseButton,connectButton,loadButton, refreshButton;
    JPasswordField pwdFld;
    JComboBox cmbSheetNames,cmbTableNames,cmbColumnNames, cmbDBTypes;
    JTable tab;
    JFileChooser fchoose;
    File fileSelected;
    Workbook wbk1;
    Sheet sheet;
    Savepoint savepoint1;
    Container pane;
    JScrollPane pan;
    int totRowCount=0;
    int badRowCount=0;
    int previousSheetIndex=0;
    int selectedSheetIndex=0;
    DefaultTableModel model;
    Object[][] sheetData,acqStateData,acqStateData_in,acqData,disTabsSheetData;
    List list=new ArrayList();
    Vector v=new Vector();
    String[][] tabStateData,tabRowData,xs,disTabsRsCs;
    JFrame frame;
    List allTabs=new ArrayList();
    List disTabs=new ArrayList();
    List disTabsCnt=new ArrayList();

    public  ExampleScreen_1()
    {
        frame = new JFrame("Nisum Loader Application");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        pane=frame.getContentPane();
        Font newFont =new Font("Serif",Font.BOLD,15);
        Font newTxtFont =new Font("Serif",Font.PLAIN,15);
        frame.addWindowListener(this);
        //Display the window.
        pane.setLayout(new GridBagLayout());
        //pane.setBackground(Color.WHITE);
        GridBagConstraints c=new GridBagConstraints();
        c.fill = GridBagConstraints.HORIZONTAL;
        c.insets=new Insets(5,5,5,5);
        //title label
        titleLabel=new JLabel("NISUM LOADER");
        c.gridx=3;
        c.gridy=0;
        c.gridwidth=3;
        titleLabel.setFont(new Font("Serif",Font.BOLD,50));
        titleLabel.setForeground(new Color(29,202,255));
        pane.add(titleLabel,c);

        //Open label
        openLabel =new JLabel("Input File :");
        c.gridx=1;
        c.gridy=1;
        pane.add(openLabel,c);
        openLabel.setFont(newFont);

        //open Txt Fld
        openTxtFld =new JTextField("",20);
        c.gridx=3;
        c.gridy=1;
        c.gridwidth=2;
        pane.add(openTxtFld,c);
        openTxtFld.addActionListener(this);
        openTxtFld.setFont(newTxtFont);
        openTxtFld.setForeground(Color.BLACK);
        openTxtFld.setBackground(new Color(255,255,160));
        openTxtFld.setEditable(false);
        openTxtFld.setToolTipText("Please Click Browse button to select the file");
        openTxtFld.addMouseListener(this);

        //Browse button
        browseButton=new JButton("Browse");
        browseButton.addActionListener(this);
        browseButton.setToolTipText("Click Browse to Select Files");
        c.gridx=5;
        c.gridy=1;
        pane.add(browseButton,c);

        //File Chooser Component
        fchoose=new JFileChooser();

        //Connections Label
        connectionLabel =new JLabel("Connections :");
        connectionLabel.setFont(new Font("Serif",Font.BOLD,20));
        connectionLabel.setForeground(new Color(29,202,255));
        c.gridx=1;
        c.gridy=2;
        c.gridwidth=2;
        pane.add(connectionLabel,c);

        //Username Label
        dbTypeLabel =new JLabel("DB Type :");
        dbTypeLabel.setFont(newFont);
        c.gridx=1;
        c.gridy=3;
        pane.add(dbTypeLabel,c);

        //ComboBox for SheetNames of Opened Workbook
        cmbDBTypes =new JComboBox();
        cmbDBTypes.addItem("MySQL");
        cmbDBTypes.addItem("ORACLE");
        cmbDBTypes.setToolTipText("Select a DB type");
        cmbDBTypes.addActionListener(this);
        cmbDBTypes.setFont(newTxtFont);
        cmbDBTypes.setForeground(Color.BLACK);
        cmbDBTypes.setBackground(new Color(255,255,160));
        c.weightx=0.5;
        c.gridx=3;
        c.gridy=3;
        pane.add(cmbDBTypes, c);

        //Username Label
        usernameLabel =new JLabel("Username :");
        usernameLabel.setFont(newFont);
        c.gridx=1;
        c.gridy=4;
        pane.add(usernameLabel,c);

        //Username text field
        usernameTxtFld =new JTextField("",30);
        usernameTxtFld.setForeground(Color.BLACK);
        usernameTxtFld.setFont(newTxtFont);
        usernameTxtFld.setBackground(new Color(255,255,160));
        usernameTxtFld.setToolTipText("Enter database username");
        c.gridx=3;
        c.gridy=4;
        c.gridwidth=1;
        pane.add(usernameTxtFld,c);
        usernameTxtFld.addActionListener(this);

        //Password label
        pwdLabel =new JLabel("Password :");
        pwdLabel.setFont(newFont);
        c.gridx=1;
        c.gridy=5;
        pane.add(pwdLabel,c);

        //Pass Field
        pwdFld =new JPasswordField("",30);
        pwdFld.setForeground(Color.BLACK);
        pwdFld.setFont(newTxtFont);
        pwdFld.setBackground(new Color(255,255,160));
        pwdFld.setToolTipText("Enter database password");
        c.gridx=3;
        c.gridy=5;
        c.gridwidth=1;
        pane.add(pwdFld,c);
        pwdFld.addActionListener(this);

        //Connect String  Label
        sidLabel =new JLabel("DB/Service Name :");
        sidLabel.setFont(newFont);
        c.gridx=1;
        c.gridy=6;
        pane.add(sidLabel,c);

        //Content String Txt Fld
        sidTxtFld =new JTextField("",30);
        sidTxtFld.setForeground(Color.BLACK);
        sidTxtFld.setFont(newTxtFont);
        sidTxtFld.setBackground(new Color(255,255,160));
        sidTxtFld.setToolTipText("Enter database connect string/service name");

        c.gridx=3;
        c.gridy=6;
        c.gridwidth=1;
        pane.add(sidTxtFld,c);
        sidTxtFld.addActionListener(this);

        //Host Label
        hostLabel =new JLabel("Host :");
        hostLabel.setFont(newFont);
        c.gridx=1;
        c.gridy=7;
        pane.add(hostLabel,c);

        //Host Txt Fld
        hostTxtFld=new JTextField("",30);
        hostTxtFld.setForeground(Color.BLACK);
        hostTxtFld.setFont(newTxtFont);
        hostTxtFld.setBackground(new Color(255,255,160));
        hostTxtFld.setToolTipText("Enter host name");

        c.gridx=3;
        c.gridy=7;
        c.gridwidth=1;
        pane.add(hostTxtFld,c);
        hostTxtFld.addActionListener(this);

        //Port Label
        portLabel=new JLabel("Port :");
        portLabel.setFont(newFont);
        c.gridx=1;
        c.gridy=8;
        pane.add(portLabel,c);

        //port Txt Fld
        portTxtFld =new JTextField("",30);
        portTxtFld.setForeground(Color.BLACK);
        portTxtFld.setFont(newTxtFont);
        portTxtFld.setBackground(new Color(255,255,160));
        portTxtFld.setToolTipText("Enter port number");

        c.gridx=3;
        c.gridy=8;
        c.gridwidth=1;
        pane.add(portTxtFld,c);
        portTxtFld.addActionListener(this);

        //Mappings Label
        mapLabel =new JLabel("Mappings :");
        mapLabel.setFont(new Font("Serif",Font.BOLD,20));
        mapLabel.setForeground(new Color(29,202,255));


        c.gridx=1;
        c.gridy=9;
        c.gridwidth=2;
        pane.add(mapLabel,c);

        //Connect button
        connectButton=new JButton("Connect");
        connectButton.setToolTipText("Click to connect to database");

        c.gridx=5;
        c.gridy=8;
        pane.add(connectButton,c);
        connectButton.addActionListener(this);

        //Refresh button
        refreshButton=new JButton("Refresh DB");
        refreshButton.setToolTipText("Click to refresh the database");

        c.gridx=5;
        c.gridy=7;
        pane.add(refreshButton,c);
        refreshButton.addActionListener(this);
        refreshButton.hide();

        //ComboBox for SheetNames of Opened Workbook
        cmbSheetNames =new JComboBox();
        cmbSheetNames.addItem("Select Sheet");
        cmbSheetNames.setToolTipText("Select a sheet in Opened Workbook");
        c.weightx=0.5;
        c.gridx=3;
        c.gridy=9;
        pane.add(cmbSheetNames,c);
        cmbSheetNames.addActionListener(this);
        cmbSheetNames.setFont(newTxtFont);
        cmbSheetNames.setForeground(Color.BLACK);
        cmbSheetNames.setBackground(new Color(255,255,160));

        //ComboBox for Tables to be loaded
        cmbTableNames =new JComboBox();
        cmbTableNames.setFont(newTxtFont);
        cmbTableNames.addActionListener(this);

        //ComboBox for Columns to be mapped to Element Selected
        cmbColumnNames =new JComboBox();
        cmbColumnNames.setFont(newTxtFont);
        cmbColumnNames.addActionListener(this);

        model=new DefaultTableModel();
        tab =new JTable(model);
        model.addColumn("Sheet Elements");
        model.addColumn("Table Name");
        model.addColumn("Column Name");
        model.addColumn("Expression");

        tab.setPreferredScrollableViewportSize(new Dimension(100,100));
        tab.getTableHeader().setReorderingAllowed(false);
        tab.getTableHeader().setFont(newFont);
        tab.setRowHeight(20);

        TableColumn sheetElementColumn=tab.getColumnModel().getColumn(0);
        TableColumn tableNameColumn=tab.getColumnModel().getColumn(1);
        TableColumn columnNameColumn=tab.getColumnModel().getColumn(2);
        TableColumn expressionColumn=tab.getColumnModel().getColumn(3);

        //code for jtable cell rendering
        DefaultTableCellRenderer renderer = new DefaultTableCellRenderer() {
            // override renderer preparation
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
                                                           boolean hasFocus,
                                                           int row, int column)
            {
                Component cell = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                cell.setFont(new Font("Serif", Font.PLAIN, 15));
                return cell;
            }
        };
        sheetElementColumn.setCellRenderer(renderer);
        tableNameColumn.setCellRenderer(renderer);
        columnNameColumn.setCellRenderer(renderer);
        expressionColumn.setCellRenderer(renderer);

        // sheetNameColumn.setCellEditor(new DefaultCellEditor(cmbSheetElements));
        tableNameColumn.setCellEditor(new DefaultCellEditor(cmbTableNames));
        columnNameColumn.setCellEditor(new DefaultCellEditor(cmbColumnNames));
        pan=new JScrollPane(tab);
        tab.getCellSelectionEnabled();
        c.gridx=1;
        c.gridy=10;
        c.gridwidth=10;
        pane.add(pan,c);

        //Load button
        loadButton=new JButton("    Load    ");
        loadButton.setToolTipText("Click to Load");

        JPanel pan2=new JPanel();
        pan2.add("START",loadButton);
        c.gridx=0;
        c.gridy=11;
        pane.add(pan2,c);
        pane.setBackground(new Color(255,255,255));
        loadButton.addActionListener(this);

        frame.pack();
        frame.setVisible(true);
        frame.setSize(800,650);
        frame.setLocation(300,50);
    }

    public String trimQuotes(String str)
    {
        if (str.startsWith("\'"))
        {
            str = str.substring(1, str.length());
        }
        if (str.endsWith("\'"))
        {
            str = str.substring(0, str.length() - 1);
        }
        return str;
    }

    public String replaceWord(String original, String find, String replacement)
    {
        int i = original.indexOf(find);
        if (i < 0)
        {
            return original;  // return original if 'find' is not in it.
        }
        String partBefore = original.substring(0, i);
        String partAfter  = original.substring(i + find.length());
        return partBefore + replacement + partAfter;
    }

    public void actionPerformed(ActionEvent e) {
        if(e.getSource()==connectButton || e.getSource()==refreshButton)
        {
            if(!connectButton.getText().equals("Connect")){
                if(connection!=null){
                    cmbColumnNames.removeAllItems();
                    cmbTableNames.removeAllItems();
                    connection = null;
                    connectButton.setText("Connect");
                    connectButton.setToolTipText("Click to connect to database");
                    JOptionPane.showMessageDialog(frame,"DB Connection closed","Alert",JOptionPane.INFORMATION_MESSAGE);
                }
            }else{
                acqUsername=usernameTxtFld.getText();
                acqPwd =new String(pwdFld.getPassword());
                acqSid =sidTxtFld.getText();
                acqHost=hostTxtFld.getText();
                acqPort=portTxtFld.getText();
                selectedDB = cmbDBTypes.getSelectedItem().toString();
                //Front End validations Need to ask the user to enter the required details if not done..using a JDialog
                System.out.println("Connection Button Pressed");
                if(acqUsername!=null && acqUsername.equals(""))
                {
                    JOptionPane.showMessageDialog(frame,"Please enter the Database Username ","Alert",JOptionPane.INFORMATION_MESSAGE);
                    usernameTxtFld.requestFocus(true);
                    connectionCheck=false;
                }
                else if (acqPwd!=null && acqPwd.equals(""))
                {
                    JOptionPane.showMessageDialog(frame,"Please enter the Database Password ","Alert",JOptionPane.INFORMATION_MESSAGE);
                    pwdFld.requestFocus(true);
                    connectionCheck=false;
                }
                else  if(acqSid!=null && acqSid.equals(""))
                {
                    JOptionPane.showMessageDialog(frame,"Please enter the Connect String / Service Name","Alert",JOptionPane.INFORMATION_MESSAGE);
                    sidTxtFld.requestFocus(true);
                    connectionCheck=false;
                }
                else  if(acqHost!=null && acqHost.equals(""))
                {
                    JOptionPane.showMessageDialog(frame,"Please enter the Host Name","Alert",JOptionPane.INFORMATION_MESSAGE);
                    hostTxtFld.requestFocus(true);
                    connectionCheck=false;
                }
                else if(acqPort!=null && acqPort.equals(""))
                {
                    JOptionPane.showMessageDialog(frame,"Please enter the Port Number ","Alert",JOptionPane.INFORMATION_MESSAGE);
                    portTxtFld.requestFocus(true);
                    connectionCheck=false;
                }
                else
                {
                    connectionCheck=true;
                }

                if(connectionCheck)
                {
                    if("ORACLE".equals(selectedDB)){
                        driverName="oracle.jdbc.driver.OracleDriver";
                        url= "jdbc:oracle:thin:@" + acqHost + ":" + acqPort + ":" + acqSid;
                    }else if("MySQL".equals(selectedDB)){
                        driverName = "com.mysql.jdbc.Driver";
                        url = "jdbc:mysql://" + acqHost + ":" + acqPort + "/" + acqSid;
                    }
                    try
                    {
                        Class.forName(driverName);
                        connection=DriverManager.getConnection(url, acqUsername, acqPwd);
                        if(connection!=null)
                        {
                            System.out.print("Connection Established");
                            connectButton.setText("Disconnect");
                            connectButton.setToolTipText("Click to disconnect");
                            JOptionPane.showMessageDialog(frame,"Connection Established. Press OK and Please Wait.....","Alert",JOptionPane.INFORMATION_MESSAGE);
                            statement =connection.createStatement();
                            if("ORACLE".equals(selectedDB)){
                                tableNamesQuery="select table_name from user_tables order by table_name";
                            }else if("MySQL".equals(selectedDB)){
                                tableNamesQuery="select table_name from information_schema.TABLES order by table_name";
                            }
                            tableNamesRset=statement.executeQuery(tableNamesQuery);
                            while(tableNamesRset.next())
                            {
                                cmbTableNames.addItem(tableNamesRset.getString("table_name"));
                            }
                            JOptionPane.showMessageDialog(frame,"Tables Loaded..Proceed to Mappings","Alert",JOptionPane.INFORMATION_MESSAGE);
                        }
                        else
                        {
                            System.out.print("Connection Failed");
                            JOptionPane.showMessageDialog(frame,"Sorry Connection Failed","Alert",JOptionPane.INFORMATION_MESSAGE);
                        }
                    }
                    catch (Exception c)
                    {
                        JOptionPane.showMessageDialog(frame,c.getMessage(),"Error",JOptionPane.ERROR_MESSAGE);
                    }
                }
            } //else - for connecting
        }
        if(e.getSource()==browseButton)
        {
            System.out.print("Your Pressed Browse Button");
            if (fchoose.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                fileSelected=fchoose.getSelectedFile();
                System.out.println("List index"+list.size());
                cmbSheetNames.setSelectedIndex(0);
                if(list.size()>0)
                {
                    openTxtFld.setText("");
                    for(int i=0;i<tab.getRowCount();i++)
                    {
                        for(int j=0;j<tab.getColumnCount();j++)
                        {
                            model.setValueAt("",i,j);
                        }
                    }
                    list.clear();
                    model.fireTableDataChanged();
                    System.out.println("Sheetnames item count"+cmbSheetNames.getItemCount());
                    int mx=cmbSheetNames.getItemCount();
                    for(int cn=cmbSheetNames.getItemCount();cn>0;cn--)
                    {
                        System.out.println("cmbSheetNames item "+cmbSheetNames.getItemAt(cn-1));
                        if(cn!=1)
                        {
                            cmbSheetNames.removeItemAt(cn-1);
                        }
                    }
                }
                System.out.println("List index"+list.size());
                System.out.println("cmb sheet names"+cmbSheetNames.getItemCount());
                System.out.println("getSelectedFile()"+fileSelected);
                openTxtFld.setText(fileSelected.toString());
                openTxtFld.setToolTipText(fileSelected.toString());
                System.out.println("The File selected"+fileSelected);
                File inputWorkbook=new File(fileSelected.toString());
                try
                {
                    wbk1=Workbook.getWorkbook(inputWorkbook);
                    sheetCount=wbk1.getNumberOfSheets();
                    System.out.println("The Number of Sheets in the Opened Workbook "+sheetCount);
                    for(int k=0;k<sheetCount;k++)
                    {
                        sheet=wbk1.getSheet(k);
                        sheetColumns=wbk1.getSheet(k).getColumns();
                        sheetRows=wbk1.getSheet(k).getRows();
                        sheetData=new Object[sheetColumns][sheetRows];
                        acqStateData=new Object[sheetColumns][4];
                        for (int i = 0; i <sheetRows; i++)
                        {
                            for (int j = 0; j <sheetColumns; j++)
                            {
                                Cell cell = sheet.getCell(j,i);
                                CellType type = cell.getType();
                                if (cell.getType() == CellType.LABEL)
                                {
                                    sheetData[j][i]="\'"+cell.getContents()+"\'";
                                }
                                if (cell.getType() == CellType.NUMBER)
                                {
                                    sheetData[j][i]="\'"+cell.getContents()+"\'";
                                }
                                if (cell.getType() == CellType.EMPTY)
                                {
                                    sheetData[j][i]="\'"+cell.getContents()+"\'";
                                }
                            }
                        }
                        for (int j = 0; j <sheetColumns; j++)
                        {
                            for(int i=0;i<4;i++)
                            {
                                if(i==0)
                                {
                                    acqStateData[j][i]=sheetData[j][0];
                                }
                                else
                                {
                                    acqStateData[j][i]=" ";
                                }
                            }
                        }
                        list.add(k,acqStateData);
                        System.out.println("length of the list "+list.lastIndexOf(acqStateData));
                    }
                   for(int sc=0;sc<sheetCount;sc++)
                    {
                        cmbSheetNames.addItem(wbk1.getSheet(sc).getName());
                    }
                }
                catch (BiffException biffex)
                {
                    JOptionPane.showMessageDialog(frame, "Please select an excel sheet with only .xls extension","Error",JOptionPane.ERROR_MESSAGE);
                    openTxtFld.setText("");
                }
                catch (Exception wbk)
                {
                    JOptionPane.showMessageDialog(frame,wbk.getMessage(),"Error",JOptionPane.ERROR_MESSAGE);
                    openTxtFld.setText("");
                }
            }
            else {
                System.out.println("No Selection ");
                JOptionPane.showMessageDialog(frame,"Ha Ha Ha..... No File Selected","Alert",JOptionPane.INFORMATION_MESSAGE);
            }
        }
        if(e.getSource()==cmbSheetNames)
        {
            selectedSheet=(String)cmbSheetNames.getSelectedItem();
            selectedSheetIndex= cmbSheetNames.getSelectedIndex();
            System.out.print("The Selected Sheet Index "+selectedSheetIndex);
            System.out.print("The Selected Sheet Name "+selectedSheet);
            if(previousSheetIndex==0 && selectedSheetIndex==0 && cmbStateCheck==true )
            {
            }
            else
            {
                if(selectedSheetIndex!=0)
                {
                    sheet=wbk1.getSheet(selectedSheetIndex-1);
                    sheetColumns=wbk1.getSheet(selectedSheetIndex-1).getColumns();
                    sheetRows=wbk1.getSheet(selectedSheetIndex-1).getRows();
                    System.out.print("The Selected Sheet Rows "+sheetRows+" Columns : "+sheetColumns);
                    System.out.print("The Selected Sheet Rows "+sheetRows+" Columns : "+sheetRows);
                    if(previousSheetIndex!=0)
                    {
                        System.out.println("previous sheet index "+previousSheetIndex);
                        sheet=wbk1.getSheet(previousSheetIndex-1);
                        pSheetColumns=wbk1.getSheet(previousSheetIndex-1).getColumns();
                        System.out.println("previous sheet columns "+pSheetColumns);
                        tabStateData=new String[pSheetColumns][4];
                        for (int i=0; i < pSheetColumns; i++)
                        {
                            for (int j=0; j <4 ; j++)
                            {
                                tabStateData[i][j]=(String)model.getValueAt(i, j);
                                System.out.println("when changed   " +model.getValueAt(i, j)+" i "+i+" j "+j );
                            }
                        }
                        list.remove(previousSheetIndex-1);
                        list.add(previousSheetIndex-1,tabStateData);
                        model.getDataVector().removeAllElements();
                        sheet=wbk1.getSheet(selectedSheetIndex-1);
                        sheetColumns=wbk1.getSheet(selectedSheetIndex-1).getColumns();
                        System.out.println("selected sheet columns "+sheetColumns);
                        acqStateData_in=(Object[][])list.get(selectedSheetIndex-1);
                        String[] ssss=new String[4];
                        tabRowData=new String[sheetColumns][4];
                        for(int i=0;i<sheetColumns;i++)
                        {
                            for(int j=0;j<4;j++)
                            {
                                tabRowData[i][j]=(String)acqStateData_in[i][j];
                                ssss[j]=tabRowData[i][j];
                            }
                            model.insertRow(i,ssss);
                        }
                    }
                    else
                    {
                        sheet=wbk1.getSheet(selectedSheetIndex-1);
                        sheetColumns=wbk1.getSheet(selectedSheetIndex-1).getColumns();
                        System.out.println("selected sheet columns "+sheetColumns);
                        acqStateData_in=(Object[][])list.get(selectedSheetIndex-1);
                        String[] ssss=new String[4];
                        tabRowData=new String[sheetColumns][4];
                        for(int i=0;i<sheetColumns;i++)
                        {
                            for(int j=0;j<4;j++)
                            {
                                tabRowData[i][j]=(String)acqStateData_in[i][j];
                                ssss[j]=tabRowData[i][j];
                            }
                            model.insertRow(i,ssss);
                        }
                    }
                }
                else if(selectedSheetIndex==0 && previousSheetIndex!=0)
                {
                    System.out.println("previous sheet index "+previousSheetIndex);
                    sheet=wbk1.getSheet(previousSheetIndex-1);
                    pSheetColumns=wbk1.getSheet(previousSheetIndex-1).getColumns();
                    System.out.println("previous sheet columns "+pSheetColumns);
                    tabStateData=new String[pSheetColumns][4];
                    for (int i=0; i < pSheetColumns; i++)
                    {
                        for (int j=0; j <4 ; j++)
                        {
                            tabStateData[i][j]=(String)model.getValueAt(i, j);
                            System.out.println("when changed   " +model.getValueAt(i, j)+" i "+i+" j "+j );
                        }
                    }
                    list.remove(previousSheetIndex-1);
                    list.add(previousSheetIndex-1,tabStateData);
                    model.getDataVector().removeAllElements();
                }
            }
            previousSheetIndex=selectedSheetIndex;
            if(selectedSheetIndex!=0)
            {
                if(connection==null)
                {
                    JOptionPane.showMessageDialog(frame,"Please connect to database","Alert",JOptionPane.INFORMATION_MESSAGE);
                    usernameTxtFld.requestFocus(true);
                }
            }
        }
        if(e.getSource()==cmbTableNames)
        {
            selectedTable=(String)cmbTableNames.getSelectedItem();
            System.out.print("The Selected Table"+selectedTable);
            if(connection == null){
                JOptionPane.showMessageDialog(frame,"Please connect to database","Alert",JOptionPane.INFORMATION_MESSAGE);
                usernameTxtFld.requestFocus(true);
            }else{
                try
                {
                    statement=connection.createStatement();
                    if("ORACLE".equals(selectedDB)){
                        columnNamesQuery="select column_name from user_tab_columns where table_name='"+selectedTable+"'";
                    }else if("MySQL".equals(selectedDB)){
                        columnNamesQuery="select column_name from information_schema.COLUMNS where table_name='"+selectedTable+"'";
                    }
                    columnNamesRset=statement.executeQuery(columnNamesQuery);
                    cmbCheck=true;
                    while(columnNamesRset.next())
                    {
                        if(cmbCheck)
                        {
                            cmbColumnNames.removeAllItems();
                            cmbCheck=false;
                        }
                        cmbColumnNames.addItem(columnNamesRset.getString("column_name"));
                    }
                }
                catch (SQLException stmt)
                {
                    stmt.printStackTrace();
                }
            }
        }
        if(e.getSource()==loadButton)
        {
            if(connection == null){
                JOptionPane.showMessageDialog(frame,"Please connect to database","Alert",JOptionPane.INFORMATION_MESSAGE);
                usernameTxtFld.requestFocus(true);
            }else if(cmbSheetNames.getItemCount()==1){
                JOptionPane.showMessageDialog(frame,"Please browse an excel sheet with .xls extension","Alert",JOptionPane.INFORMATION_MESSAGE);
                usernameTxtFld.requestFocus(true);
            }
            else{
                model.fireTableDataChanged();
                if( selectedSheetIndex!=0)
                {
                    int numRows = tab.getRowCount();
                    int numCols = tab.getColumnCount();
                    Object[][] acqData_s=new Object[numRows][numCols] ;
                    javax.swing.table.TableModel model = tab.getModel();
                    for (int i=0; i < numRows; i++)
                    {
                        for (int j=0; j < numCols; j++)
                        {
                            acqData_s[i][j]=model.getValueAt(i, j);
                            System.out.println("acqData_s[i][j]"+acqData_s[i][j]);
                        }
                    }
                    list.remove(selectedSheetIndex-1);
                    list.add(selectedSheetIndex-1,acqData_s);
                }
                int loadOption=JOptionPane.showConfirmDialog(frame,"Do you want to continue","Alert",JOptionPane.YES_NO_OPTION);
                if(loadOption==JOptionPane.YES_OPTION)
                {
                    logFileDirPath=fchoose.getCurrentDirectory()+"\\NisumLoader_LogFile.txt";
                    badFileDirPath=fchoose.getCurrentDirectory()+"\\NisumLoader_BadFile.txt";
                    dataFileDirPath=fchoose.getCurrentDirectory()+"\\NisumLoader_DataFile.txt";
                    File logf=new File(logFileDirPath);
                    File badf=new File(badFileDirPath);
                    File dataf=new File(dataFileDirPath);
                    if(logf.exists())
                    {
                        logf.delete();
                    }
                    else
                    {
                        logf.getParentFile().mkdirs();
                    }

                    if(badf.exists())
                    {
                        badf.delete();
                    }
                    else
                    {
                        badf.getParentFile().mkdirs();
                    }

                    if(dataf.exists())
                    {
                        dataf.delete();
                    }
                    else
                    {
                        dataf.getParentFile().mkdirs();
                    }
                    try
                    {
                        FileWriter logfw=new FileWriter(logf);
                        FileWriter datafw=new FileWriter(dataf);
                        FileWriter badfw=new FileWriter(badf);

                        logfw.write("The Statistics of Data loading process : ");
                        logfw.write("\n----------------------------------------------------------");
                        datafw.write("The Statistics of Data loaded : ");
                        datafw.write("\n----------------------------------------------------------");
                        badfw.write("The Statistics of Data not loaded : ");
                        badfw.write("\n----------------------------------------------------------");

                        logfw.write("\n\nLog File Directory : "+logFileDirPath+"\n");
                        logfw.write("Data File Directory : "+dataFileDirPath+"\n");
                        logfw.write("Bad File Directory : "+badFileDirPath+"\n");

                        datafw.write("\n\nLog File Directory : "+logFileDirPath+"\n");
                        datafw.write("Data File Directory : "+dataFileDirPath+"\n");
                        datafw.write("Bad File Directory : "+badFileDirPath+"\n");

                        badfw.write("\n\nLog File Directory : "+logFileDirPath+"\n");
                        badfw.write("Data File Directory : "+dataFileDirPath+"\n");
                        badfw.write("Bad File Directory : "+badFileDirPath+"\n");

                        logfw.write("\n**************************************************************\n");
                        datafw.write("\n**************************************************************\n");
                        badfw.write("\n**************************************************************\n");

                        for(int k=0;k<sheetCount;k++)
                        {
                            sheet=wbk1.getSheet(k);
                            sheetColumns=wbk1.getSheet(k).getColumns();
                            sheetRows=wbk1.getSheet(k).getRows();
                            logfw.write("\nSheet Name : "+sheet.getName());
                            logfw.write("\nNumber of sheet Rows(including column names row) : "+sheetRows);
                            logfw.write("\nNumber of sheet Columns : "+sheetColumns);

                            datafw.write("\nSheet Name : "+sheet.getName());
                            datafw.write("\nNumber of sheet Rows(including column names row) : "+sheetRows);
                            datafw.write("\nNumber of sheet Columns : "+sheetColumns);

                            badfw.write("\nSheet Name : "+sheet.getName());
                            badfw.write("\nNumber of sheet Rows(including column names row) : "+sheetRows);
                            badfw.write("\nNumber of sheet Columns : "+sheetColumns);

                            System.out.print("The sheet Index "+k+" Sheet rows "+sheetRows+" sheet Columns "+sheetColumns);

                            sheetData=new Object[sheetColumns][sheetRows];
                            for (int i = 0; i <sheetRows; i++)
                            {
                                for (int j = 0; j <sheetColumns; j++)
                                {
                                    Cell cell = sheet.getCell(j,i);
                                    CellType type = cell.getType();
                                    if (cell.getType() == CellType.LABEL)
                                    {
                                        sheetData[j][i]="\'"+cell.getContents().trim()+"\'";
                                    }
                                    if (cell.getType() == CellType.NUMBER)
                                    {
                                        sheetData[j][i]="\'"+cell.getContents().trim()+"\'";
                                    }
                                    if (cell.getType() == CellType.EMPTY)
                                    {
                                        sheetData[j][i]="\'"+cell.getContents().trim()+"\'";
                                    }
                                }
                            }

                            acqData=(Object[][])list.get(k);
                            xs=new String[sheetColumns][4];
                            //xs holds the front end table states i.e, the sheet elements mapped to table names and column names
                            //javax.swing.table.TableModel model = tab.getModel();
                            for (int i=0; i < sheetColumns; i++)
                            {
                                for (int j=0; j < 4; j++)
                                {
                                    xs[i][j]=(String)acqData[i][j];
                                    System.out.println("  " + xs[i][j]);
                                }
                            }
                            //get all tables selected from xs[][] and add to allTabs list
                            for(int i=0;i<sheetColumns;i++)
                            {
                                allTabs.add(xs[i][1]);
                                //  System.out.println("All selected tables are :"+allTabs.get(i));
                            }
                            //getting distinct tables from allTabs list and add them to disTabs
                            for(int i=0;i<allTabs.size();i++)
                            {
                                if(!disTabs.contains(allTabs.get(i)))
                                {
                                    disTabs.add(allTabs.get(i));
                                }
                            }

                            for(int i=0;i<disTabs.size();i++)
                            {
                                int tempTabsCnt=0;
                                for(int j=0;j<allTabs.size();j++)
                                {
                                    if(allTabs.get(j)==disTabs.get(i))
                                    {
                                        tempTabsCnt=tempTabsCnt+1;
                                    }
                                }
                                Integer intObj=new Integer(tempTabsCnt);
                                disTabsCnt.add(intObj);
                            }
                            for(int h=0;h<disTabs.size();h++)
                            {
                                System.out.println("disTabs.get(0)"+disTabs.get(h).toString());
                            }
                            //getting respective rows from JTable table by table using distinct tables

                            if(!disTabs.get(0).toString().trim().equalsIgnoreCase(""))
                            {
                                for(int i=0;i<disTabs.size();i++)
                                {
                                    int v=Integer.parseInt(disTabsCnt.get(i).toString());
                                    disTabsRsCs=new String[v][4];
                                    int noOfSheetCols=0; //no.of columns in the sheet for a particular table
                                    for(int r=0;r<sheetColumns;r++)
                                    {
                                        if(xs[r][1]==disTabs.get(i))
                                        {
                                            for(int c=0;c<4;c++)
                                            {
                                                disTabsRsCs[noOfSheetCols][c]=xs[r][c];
                                            }
                                            noOfSheetCols=noOfSheetCols+1;
                                        }
                                    }
                                    disTabsSheetData=new Object[sheetRows][noOfSheetCols];
                                    String[][] ObjToStr=new String[sheetRows][sheetColumns];
                                    for(int f=0;f<sheetRows;f++)
                                    {
                                        for(int j=0;j<sheetColumns;j++)
                                        {
                                            ObjToStr[f][j]=(String)sheetData[j][f];
                                        }
                                    }
                                    for(int j=0;j<noOfSheetCols;j++)
                                    {
                                        for(int sc=0;sc<sheetColumns;sc++)
                                        {
                                            if(trimQuotes(disTabsRsCs[j][0]).equalsIgnoreCase(trimQuotes(ObjToStr[0][sc])))
                                            {
                                                for(int sr=0;sr<sheetRows;sr++)
                                                {
                                                    disTabsSheetData[sr][j]=ObjToStr[sr][sc];
                                                }
                                            }
                                        }
                                    }
                                    try
                                    {
                                        if(connection!=null)
                                        {
                                            statement =connection.createStatement();
                                            connection.setAutoCommit(false);
                                        }
                                        String queryColumns="";
                                        for (int qc=0; qc < noOfSheetCols; qc++)
                                        {
                                            if(qc==noOfSheetCols-1)
                                            {
                                                queryColumns=queryColumns+disTabsRsCs[qc][2] ;
                                            }
                                            else
                                            {
                                                queryColumns=queryColumns+disTabsRsCs[qc][2]+",";
                                            }
                                            System.out.println("queryColumns string "+ queryColumns);
                                        }
                                        String insRowData="";
                                        datafw.write("\n\n<<<Table Name>>>: "+disTabsRsCs[0][1]);
                                        datafw.write("\n\n<<Queried Column Names>>: "+queryColumns+"\n\n");
                                        datafw.write("The Loaded Rows are :\n");
                                        datafw.write("............................\n\n");


                                        badfw.write("\n\n<<<Table Name>>>: "+disTabsRsCs[0][1]);
                                        badfw.write("\n\n<<Queried Column Names>>: "+queryColumns+"\n\n");
                                        badfw.write("The Rows that are not loaded :\n");
                                        badfw.write("............................\n\n");

                                        for (int rd = 0; rd <sheetRows; rd++)
                                        {
                                            for (int j = 0; j <noOfSheetCols; j++)
                                            {
                                                if(!disTabsRsCs[j][3].trim().equalsIgnoreCase("null") && !disTabsRsCs[j][3].trim().equals(""))
                                                {
                                                    expForSheetData=replaceWord(disTabsRsCs[j][3],trimQuotes((String)disTabsRsCs[j][0]),(String)disTabsSheetData[rd][j]);
                                                    System.out.println("Expression..."+expForSheetData+"Expression Before"+disTabsRsCs[j][3]+"....");
                                                }
                                                else
                                                {
                                                    expForSheetData=(String)disTabsSheetData[rd][j];
                                                }

                                                if (j==noOfSheetCols-1)
                                                {
                                                    insRowData=insRowData+expForSheetData;
                                                }
                                                else
                                                {
                                                    insRowData=insRowData+expForSheetData+",";
                                                }
                                            }
                                            if(rd!=0)
                                            {
                                                insertStmtQuery="insert into "+disTabsRsCs[0][1]+"("+queryColumns+") select "+insRowData+" from dual";
                                                try
                                                {
                                                    savepoint1 = connection.setSavepoint();
                                                    statement.executeUpdate(insertStmtQuery);
                                                    totRowCount=totRowCount+1;
                                                    datafw.write("<Row> :"+rd+"->"+insRowData+"\n");
                                                }
                                                catch (Exception sv)
                                                {
                                                    sv.printStackTrace();
                                                    connection.rollback(savepoint1);
                                                    badRowCount=badRowCount+1;
                                                    if(badRowCount>=1)
                                                    {
                                                        badfw.write("<Row> :"+rd+"->"+insRowData);
                                                        badfw.write("\n\nReason :-> "+sv.getMessage()+"\n");
                                                    }

                                                    System.out.println("Bad Row Number :"+badRowCount);
                                                }
                                            }
                                            System.out.println("Insert stmt is : "+insertStmtQuery);
                                            System.out.println("Row Data : "+ insRowData);
                                            insRowData="";
                                        }

                                        System.out.println("After Insert stmt...disTabsRsCs"+disTabsRsCs[0][1]);
                                        logfw.write("\n\n<<<Table Name>>>: "+disTabsRsCs[0][1]);
                                        logfw.write("\n\nTotal Number of Records loaded : "+totRowCount);
                                        logfw.write("\nTotal Number of Bad Records : "+badRowCount);

                                        datafw.write("\nTotal Number of Records loaded : "+totRowCount);
                                        datafw.write("\nTotal Number of Bad Records : "+badRowCount);

                                        if (badRowCount>=1)
                                        {
                                            badfw.write("Total Number of Bad Records : "+badRowCount);
                                        }
                                        else
                                        {
                                            badfw.write("Zero Bad Records");
                                        }
                                        badRowCount=0;
                                        totRowCount=0;
                                        insertStmtQuery="";
                                        connection.commit();
                                        System.out.println("end of load data button");
                                    }
                                    catch (SQLException inStmt)
                                    {
                                        System.out.println("SQL Exception :"+inStmt.getMessage());
                                    }
                                    datafw.write("\n");
                                    badfw.write("\n");
                                    logfw.write("\n");
                                }//end loop of tables in a sheet

                                allTabs.clear();
                                disTabs.clear();
                                disTabsCnt.clear();
                            }//end of cnttab if block
                            else
                            {
                                logfw.write("\n\nZero-mappings : No Mappings Found for this Sheet\n");
                                datafw.write("\n\nZero-mappings : No Mappings Found for this Sheet\n");
                                badfw.write("\n\nZero-mappings : No Mappings Found for this Sheet\n");
                            }
                            logfw.write("\n**************************************************************\n");
                            datafw.write("\n**************************************************************\n");
                            badfw.write("\n**************************************************************\n");
                            allTabs.clear();
                            disTabs.clear();
                            disTabsCnt.clear();
                        } //end loop of sheets

                        logfw.flush();
                        logfw.close();

                        datafw.flush();
                        datafw.close();

                        badfw.flush();
                        badfw.close();

                        JOptionPane.showMessageDialog(frame,"Loading Process Completed","Status",JOptionPane.INFORMATION_MESSAGE);
                    } // end of fw try block
                    catch (IOException fwexcep)
                    {
                        System.out.println(fwexcep.getMessage());
                    }
                }
                else if(loadOption==JOptionPane.NO_OPTION)
                {
                    frame.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
                }
            }
        }
    }
    //mouse abstract menthods
    public void mousePressed(MouseEvent e) { }
    public void mouseReleased(MouseEvent e) {}
    public void mouseEntered(MouseEvent e) {  }
    public void mouseExited(MouseEvent e) {}
    public void mouseClicked(MouseEvent me) {
        if(me.getSource()==openTxtFld)
        {
            JOptionPane.showMessageDialog(frame,"Please Click Browse Button, You cannot enter here","Alert",JOptionPane.INFORMATION_MESSAGE);
            browseButton.requestFocus(true);
        }
    }

    //windowlistener abstract methods
    public void  windowActivated(WindowEvent e) {}
    public void  windowClosed(WindowEvent e) {}
    public void  windowClosing(WindowEvent e) {
        int winC=JOptionPane.showConfirmDialog(frame,"Do you want to Exit","Alert",JOptionPane.YES_NO_OPTION);
        if(winC==JOptionPane.YES_OPTION)
        {
            System.exit(0);
        }
        else if(winC==JOptionPane.NO_OPTION)
        {
            frame.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
        }
    }
    public void  windowDeactivated(WindowEvent e) {}
    public void  windowDeiconified(WindowEvent e) {}
    public void  windowIconified(WindowEvent e) {}
    public void  windowOpened(WindowEvent e) {}
    public static void main(String[] args)
    {
        try
        {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        new ExampleScreen_1();
    }
}