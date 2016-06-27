
# ImportExportData Macro Action

 **Last modified:** July 28, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You can use the  **ImportExportData** action to import or export data between the current Access database (.mdb or .accdb) or Access project (.adp) and another database. For Microsoft Access databases, you can also link a table to the current Access database from another database. With a linked table, you have access to the table's data while the table itself remains in the other database.


 **Note**  This action will not be allowed if the database is not trusted. For more information about enabling macros, see the links in the  **See Also** section of this article.


## Settings

The  **ImportExportData** action has the following arguments.



|**Action argument**|**Description**|
|:-----|:-----|
|**Transfer Type**|The type of transfer you want to make. Select  **Import**,  **Export**, or  **Link** in the **Transfer Type** box in the **Action Arguments** section of the Macro Builder pane. The default is **Import**.
 **Note**  The  **Link** transfer type is not supported for Access projects (.adp).

|
|**Database Type**|The type of database to import from, export to, or link to. You can select  **Microsoft Access** or one of a number of other database types in the **Database Type** box. The default is **Microsoft Access**.|
|**Database Name**|The name of the database to import from, export to, or link to. Include the full path. This is a required argument. For types of databases that use separate files for each table, such as FoxPro, Paradox, and dBASE, enter the directory containing the file. Enter the file name in the  **Source** argument (to import or link) or the **Destination** argument (to export). For ODBC databases, type the full Open Database Connectivity (ODBC) connection string.To see an example of a connection string, link an external table to Access: 
<ol xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p /></li><li><p>In the <span class="ui">Get External Data</span> dialog box, enter  the path of your source database in the <span class="ui">File name</span> box.</p></li><li><p>Click <span class="ui">Link to the data source by creating a linked table</span>, and click <span class="ui">OK</span>.</p></li><li><p>Select a table in the <span class="ui">Link Tables</span> dialog box, and click <span class="ui">OK</span>.</p></li></ol>Open the newly linked table in Design view and view the table properties by clicking  **Property Sheet** on the **Design** tab, under **Tools**. The text in the  **Description** property setting is the connection string for this table.For more information on ODBC connection strings, see the Help file or other documentation for the ODBC driver of this type of ODBC database.|
|**Object Type**|The type of object to import or export. If you select  **Microsoft Access** for the **Database Type** argument, you can select **Table**,  **Query**,  **Form**,  **Report**,  **Macro**,  **Module**,  **Data Access Page**,  **Server View**,  **Diagram**,  **Stored Procedure**, or  **Function** in the **Object Type** box. The default is **Table**. If you select any other type of database, or if you select  **Link** in the **Transfer Type** box, this argument is ignored. If you are exporting a select query to an Access database, select **Table** in this argument to export the result set of the query, and select **Query** to export the query itself. If you are exporting a select query to another type of database, this argument is ignored and the result set of the query is exported.|
|**Source**|The name of the table, select query, or Access object that you want to import, export, or link. For some types of databases, such as FoxPro, Paradox, or dBASE, this is a file name. Include the file name extension (such as .dbf) in the file name. This is a required argument.|
|**Destination**|The name of the imported, exported, or linked table, select query, or Access object in the destination database. For some types of databases, such as FoxPro, Paradox, or dBASE, this is a file name. Include the file name extension (such as .dbf) in the file name. This is a required argument. If you select  **Import** in the **Transfer Type** argument and **Table** in the **Object Type** argument, Access creates a new table containing the data in the imported table. If you import a table or other object, Access adds a number to the name if it conflicts with an existing name. For example, if you import Employees and Employees already exists, Access renames the imported table or other object Employees1. If you export to an Access database or another database, Access automatically replaces any existing table or other object that has the same name.|
|**Structure Only**|Specifies whether to import or export only the structure of a database table without any of its data. Select  **Yes** or **No**. The default is  **No**.|

## Remarks

You can import and export tables between Access and other types of databases. You can also export Access select queries to other types of databases. Access exports the result set of the query in the form of a table. You can import and export any Access database object if both databases are Access databases.

If you import a table from another Access database (.mdb or .accdb) that's a linked table in that database, it will still be linked after you import it. That is, the link is imported, not the table itself.

If the database you're accessing requires a password, a dialog box appears when you run the macro. Type the password in this dialog box.

The  **ImportExportData** action is similar to the commands on the **External Data** tab, under **Import** or **Export**. You can use these commands to select a source of data, such as an Access database or another type of database, a spreadsheet, or a text file. If you select a database, one or more dialog boxes appear in which you select the type of object to import or export (for Access databases), the name of the object, and other options, depending on the database you are importing from or exporting or linking to. The arguments for the  **ImportExportData** action reflect the options in these dialog boxes.

If you want to supply index information for a linked dBASE table, first link the table: 


1. 
    
2. Click  **dBASE File**.
    
3. In the  **Get External Data** dialog box, enter the path for the dBASE file in the **File name** box.
    
4. Click  **Link to the data source by creating a linked table**, then click  **OK**.
    
5. Specify the indexes in the dialog boxes for this command. Access stores the index information in a special information (.inf) file, located in the Microsoft Office folder.
    
6. You can then delete the link to the linked table. 
    
The next time you use the  **ImportExportData** action to link this dBASE table, Access uses the index information that you've specified.


 **Note**  If you query or filter a linked table, the query or filter is case-sensitive.

To run the  **ImportExportData** action in a Visual Basic for Applications (VBA) module, use the **TransferDatabase** method of the **DoCmd** object.

