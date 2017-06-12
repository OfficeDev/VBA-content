---
title: About Connecting to Data in Visio
ms.prod: visio
ms.assetid: 2057123f-faeb-f705-5fe7-75d3b76fa1a5
ms.date: 06/08/2017
---


# About Connecting to Data in Visio

 **Note**  Data-connectivity features are available only to licensed users of Microsoft Visio Professional 2013.

There are four aspects of data connectivity in Visio:

- Connecting to a data source
    
- Linking shapes to data
    
- Displaying linked data graphically
    
- Refreshing linked data that has changed in the data source, updating linked shapes, and resolving any subsequent conflicts that may arise
    
Typically, you approach these aspects in the order in which they are listed; that is, you first connect your Visio drawing to a data source, then link shapes in your drawing to data in the data source, display the data in linked shapes graphically, and refresh the linked data when necessary. 
Each of these aspects has new objects and members associated with it in the Visio object model. This topic deals with the first of these aspects, connecting your Visio drawing to a data source. For more information about the other aspects of data connectivity, see the following topics: 

-  [About Linking Shapes to Data](about-linking-shapes-to-data.md)
    
-  [About Displaying Data Graphically](about-displaying-data-graphically-visio.md)
    
To connect your Visio drawing to a data source programmatically, you can use the Visio API for data connectivity, which includes the following objects and their associated members:

-  **[DataRecordsets](datarecordsets-object-visio.md)** collection
    
-  **[DataRecordset](datarecordset-object-visio.md)** object
    
-  **[DataConnection](dataconnection-object-visio.md)** object
    
-  **[DataRecordsetChangedEvent](datarecordsetchangedevent-object-visio.md)** object
    
-  **[DataColumns](datacolumns-object-visio.md)** collection
    
-  **[DataColumn](datacolumn-object-visio.md)** object
    

## About Data Recordsets and Data Connections

Each Visio  **Document** object has a **DataRecordsets** collection, which is empty until you make a connection to a data source. To connect a Visio document to a data source, you add a **DataRecordset** object to the **DataRecordsets** collection of the document. A **DataRecordset** object in turn has a **DataColumns** collection of **DataColumn** objects, each of which is mapped to a corresponding column (field) in the data source.

Data sources you can connect to include Excel spreadsheets, Access and SQL Server databases, SharePoint lists, and other OLEDB or ODBC data sources, such as an Oracle database. When you add a  **DataRecordset** object by connecting to one of these data sources, Visio abstracts the connection in a **DataConnection** object, and the **DataRecordset** object is said to be connected.

You can also add a  **DataRecordset** object by using an XML file that conforms to the ADO Classic (ADO version 2.x) Data Recordset XML schema as the data source. The resulting **DataRecordset** object is said to be connection-less. The connection between a data source and a **DataRecordset** object only goes one way—from the data source to the Visio drawing. If data in the source changes, you can refresh the data in the drawing to reflect those changes. You cannot, however, make changes in the data in the drawing and then push those changes back to the data source.


## Adding DataRecordset Objects

To add a  **DataRecordset** object to the **DataRecordsets** collection, you can use one of the following three methods, depending on the data source you want to connect to and whether you want to pass the method a connection string and query command string or a saved Office Data Connection (ODC) file that contains the connection and query information:


-  **[DataRecordsets.Add](datarecordsets-add-method-visio.md)**
    
-  **[DataRecordsets.AddFromConnectionFile](datarecordsets-addfromconnectionfile-method-visio.md)**
    
-  **[DataRecordsets.AddFromXML](datarecordsets-addfromxml-method-visio.md)**
    
The following Visual Basic for Applications (VBA) sample macro shows how you might use the  **Add** method to connect a Visio drawing to data in an Excel worksheet, in this case, in the ORGDATA.XLS sample workbook that is included with Visio:




```vb
Public Sub AddDataRecordset()

    Dim strConnection As String
    Dim strCommand As String
    Dim strOfficePath As String
    Dim vsoDataRecordset As Visio.DataRecordset
    strOfficePath = Visio.Application.Path    
    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" _
                       &; "User ID=Admin;" _
                       &; "Data Source=" + strOfficePath + "SAMPLES\1033\ORGDATA.XLS;" _
                       &; "Mode=Read;" _
                       &; "Extended Properties=""HDR=YES;IMEX=1;MaxScanRows=0;Excel 12.0;"";" _
                       &; "Jet OLEDB:Engine Type=34;"
    strCommand = "SELECT * FROM [Sheet1$]"
    Set vsoDataRecordset = ActiveDocument.DataRecordsets.Add(strConnection, strCommand, 0, "Data")

End Sub
```


 **Note**  If you run this code on non-English builds of Visio, the path to the file ORGDATA.XLS, shown here as "Samples\1033\ORGDATA.XLS," will be different. Substitute the correct path for your version of Visio in your code.

The  **Add** method returns a **DataRecordset** object and takes four parameters:


-  _ConnectionIDOrString_ The ID of an existing **DataConnection** object or the connection string to specify a new data-source connection. If you pass the ID of an existing **DataConnection** object that is currently being used by one or more other data recordsets, the data recordsets become a transacted group recordset. All data recordsets in the group are refreshed whenever a data-refresh operation occurs. You can determine an appropriate connection string by first using the **Data Selector Wizard** in the Visio user interface (UI) to make the same connection, recording a macro while running the wizard, and then copying the connection string from the macro code.
    
-  _CommandString_ The string that specifies the database table or Excel worksheet and specifies the fields (columns) within the table or worksheet that contain the data you want to query. The command string is also passed to the **[DataRecordset.Refresh](datarecordset-refresh-method-visio.md)** method when the data in the recordset is refreshed.
    
-  _AddOptions_ A combination of one or more values from the **[VisDataRecordsetAddOptions](visdatarecordsetaddoptions-enumeration-visio.md)** enumeration. These values specify certain data recordset behaviors, and make it possible, for example, to prevent the queried data in the recordset from appearing in the **External Data** window in the Visio UI or from being refreshed by user actions. Afteryou assign this value, you cannot change it for the duration of the **DataRecordset** object.
    
-  _Name_ An optional string that gives the data recordset a display name. If you specify for data from the recordset to be displayed in the **External Data** window in the Visio UI, the name you pass appears on the tab of that window that corresponds to the data recordset. In our example, there is no existing data connection, so for the first parameter of the **Add** method, we pass _strConnection_, the connection string we defined. For the second parameter, we pass  _strCommand_, the command string we defined, which directs Visio to select all columns from the worksheet we specified. For the third parameter of the  **Add** method, we pass zero to specify default behavior of the data recordset, and for the last parameter, we pass _Org Data_, the display name we defined for the data recordset.
    
The following sample code shows how to get the  **DataConnection** object that was created when we added a **DataRecordset** object to the **DataRecordsets** collection. It prints the connection string associated with the **DataConnection** object in the **Immediate** window by accessing the **[ConnectionString](dataconnection-connectionstring-property-visio.md)** property of the **DataConnection** object.




```vb
Public Sub GetDataConnectionObject(vsoDataRecordset As Visio.DataRecordset) 
 
    Dim vsoDataConnection As DataConnection 
    Set vsoDataConnection = vsoDataRecordset.DataConnection 
    Debug.Print vsoDataConnection.ConnectionString 
 
End Sub
```

Just as you can get the connection string associated with a  **DataConnection** object by accessing its **ConnectionString** property, you can get the command string associated with a **DataRecordset** object by accessing its **[CommandString](datarecordset-commandstring-property-visio.md)** property. Both of these properties are assignable, so you can change the data source associated with a **DataRecordset** object or the query associated with a **DataConnection** object at any time, although changes are not reflected in your drawing until you refresh the data. For more information about refreshing data, see [About Linking Shapes to Data](about-linking-shapes-to-data.md). 


## Accessing Data in Data Recordsets Programmatically

When you import data, Visio assigns integer row IDs, starting with the number 1, to each data row in the resulting data recordset, based upon the order of rows in the original data source. Visio uses data row IDs to track the rows when they are linked to shapes and when data is refreshed. If you want to access data rows programmatically, you must use these data row IDs. For information about how data-refresh operations affect row order, see [About Linking Shapes to Data](about-linking-shapes-to-data.md). 

You can use the  **[DataRecordset.GetDataRowIDs](datarecordset-getdatarowids-method-visio.md)** method to get an array of the IDs of all the rows in a data recordset, where each row represents a single data record. The **GetDataRowIDs** method takes as its parameter a criteria string, which is a string that conforms to the guidelines specified in the ActiveX Data Object (ADO) API for setting the **ADO.Filter** property. By specifying appropriate criteria and using AND and OR operators to separate clauses, you can filter the information in the data recordset to return only certain data recordset rows selectively. To apply no filter (that is, to get all the rows), pass an empty string (""). For more information about criteria strings, see the [Filter Property](http://msdn.microsoft.com/en-us/library/ms676691%28VS.85%29.aspx) topic in the ADO 2.x API Reference. After you retrieve the data-row IDs, you can use the **[DataRecordset.GetRowData](datarecordset-getrowdata-method-visio.md)** method to get all the data stored in each column in the data row. For more information about data columns, see [About Linking Shapes to Data](about-linking-shapes-to-data.md). 

The following sample code shows how to use the  **GetDataRowIDs** and **GetRowData** methods to return the row ID of each row and then get the data stored in each column in every row of the data recordset passed in. It uses two nested **For…Next** loops to iterate through all the rows in the recordset and then, for each row, iterate through all the columns in that row. The code displays the information returned in the **Immediate** window. Note that you pass an empty string to the **GetDataRowIDs** method to bypass filtering and get all the rows in the recordset. After you call the procedure, note that the first set of data shown (corresponding to the first data row) contains the headings for all the data columns in the data recordset.




```vb
Public Sub GetDataRecords(vsoDataRecordset As Visio.DataRecordset)

    Dim lngRowIDs() As Long
    Dim lngRow As Long
    Dim lngColumn As Long
    Dim varRowData As Variant

    'Get the row IDs of all the rows in the recordset
    lngRowIDs = vsoDataRecordset.GetDataRowIDs("")

    'Iterate through all the records in the recordset.
    For lngRow = LBound(lngRowIDs) To UBound(lngRowIDs)
        varRowData = vsoDataRecordset.GetRowData(lngRow)

        'Print a separator between rows
        Debug.Print "------------------------------"

       'Print the data stored in each column of a particular data row.
        For lngColumn = LBound(varRowData) To UBound(varRowData)
            Debug.Print vsoDataRecordset.DataColumns(lngColumn + 1).Name _
               & Trim(Str(lngColumn)) & " = " & VarRowData(lngColumn)
        Next lngColumn
    Next lngRow

End Sub
```


