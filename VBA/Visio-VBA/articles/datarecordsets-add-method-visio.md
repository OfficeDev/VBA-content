---
title: DataRecordsets.Add Method (Visio)
keywords: vis_sdr.chm16316005
f1_keywords:
- vis_sdr.chm16316005
ms.prod: visio
api_name:
- Visio.DataRecordsets.Add
ms.assetid: 9eb136ce-d543-75c3-3a72-cb23dfc8df78
ms.date: 06/08/2017
---


# DataRecordsets.Add Method (Visio)

Adds a  **[DataRecordset](datarecordset-object-visio.md)** object to the **[DataRecordsets](datarecordsets-object-visio.md)** collection by connecting to and retrieving data from an OLEDB or ODBC data source.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Add**( **_ConnectionIDOrString_** , **_CommandString_** , **_AddOptions_** , **_Name_** )

 _expression_ A variable that represents a **DataRecordsets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ConnectionIDOrString_|Required| **Variant**|The ID of an existing  **[DataConnection](dataconnection-object-visio.md)** object or the connection string to specify a new data-source connection.|
| _CommandString_|Required| **String**|Query string that specifies the database table or Excel worksheet and the fields (columns) within the table or worksheet that contain the data you want to query.|
| _AddOptions_|Required| **Long**|Options that determine properties of the data recordset to add. A combination of one or more enumerated value from  **[VisDataRecordsetAddOptions](visdatarecordsetaddoptions-enumeration-visio.md)** . For more information, see Remarks.|
| _Name_|Optional| **String**|Assigns a display name to the  **DataRecordset** object being added.|

### Return Value

DataRecordset


## Remarks

You can determine an appropriate connection string to pass to the ConnectionIDOrString parameter by first using the  **Data Selector Wizard** in the Visio user interface (UI) to make the same connection, recording a macro while running the wizard, and then copying the connection string from the macro code.

An easy way to reuse an existing data connection is to pass the  **DataConnection** property value of an existing **DataRecordset** object for the ConnectionIDOrString parameter. Use the following syntax:




```
NewDataRecordset  = DataRecordsets.Add(ExistingDataRecordset .DataConnection.ID, CommandString, AddOptions, Name)
```

For the ConnectionIDOrString parameter, if you pass the ID of an existing  **DataConnection** object that is currently being used by one or more other data recordsets, all the data recordsets become a _transacted group recordset_ . All data recordsets in the group are refreshed whenever a data-refresh operation occurs.

The AddOptions parameter can be a combination of one or more of the following values from the  **VisDataRecordsetAddOptions** enumeration, which is declared in the Visio type library. The default is zero (0), which specifies that none of the options be set.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDataRecordsetNoExternalDataUI**|1|Prevents data in the new data recordset from being displayed in the  **External Data** window.|
| **visDataRecordsetNoRefreshUI**|2|Prevents the data recordset from being displayed in the  **Refresh Data** dialog box.|
| **visDataRecordsetNoAdvConfig**|4|Prevents the data recordset from being displayed in the  **Configure Refresh** dialog box.|
| **visDataRecordsetDelayQuery**|8|Adds a data recordset but does not execute the CommandString query until the next time you call the  **Refresh** method.|
| **visDataRecordsetDontCopyLinks**|16|Adds a data recordset, but shape-data links are not copied to the Clipboard when shapes are copied or cut.|
 Once you assign these values, you cannot change them for the life of the **DataRecordset** object.

The Name parameter is an optional string that lets you assign the data recordset a display name. If you specify that the  **External Data** window be displayed in the Visio UI, the name you pass for this argument appears on the tab of the **External Data** window that corresponds to the data recordset added.

If the  **Add** method succeeds, it performs the following actions:


- Creates a  **DataRecordset** object and assigns it the name specified in the Name parameter. If you do not specify a name, Visio assigns the data recordset the name of the database table that is the source of the data.
    
- Associates a new or existing  **DataConnection** object with the **DataRecordset** object.
    
- Opens the  **External Data** window in the Visio UI, unless **visDataRecordsetNoExternalDataUI** is set.
    
Unless you pass  **visDataRecordsetDelayQuery** as part of the AddOptions parameter, the **Add** method also does the following:


- Executes the query string specified in the CommandString parameter and retreive the resulting data.
    
- Maps the data types of the columns of the data source to equivalent Visio data types, while filtering the results to remove data-source columns that cannot be linked to Visio shapes because they have no equivalent Visio data type. In particular, you cannot import binary data or esoteric data types such as  **UserDefined** , **Chapter** , and **IDispatch** .
    
-  Assigns a row ID to each row in the data recordset. For more information about row IDs, see **[DataRecordset.GetDataRowIDs ](datarecordset-getdatarowids-method-visio.md)** topic.
    



 **Note**  The  **Add** method fails and returns an exception if it encounters network connection errors, network time outs, or database permission errors. If the **visDataRecordsetDelayQuery** option is set, under the same circumstances **Add** may successfully add a new data recordset, but refresh may fail.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how you might use the  **Add** method to connect a Visio drawing to data in ORGDATA.XLS, a Microsoft Office Excel workbook that is installed at C:\Program Files\Microsoft Office\OFFICE12\SAMPLES\1033\ when you install Visio at the default file path.

In this example, there is no existing data connection, so for the first parameter of the  **Add** method, we pass _strConnection_ , the connection string. For the second parameter, we pass _strCommand_ , the command string, which directs Visio to select all columns from the worksheet we specify. For the third parameter of the **Add** method, we pass zero to specify default behavior of the data recordset, and for the last parameter, we pass _"Org Data"_ , the display name we want to assign to the data recordset.




```vb
Public Sub AddDataRecordset_Example() 
 
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
 
    Set vsoDataRecordset = ActiveDocument.DataRecordsets.Add(strConnection, strCommand, 0, "Org Data") 
 
End Sub
```


