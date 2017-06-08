---
title: DataRecordsets.AddFromConnectionFile Method (Visio)
keywords: vis_sdr.chm16360275
f1_keywords:
- vis_sdr.chm16360275
ms.prod: visio
api_name:
- Visio.DataRecordsets.AddFromConnectionFile
ms.assetid: 7118bd4d-484b-dc22-e6f8-925376a5a67a
ms.date: 06/08/2017
---


# DataRecordsets.AddFromConnectionFile Method (Visio)

Adds a  **[DataRecordset](datarecordset-object-visio.md)** object to the **[DataRecordsets](datarecordsets-object-visio.md)** collection by using the connection and query information contained in an Office Data Connection (ODC) file to connect to and retrieve data from an OLEDB or ODBC data source.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **AddFromConnectionFile**( **_FileName_** , **_AddOptions_** , **_Name_** )

 _expression_ An expression that returns a **DataRecordsets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the ODC connection file to use. |
| _AddOptions_|Required| **Long**|Options that determine properties of the data recordset to be added. A combination of one or more enumerated value from  **[VisDataRecordsetAddOptions](visdatarecordsetaddoptions-enumeration-visio.md)** . For more information, see Remarks.|
| _Name_|Optional| **String**|Assigns a display name to the  **DataRecordset** object being added.|

### Return Value

DataRecordset


## Remarks

For the FileName parameter, pass the name and full path of an ODC file that contains a connection string that specifies how to connect to an OLEDB or ODBC data source and a query string that specifies how to extract the desired data from the data source. 

An ODC file uses HTML and XML to store connection and query information. You can view or edit the contents of the file in any text editor. ODC files have the .odc file name extension. You can use the Data Connection Wizard in Microsoft Access or Microsoft Excel to create an ODC file that will connect to and retrieve the data you want.

The AddOptions parameter can be a combination of one or more of the following values from the  **VisDataRecordsetAddOptions** enumeration, which is declared in the Visio type library. The default is zero (0), which specifies that none of the options be set.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDataRecordsetNoExternalDataUI**|1|Prevents data in the new data recordset from being displayed in the  **External Data** window.|
| **visDataRecordsetNoRefreshUI**|2|Prevents the data recordset from being included in the refresh operation and displayed in the  **Refresh Data** dialog box.|
| **visDataRecordsetNoAdvConfig**|4|Prevents the data recordset from being displayed in the  **Configure Refresh** dialog box.|
| **visDataRecordsetDelayQuery**|8|Adds a data recordset but does not execute the CommandString query until the next time you call the  **Refresh** method.|
| **visDataRecordsetDontCopyLinks**|16|Adds a data recordset, but shape-data links are not cut or copied.|
 Once you assign these values, you cannot change them for the life of the **DataRecordset** object.

The Name argument is an optional string that lets you assign the data recordset a display name. If you specify that the  **External Data** window display in the Visio UI, the name you pass for this argument appears on the tab of the **External Data** window that corresponds to the data recordset added.

If the  **AddFromConnectionFile** method succeeds, it performs the following actions:


- Creates a  **DataRecordset** object and assigns it the name specified in the Name parameter. If you do not specify a name, Visio assigns the data recordset the name of the database table that is the source of the data.
    
- Associates a new or existing  **DataConnection** object with the **DataRecordset** object.
    
- Executes the query string specified in the command string within the ODC file and retreives the resulting data.
    
- Maps the data types of the columns of the data source to equivalent Visio data types, while filtering the results to remove data-source columns that cannot be linked to Visio shapes because they have no equivalent Visio data type. 
    
-  Assigns a row ID to each row in the data recordset. For more information about row IDs, see the **[DataRecordset.GetDataRowIDs](datarecordset-getdatarowids-method-visio.md)** property topic.
    



 **Note**  The  **AddFromConnectionFile** method fails and return an exception if it encounters network connection errors, network time outs, or database permission errors.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how you might use the  **AddFromConnectionString** method to connect a Visio drawing to connect to data in the Products table of the Northwind database that is supplied with Microsoft Access. Before running this sample code, use the Data Connection Wizard to create an ODC file, and replace the value of the _strFile_ variable with the full path to and file name of the ODC file you created. Optionally, supply a different value for the _strName_ variable.


```vb
Public Sub AddFromConnectionFile_Example() 
 
    Dim strFile As String 
    Dim strName As String 
    Dim vsoDataRecordset As Visio.DataRecordset 
 
    strFile = "C:\Users\username \Documents\My Data Sources\Northwind.mdb Products.odc" 
 
    strName = "Data from ODC" 
 
    Set vsoDataRecordset = ThisDocument.DataRecordsets.AddFromConnectionFile(strFile, 0, strName) 
 
End Sub
```


