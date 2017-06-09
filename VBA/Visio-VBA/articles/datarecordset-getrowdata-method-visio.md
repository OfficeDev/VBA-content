---
title: DataRecordset.GetRowData Method (Visio)
keywords: vis_sdr.chm16460315
f1_keywords:
- vis_sdr.chm16460315
ms.prod: visio
api_name:
- Visio.DataRecordset.GetRowData
ms.assetid: 969d7702-e78c-736f-87d8-c8e7e8c5a778
ms.date: 06/08/2017
---


# DataRecordset.GetRowData Method (Visio)

Gets the data in all columns in the specified row.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetRowData**( **_DataRowID_** )

 _expression_ An expression that returns a **DataRecordset** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRowID_|Required| **Long**|The ID of the row in the data recordset from which you want to retrieve data.|

### Return Value

Variant()


## Remarks

To get the row IDs of all the rows in the data recordset, use the  **[GetDataRowIDs](datarecordset-getdatarowids-method-visio.md)** method. See the example in this topic.

If you pass a row ID of zero for the DataRowID parameter, the  **GetRowData** method returns the names of the columns in the data recordset. If you pass any other valid row ID than zero, the **GetRowData** method returns values for all the columns in the specified row, in the same order as the column names that the method returns when you pass zero.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how you can use the  **GetDataRowIDs** and **GetRowData** methods to return the row ID of each row and then get the data stored in each column in every row of the specified data recordset. The macro gets the count of all data recordsets associated with the current document and gets row data for the one most recently added. It iterates through all the rows in the data recordset and then, for each row, iterates through all the columns in that row. The code displays the information returned in the **Immediate** window.

Before running this macro, create at least one data recordset in the current document.

Note that the macro passes an empty string to the  **GetDataRowIDs** method to bypass filtering and get all the rows in the recordset. After you run the macro, note that the first set of data shown (corresponding to the first data row) contains the headings for all the data columns in the data recordset.




```vb
Public Sub GetRowData_Example() 
 
     
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intCount As Integer 
    Dim lngRowIDs() As Long 
    Dim lngRow As Long 
    Dim lngColumn As Long 
    Dim varRowData As Variant 
 
    'Get the count of all data recordsets in the current document. 
    intCount = ThisDocument.DataRecordsets.Count 
 
    Set vsoDataRecordset = ThisDocument.DataRecordsets(intCount) 
 
    'Get the row IDs of all the rows in the data recordset 
    lngRowIDs = vsoDataRecordset.GetDataRowIDs("") 
 
    'Iterate through all the records in the data recordset. 
    For lngRow = LBound(lngRowIDs) To UBound(lngRowIDs) + 1 
        varRowData = vsoDataRecordset.GetRowData(lngRow) 
 
        'Print a separator between rows 
        Debug.Print "------------------------------" 
 
       'Print the data stored in each column of a particular data row. 
        For lngColumn = LBound(varRowData) To UBound(varRowData) 
            Debug.Print varRowData(lngColumn) 
        Next lngColumn 
    Next lngRow 
 
End Sub
```


