---
title: DataRecordset.GetDataRowIDs Method (Visio)
keywords: vis_sdr.chm16460310
f1_keywords:
- vis_sdr.chm16460310
ms.prod: visio
api_name:
- Visio.DataRecordset.GetDataRowIDs
ms.assetid: d76874eb-c25b-df65-5d00-64de288d086e
ms.date: 06/08/2017
---


# DataRecordset.GetDataRowIDs Method (Visio)

Gets an array of the IDs of all the rows in the data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetDataRowIDs**( **_CriteriaString_** )

 _expression_ An expression that returns a **DataRecordset** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CriteriaString_|Required| **String**|A string that conforms to the guidelines specified in the Microsoft ActiveX Data Object (ADO) API for setting the  **ADO.Filter** property|

### Return Value

Long()


## Remarks

You can use the  **GetDataRowIDs** method to get an array of the IDs of all the rows in a data recordset, where each row represents a single data record. The **GetDataRowIDs** method takes as its parameter a criteria string, which is a string that conforms to the guidelines specified in the ADO API for setting the **ADO.Filter** property. By specifying appropriate criteria and using AND and OR operators to separate clauses, you can filter the information in the data recordset to return only certain data recordset rows selectively. To apply no filter (that is, to get all the rows), pass an empty string ("").

For more information about criteria strings, see [Filter Property](http://msdn.microsoft.com/en-us/library/ms676691%28VS.85%29.aspx) in the ADO 2.8 API Reference.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how you can use the  **GetDataRowIDs** and **[GetRowData](datarecordset-getrowdata-method-visio.md)** methods to return the row ID of each row and then get the data stored in each column in every row of the specified data recordset. The macro gets the count of all data recordsets associated with the current document and gets row data for the one most recently added. It iterates through all the rows in the data recordset and then, for each row, iterates through all the columns in that row. The code displays the information returned in the **Immediate** window.

Before running this macro, create at least one data recordset in the current document.

Note that the macro passes an empty string to the  **GetDataRowIDs** method to bypass filtering and get all the rows in the recordset. After you run the macro, note that the first set of data shown (corresponding to the first data row) contains the headings for all the data columns in the data recordset.




```vb
Public Sub GetDataRowIDs_Example() 
 
     
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


