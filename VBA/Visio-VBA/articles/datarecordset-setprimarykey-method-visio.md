---
title: DataRecordset.SetPrimaryKey Method (Visio)
keywords: vis_sdr.chm16460295
f1_keywords:
- vis_sdr.chm16460295
ms.prod: visio
api_name:
- Visio.DataRecordset.SetPrimaryKey
ms.assetid: 5ec125ff-b4a8-abcb-0d9d-140e97de6db2
ms.date: 06/08/2017
---


# DataRecordset.SetPrimaryKey Method (Visio)

Sets the primary key setting value and the name of the primary key column or columns for the data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **SetPrimaryKey**( **_PrimaryKeySettings_** , **_PrimaryKey()_** )

 _expression_ An expression that returns a **DataRecordset** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PrimaryKeySettings_|Required| **VisPrimaryKeySettings**|The primary key setting for the data recordset. See Remarks for possible values.|
| _PrimaryKey()_|Required| **[SAFE-ARRAY]**|An array of  **String** variables.|

### Return Value

Nothing


## Remarks

You can use the  **SetPrimaryKey** method to specify the primary key setting and the name of the primary key column or columns for the data recordset. You specify the primary key setting for the data recordset by passing a value from the **VisPrimaryKeySettings** enumeration for the PrimaryKeySettings parameter. The default (when you don't specify a primary key) is **visKeyRowOrder** , which means that Visio identifies data recordset rows by row order.

You can specify that the data recordset have either a single-column or a composite primary key. A single-column primary key bases row identification on the values in a single column. A composite primary key uses two or more columns to identify a row uniquely. Possible values for PrimaryKeySettings are shown in this table.



|**Constant**|**Value **|**Description**|
|:-----|:-----|:-----|
| **visKeyRowOrder**|1|Use row order as the primary key.|
| **visKeySingle**|2|Use a single column as the primary key column.|
| **visKeyComposite**|3|Use multiple columns as primary key columns.|
For the PrimaryKey() parameter, pass an array of one or more strings that represent the name of the column or columns you want to set as the primary key column(s). The value you pass for the PrimaryKeySettings parameter must be consistent with the number of array items. When you set primary keys, make sure that the column or columns you pick to be primary key columns contain unique values (or value sets) for each row.

You can use the  **[GetPrimaryKey](datarecordset-getprimarykey-method-visio.md)** method to determine the current primary key setting for the data recordset as well as the name of the column or columns, if any, that are currently set as the primary key column or columns.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how you can use the  **SetPrimaryKey** method to specify the primary key setting for a data recordset as well as the name of the primary key column. The macro finds the most recently created data recordset associated with the document, specifies the primary key setting ( **visKeySingle** , to indicate a single-column primary key), and sets the name of the primary key column.

Before running this macro, create at least one data recordset in the current document, and replace the variable  _columnName_ in the code with the name of the column in the data recordset that you want to specify as the primary key column.




```vb
Public Sub SetPrimaryKey_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intCount As Integer 
    Dim aPrimaryKeyColumns() As String 
     
    intCount = ThisDocument.DataRecordsets.Count 
    aPrimaryKeyColumns(0) = "columnName " 
    Set vsoDataRecordset = ThisDocument.DataRecordsets(intCount) 
    vsoDataRecordset.SetPrimaryKey visKeySingle, aPrimaryKeyColumns 
    
End Sub
```


