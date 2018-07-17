---
title: DataRecordset.GetPrimaryKey Method (Visio)
keywords: vis_sdr.chm16460290
f1_keywords:
- vis_sdr.chm16460290
ms.prod: visio
api_name:
- Visio.DataRecordset.GetPrimaryKey
ms.assetid: 4f056424-4668-7859-5ed1-bd28a051ddc0
ms.date: 06/08/2017
---


# DataRecordset.GetPrimaryKey Method (Visio)

Gets the primary key setting and the name of the primary key column or columns for the data recordset.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetPrimaryKey**( **_PrimaryKeySettings_** , **_PrimaryKey()_** )

 _expression_ An expression that returns a **DataRecordset** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PrimaryKeySettings_|Required| **VisPrimaryKeySettings**|Out parameter. The primary key setting for the data recordset. See Remarks for possible values.|
| _PrimaryKey()_|Required| **String**|Out parameter. An array of  **String** variables.|

### Return Value

Nothing


## Remarks

You can use the  **GetPrimaryKey** method to determine the existing primary key setting for a data recordset and the name of the primary key column or columns if a primary key has been specified. The method returns the primary key setting for the data recordset in the PrimaryKeySettings out parameter, as a value from the **VisPrimaryKeySettings** enumeration. The default (when no primary key has been specified) is **visKeyRowOrder** , which means that Microsoft Visio identifies data recordset rows by row order.

A data recordset for which a primary key has been specified can have single or composite primary key columns. A single-column primary key bases row identification on the values in a single column. A composite primary key uses two or more columns to identify a row uniquely. Possible values for PrimaryKeySettings are shown in this table.



|**Constant**|**Value **|**Description**|
|:-----|:-----|:-----|
| **visKeyRowOrder**|1|Use row order as the primary key.|
| **visKeySingle**|2|Use a single column as the primary key column.|
| **visKeyComposite**|3|Use multiple columns as primary key columns.|
For the PrimaryKey() out parameter, pass a dimensionless array of strings. If the primary key setting returned is  **visKeySingle** or **visKeyComposite** , the method also returns an array of primary key column name strings in the PrimaryKey() out parameter. If the primary key setting is **visKeyRowOrder** , the default, the method returns an empty array.

You can use the  **[DataRecordset.SetPrimaryKey](datarecordset-setprimarykey-method-visio.md)** method to specify the primary key setting for the data recordset as well as the name of the column or columns that you want to set as the primary key column or columns. When you set primary keys, make sure that the column or columns you pick to be primary key columns contain unique values (or value sets) for each row.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how you can use the  **GetPrimaryKey** method to determine the primary key setting for a data recordset as well as the name of the first primary key column. The macro finds the most recently created data recordset associated with the document and, if a primary key has been specified, prints in the **Immediate** window the value of the primary key setting and the name of the first primary key column for the data recordset. If no primary key exists, it prints the primary key setting and the statement "No primary key."

Before running this macro, create at least one data recordset in the current document and, if you want, specify a primary key by using the  **SetPrimaryKey** method.




```vb
Public Sub GetPrimaryKey_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intCount As Integer 
    Dim astrPrimaryKeyColumns() As String 
    Dim vsoKeySettings As VisPrimaryKeySettings 
 
 
    intCount = ThisDocument.DataRecordsets.Count 
    Set vsoDataRecordset = ThisDocument.DataRecordsets(intCount) 
    vsoDataRecordset.GetPrimaryKey vsoKeySettings, astrPrimaryKeyColumns 
 
    If vsoKeySettings = visKeyRowOrder Then 
        Debug.Print vsoKeySettings, "No primary key" 
    Else 
        Debug.Print vsoKeySettings, astrPrimaryKeyColumns(0) 
    End If 
    
End Sub
```


