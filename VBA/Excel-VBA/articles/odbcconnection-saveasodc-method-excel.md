---
title: ODBCConnection.SaveAsODC Method (Excel)
keywords: vbaxl10.chm796085
f1_keywords:
- vbaxl10.chm796085
ms.prod: excel
api_name:
- Excel.ODBCConnection.SaveAsODC
ms.assetid: a499de7c-ee4a-22d2-ff35-33489fcf4fe1
ms.date: 06/08/2017
---


# ODBCConnection.SaveAsODC Method (Excel)

Saves the ODBC connection as a Microsoft Office Data Connection file.


## Syntax

 _expression_ . **SaveAsODC**( **_ODCFileName_** , **_Description_** , **_Keywords_** )

 _expression_ A variable that represents an **ODBCConnection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ODCFileName_|Required| **String**|Location to save the file.|
| _Description_|Optional| **Variant**|Description that will be saved in the file.|
| _Keywords_|Optional| **Variant**|Space-separated keywords that can be used to search for this file.|

### Return Value

Nothing


## Example

The following example saves the connection as an ODC file titled "ODCFile". This example assumes an ODBC connection exists on the active worksheet.


```vb
Sub UseSaveAsODC() 
 
 Application.ActiveWorkbook.ODBCConnection.SaveAsODC ("ODCFile") 
 
End Sub
```


## See also


#### Concepts


[ODBCConnection Object](odbcconnection-object-excel.md)

