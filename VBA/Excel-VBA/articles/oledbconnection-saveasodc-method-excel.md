---
title: OLEDBConnection.SaveAsODC Method (Excel)
keywords: vbaxl10.chm794089
f1_keywords:
- vbaxl10.chm794089
ms.prod: excel
api_name:
- Excel.OLEDBConnection.SaveAsODC
ms.assetid: da83acf3-c935-c36f-944e-35b46e54cabf
ms.date: 06/08/2017
---


# OLEDBConnection.SaveAsODC Method (Excel)

Saves the OLE DB connection as an Microsoft Office Data Connection file.


## Syntax

 _expression_ . **SaveAsODC**( **_ODCFileName_** , **_Description_** , **_Keywords_** )

 _expression_ A variable that represents an **OLEDBConnection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ODCFileName_|Required| **String**|Location to save the file.|
| _Description_|Optional| **Variant**|Description that will be saved in the file.|
| _Keywords_|Optional| **Variant**|Space-separated keywords that can be used to search for this file.|

## Example

The following example saves the connection as an ODC file titled "ODCFile". This example assumes an OLE DB connection exists on the active worksheet.


```vb
Sub UseSaveAsODC() 
 
 Application.ActiveWorkbook.OLEDBConnection.SaveAsODC ("ODCFile") 
 
End Sub
```


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

