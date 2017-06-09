---
title: PivotCache.SaveAsODC Method (Excel)
keywords: vbaxl10.chm227106
f1_keywords:
- vbaxl10.chm227106
ms.prod: excel
api_name:
- Excel.PivotCache.SaveAsODC
ms.assetid: d7b553a5-70b1-41e7-9e35-088c23357570
ms.date: 06/08/2017
---


# PivotCache.SaveAsODC Method (Excel)

Saves the PivotTable cache source as an Microsoft Office Data Connection file.


## Syntax

 _expression_ . **SaveAsODC**( **_ODCFileName_** , **_Description_** , **_Keywords_** )

 _expression_ A variable that represents a **PivotCache** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ODCFileName_|Required| **String**|Location to save the file.|
| _Description_|Optional| **Variant**|Description that will be saved in the file.|
| _Keywords_|Optional| **Variant**|Space-separated keywords that can be used to search for this file.|

## Example

The following example saves the cache source as an ODC file titled "ODCFile". This example assumes a PivotTable cache exists on the active worksheet.


```vb
Sub UseSaveAsODC() 
 
 Application.ActiveWorkbook.PivotCaches.Item(1).SaveAsODC ("ODCFile") 
 
End Sub
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

