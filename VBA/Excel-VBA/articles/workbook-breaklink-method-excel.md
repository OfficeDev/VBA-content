---
title: Workbook.BreakLink Method (Excel)
keywords: vbaxl10.chm199198
f1_keywords:
- vbaxl10.chm199198
ms.prod: excel
api_name:
- Excel.Workbook.BreakLink
ms.assetid: 1e9d70c1-908e-92eb-26b8-d6ac753cc9c2
ms.date: 06/08/2017
---


# Workbook.BreakLink Method (Excel)

Converts formulas linked to other Microsoft Excel sources or OLE sources to values.


## Syntax

 _expression_ . **BreakLink**( **_Name_** , **_Type_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the link.|
| _Type_|Required| **[XlLinkType](xllinktype-enumeration-excel.md)**|The type of link.|

## Example

In this example, Microsoft Excel converts the first link (an Excel link type) in the active workbook. 




 **Note**  This example assumes at least one formula exists in the active workbook that links to another Excel source.




```vb
Sub UseBreakLink() 
 
 Dim astrLinks As Variant 
 
 ' Define variable as an Excel link type. 
 astrLinks = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks) 
 
 ' Break the first link in the active workbook. 
 ActiveWorkbook.BreakLink _ 
 Name:=astrLinks(1), _ 
 Type:=xlLinkTypeExcelLinks 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

