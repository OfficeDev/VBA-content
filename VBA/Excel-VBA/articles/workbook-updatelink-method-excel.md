---
title: Workbook.UpdateLink Method (Excel)
keywords: vbaxl10.chm199160
f1_keywords:
- vbaxl10.chm199160
ms.prod: excel
api_name:
- Excel.Workbook.UpdateLink
ms.assetid: 2aef72cc-a820-3e91-1f46-50c739faf2bb
ms.date: 06/08/2017
---


# Workbook.UpdateLink Method (Excel)

Updates a Microsoft Excel, DDE, or OLE link (or links).


## Syntax

 _expression_ . **UpdateLink**( **_Name_** , **_Type_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional| **Variant**|The name of the Microsoft Excel or DDE/OLE link to be updated, as returned from the  **[LinkSources](workbook-linksources-method-excel.md)** method.|
| _Type_|Optional| **Variant**|One of the constants of  **[XlLinkType](xllinktype-enumeration-excel.md)** specifying the type of link.|

## Remarks

When the  **UpdateLink** method is called without any parameters, Excel defaults to updating all worksheet links.


## Example

This example updates all links in the active workbook.


```vb
ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

