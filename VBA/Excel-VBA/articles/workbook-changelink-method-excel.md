---
title: Workbook.ChangeLink Method (Excel)
keywords: vbaxl10.chm199083
f1_keywords:
- vbaxl10.chm199083
ms.prod: excel
api_name:
- Excel.Workbook.ChangeLink
ms.assetid: 9b2c0b82-73ff-3bdb-63df-82c0708cb703
ms.date: 06/08/2017
---


# Workbook.ChangeLink Method (Excel)

Changes a link from one document to another.


## Syntax

 _expression_ . **ChangeLink**( **_Name_** , **_NewName_** , **_Type_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the Microsoft Excel or DDE/OLE link to be changed, as it was returned from the  **[LinkSources](workbook-linksources-method-excel.md)** method.|
| _NewName_|Required| **String**|The new name of the link.|
| _Type_|Optional| **[XlLinkType](xllinktype-enumeration-excel.md)**|The link type.|

## Example

This example changes a Microsoft Excel link.




 **Note**  This example assumes at least one formula exists in the active workbook that links to another Excel source.




```vb
ActiveWorkbook.ChangeLink "c:\excel\book1.xls", _ 
 "c:\excel\book2.xls", xlExcelLinks
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

