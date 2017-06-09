---
title: Workbook.SetLinkOnData Method (Excel)
keywords: vbaxl10.chm199151
f1_keywords:
- vbaxl10.chm199151
ms.prod: excel
api_name:
- Excel.Workbook.SetLinkOnData
ms.assetid: b500a579-6e4c-5712-05cf-27c6393b3bcd
ms.date: 06/08/2017
---


# Workbook.SetLinkOnData Method (Excel)

Sets the name of a procedure that runs whenever a DDE link is updated.


## Syntax

 _expression_ . **SetLinkOnData**( **_Name_** , **_Procedure_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the DDE/OLE link, as returned from the  **[LinkSources](workbook-linksources-method-excel.md)** method.|
| _Procedure_|Optional| **Variant**|The name of the procedure to be run when the link is updated. This can be either a Microsoft Excel 4.0 macro or a Visual Basic procedure. Set this argument to an empty string ("") to indicate that no procedure should run when the link is updated.|

## Example

This example sets the name of the procedure that runs whenever the DDE link is updated.


```vb
ActiveWorkbook.SetLinkOnData _ 
 "WinWord|'C:\MSGFILE.DOC'!DDE_LINK1", _ 
 "my_Link_Update_Macro"
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

