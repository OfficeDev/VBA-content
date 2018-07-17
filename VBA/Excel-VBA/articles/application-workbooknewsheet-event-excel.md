---
title: Application.WorkbookNewSheet Event (Excel)
keywords: vbaxl10.chm504087
f1_keywords:
- vbaxl10.chm504087
ms.prod: excel
api_name:
- Excel.Application.WorkbookNewSheet
ms.assetid: 5190254f-b7f4-10e5-41f5-704b1466ff68
ms.date: 06/08/2017
---


# Application.WorkbookNewSheet Event (Excel)

Occurs when a new sheet is created in any open workbook.


## Syntax

 _expression_ . **WorkbookNewSheet**( **_Wb_** , **_Sh_** )

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The workbook.|
| _Sh_|Required| **Object**|The new sheet.|

### Return Value

Nothing


## Example

This example moves the new sheet to the end of the workbook.


```vb
Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, _ 
 ByVal Sh As Object) 
 Sh.Move After:=Wb.Sheets(Wb.Sheets.Count) 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

