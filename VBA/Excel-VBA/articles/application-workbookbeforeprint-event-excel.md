---
title: Application.WorkbookBeforePrint Event (Excel)
keywords: vbaxl10.chm504086
f1_keywords:
- vbaxl10.chm504086
ms.prod: excel
api_name:
- Excel.Application.WorkbookBeforePrint
ms.assetid: 27cb5f84-fda3-dc89-6e12-0c31ed16f47c
ms.date: 06/08/2017
---


# Application.WorkbookBeforePrint Event (Excel)

Occurs before any open workbook is printed.


## Syntax

 _expression_ . **WorkbookBeforePrint**( **_Wb_** , **_Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The workbook.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the workbook isn't printed when the procedure is finished.|

### Return Value

Nothing


## Example

This example recalculates all worksheets in the workbook before printing anything.


```vb
Private Sub App_WorkbookBeforePrint(ByVal Wb As Workbook, _ 
 Cancel As Boolean) 
 For Each wk in Wb.Worksheets 
 wk.Calculate 
 Next 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

