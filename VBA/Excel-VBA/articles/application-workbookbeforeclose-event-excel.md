---
title: Application.WorkbookBeforeClose Event (Excel)
keywords: vbaxl10.chm504084
f1_keywords:
- vbaxl10.chm504084
ms.prod: excel
api_name:
- Excel.Application.WorkbookBeforeClose
ms.assetid: 9c3618ea-0e5e-e4fe-20af-279826bfa7c3
ms.date: 06/08/2017
---


# Application.WorkbookBeforeClose Event (Excel)

Occurs immediately before any open workbook closes.


## Syntax

 _expression_ . **WorkbookBeforeClose**( **_Wb_** , **_Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The workbook that's being closed|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the workbook doesn't close when the procedure is finished.|

### Return Value

Nothing


## Example

This example prompts the user for a yes or no response before closing any workbook. For more information about how to use event procedures with the  **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/0063feba-47fd-29be-d2d5-8fcf47e70cbc%28Office.15%29.aspx).


```vb
Private Sub App_WorkbookBeforeClose(ByVal Wb as Workbook, _ 
 Cancel as Boolean) 
 a = MsgBox("Do you really want to close the workbook?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

