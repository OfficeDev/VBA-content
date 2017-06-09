---
title: Application.ThisWorkbook Property (Excel)
keywords: vbaxl10.chm183111
f1_keywords:
- vbaxl10.chm183111
ms.prod: excel
api_name:
- Excel.Application.ThisWorkbook
ms.assetid: 04b713dd-fd93-3cbc-f10b-05b9c3d107b1
ms.date: 06/08/2017
---


# Application.ThisWorkbook Property (Excel)

Returns a  **[Workbook](workbook-object-excel.md)** object that represents the workbook where the current macro code is running. Read-only.


## Syntax

 _expression_ . **ThisWorkbook**

 _expression_ A variable that represents an **Application** object.


## Remarks

Use this property to refer to the workbook that contains your macro code.  **ThisWorkbook** is the only way to refer to an add-in workbook from inside the add-in itself. The **ActiveWorkbook** property doesn't return the add-in workbook; it returns the workbook that's calling the add-in.

The  **Workbooks** property may fail, as the workbook name probably changed when you created the add-in. **ThisWorkbook** always returns the workbook in which the code is running.

For example, use code such as the following to activate a dialog sheet stored in your add-in workbook.

 `ThisWorkbook.DialogSheets(1).Show`

This property can be used only from inside Microsoft Excel. You cannot use it to access a workbook from any other application.


## Example

This example closes the workbook that contains the example code. Changes to the workbook, if any, aren't saved.


```vb
ThisWorkbook.Close SaveChanges:=False
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

