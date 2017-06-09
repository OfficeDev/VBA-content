---
title: Worksheet.PrintedCommentPages Property (Excel)
keywords: vbaxl10.chm175164
f1_keywords:
- vbaxl10.chm175164
ms.prod: excel
api_name:
- Excel.Worksheet.PrintedCommentPages
ms.assetid: 3ade9c86-c6b9-08fa-3bc6-a040dd1da36a
ms.date: 06/08/2017
---


# Worksheet.PrintedCommentPages Property (Excel)

Returns the number of comment pages that will be printed for the current worksheet. Read-only


## Syntax

 _expression_ . **PrintedCommentPages**

 _expression_ A variable that returns a **Worksheet** object.


### Return Value

Long


## Remarks

The  **PrintedCommentPages** property will only return a number greater than zero if the **Comments** setting on the **Sheet** tab of the **Page Setup** dialog box is set to **At end of sheet**. This property will return zero if the sheet is a  **Chart** sheet or a **MS Excel 5.0 Dialog** sheet.


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

