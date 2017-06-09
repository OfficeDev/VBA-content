---
title: Worksheet.Comments Property (Excel)
keywords: vbaxl10.chm175139
f1_keywords:
- vbaxl10.chm175139
ms.prod: excel
api_name:
- Excel.Worksheet.Comments
ms.assetid: c2ad8ea7-0fa3-7cde-e3f2-49bbdb81d261
ms.date: 06/08/2017
---


# Worksheet.Comments Property (Excel)

Returns a  **[Comments](comments-object-excel.md)** collection that represents all the comments for the specified worksheet. Read-only.


## Syntax

 _expression_ . **Comments**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example deletes all comments added by Jean Selva on the active sheet.


```vb
For Each c in ActiveSheet.Comments 
 If c.Author = "Jean Selva" Then c.Delete 
Next
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

