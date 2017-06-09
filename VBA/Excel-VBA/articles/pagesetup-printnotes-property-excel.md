---
title: PageSetup.PrintNotes Property (Excel)
keywords: vbaxl10.chm473095
f1_keywords:
- vbaxl10.chm473095
ms.prod: excel
api_name:
- Excel.PageSetup.PrintNotes
ms.assetid: 6609fe58-6015-9ae2-4cc0-107e29cd7b9d
ms.date: 06/08/2017
---


# PageSetup.PrintNotes Property (Excel)

 **True** if cell notes are printed as end notes with the sheet. Applies only to worksheets. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintNotes**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

Use the  **PrintComments** property to print comments as text boxes or end notes.


## Example

This example turns off the printing of notes.


```vb
Worksheets("Sheet1").PageSetup.PrintNotes = False
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

