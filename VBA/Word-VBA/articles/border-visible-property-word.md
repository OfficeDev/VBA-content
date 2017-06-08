---
title: Border.Visible Property (Word)
keywords: vbawd10.chm154861568
f1_keywords:
- vbawd10.chm154861568
ms.prod: word
api_name:
- Word.Border.Visible
ms.assetid: 7040aa03-17dc-073c-c9db-e4a7cc2e7ef9
ms.date: 06/08/2017
---


# Border.Visible Property (Word)

 **True** if the specified object is visible. Read/write **Boolean** .


## Syntax

 _expression_ . **Visible**

 _expression_ Required. A variable that represents a **[Border](border-object-word.md)** object.


## Remarks

For any object, some methods and properties may be unavailable if the  **Visible** property is **False** .


## Example

This example creates a table in the active document and removes the default borders from the table.


```vb
Set myTable = ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=12, NumColumns:=5) 
For Each aBorder In myTable.Borders 
 aBorder.Visible = False 
Next aBorder
```


## See also


#### Concepts


[Border Object](border-object-word.md)

