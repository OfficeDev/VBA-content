---
title: TextColumn.Width Property (Word)
ms.prod: word
api_name:
- Word.TextColumn.Width
ms.assetid: 4050636e-0721-56b2-7a63-3f56906e3ca6
ms.date: 06/08/2017
---


# TextColumn.Width Property (Word)

Returns or sets the width, in points, of the specified text columns. Read/write  **Long** .


## Syntax

 _expression_ . **Width**

 _expression_ A variable that represents a **[TextColumn](textcolumn-object-word.md)** object.


## Example

This example formats the section that includes the selection as three columns. The  **For Each** loop is used to display the width of each column in the **TextColumns** collection.


```vb
Selection.PageSetup.TextColumns.SetCount NumColumns:=3 
For Each acol In Selection.PageSetup.TextColumns 
 MsgBox "Width= " &; PointsToInches(acol.Width) 
Next acol
```


## See also


#### Concepts


[TextColumn Object](textcolumn-object-word.md)

