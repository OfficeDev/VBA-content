---
title: TextColumn Object (Word)
ms.prod: word
api_name:
- Word.TextColumn
ms.assetid: 660614a8-ad5b-dae4-887e-0f75e1172c10
ms.date: 06/08/2017
---


# TextColumn Object (Word)

Represents a single text column. The  **TextColumn** object is a member of the **[TextColumns](textcolumns-objectword.md)** collection. The **TextColumns** collection includes all the columns in a document or section of a document.


## Remarks

Use  **TextColumns** (Index), where Index is the index number, to return a single **TextColumn** object. The index number represents the position of the column in the **TextColumns** collection (counting from left to right).

The following example sets the space after the first text column in the active document to 0.5 inch.




```vb
ActiveDocument.PageSetup.TextColumns(1).SpaceAfter = _ 
 InchesToPoints(0.5)
```

Use the  **Add** method to add a column to the collection of columns. By default, there is one text column in the **TextColumns** collection. The following example adds a 2.5-inch-widecolumn to the active document.




```vb
ActiveDocument.PageSetup.TextColumns.Add _ 
 Width:=InchesToPoints(2.5), _ 
 Spacing:=InchesToPoints(0.5), EvenlySpaced:=False
```

Use the  **SetCount** method to arrange text into columns. The following example arranges the text in the active document into three columns.




```vb
ActiveDocument.PageSetup.TextColumns.SetCount NumColumns:=3
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


