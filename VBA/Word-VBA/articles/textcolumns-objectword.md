---
title: TextColumns Object (Word)
keywords: vbawd10.chm2419
f1_keywords:
- vbawd10.chm2419
ms.prod: word
ms.assetid: 00b62c93-db7d-00b9-cc84-9a21e427d0cd
ms.date: 06/08/2017
---


# TextColumns Object (Word)

A collection of  **TextColumn** objects that represent all the columns of text in a document or a section of a document.


## Remarks

Use the  **TextColumns** property to return the **TextColumns** collection. The following example formats the columns in the first section in the active document to be evenly spaced, with a line between the columns.


```vb
With ActiveDocument.Sections(1).PageSetup.TextColumns 
 .EvenlySpaced = True 
 .LineBetween = True 
End With
```

Use the  **Add** method to add a column to the collection of columns. By default, there is one text column in the **TextColumns** collection. The following example adds a 2.5-inch-wide column to the active document.




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


