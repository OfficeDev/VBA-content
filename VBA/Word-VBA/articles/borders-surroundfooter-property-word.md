---
title: Borders.SurroundFooter Property (Word)
keywords: vbawd10.chm154927129
f1_keywords:
- vbawd10.chm154927129
ms.prod: word
api_name:
- Word.Borders.SurroundFooter
ms.assetid: 890b0ba3-6815-6836-591d-f73d90758c4b
ms.date: 06/08/2017
---


# Borders.SurroundFooter Property (Word)

 **True** if a page border encompasses the document footer. Read/write **Boolean** .


## Syntax

 _expression_ . **SurroundFooter**

 _expression_ An expression that returns a **[Borders](borders-object-word.md)** collection object.


## Example

This example formats the page border in section one of the active document so that it encompasses the header and footer on each page in the section.


```vb
With ActiveDocument.Sections(1).Borders 
 .SurroundFooter = True 
 .SurroundHeader = True 
End With
```

This example adds a graphical page border around each page in section one. The page border doesn't encompass the header and footer areas.




```vb
With ActiveDocument.Sections(1) 
 .Borders.SurroundFooter = False 
 .Borders.SurroundHeader = False 
 For Each aBord In .Borders 
 aBord.ArtStyle = wdArtPeople 
 aBord.ArtWidth = 15 
 Next aBord 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

