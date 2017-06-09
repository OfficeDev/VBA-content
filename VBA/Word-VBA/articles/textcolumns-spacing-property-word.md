---
title: TextColumns.Spacing Property (Word)
keywords: vbawd10.chm158531687
f1_keywords:
- vbawd10.chm158531687
ms.prod: word
api_name:
- Word.TextColumns.Spacing
ms.assetid: af171eb4-fa49-370c-6a8f-bf95abd57c31
ms.date: 06/08/2017
---


# TextColumns.Spacing Property (Word)

Returns or sets the spacing (in points) between between columns. Read/write  **Single** .


## Syntax

 _expression_ . **Spacing**

 _expression_ Required. A variable that represents a **[TextColumns](textcolumns-objectword.md)** collection.


## Remarks

After this property has been set for a  **TextColumns** object, the **EvenlySpaced** property is set to **True** . To return or set the spacing for a single text column when **EvenlySpaced** is **False** , use the **SpaceAfter** property of the **TextColumn** object.


## Example

This example formats the active document to display text in two columns with 0.5 inch (36 points) spacing between the columns.


```vb
With ActiveDocument.PageSetup.TextColumns 
 .SetCount NumColumns:=2 
 .LineBetween = False 
 .EvenlySpaced = True 
 .Spacing = 36 
End With
```


## See also


#### Concepts


[TextColumns Collection Object](textcolumns-objectword.md)

