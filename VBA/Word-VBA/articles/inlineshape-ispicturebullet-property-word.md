---
title: InlineShape.IsPictureBullet Property (Word)
keywords: vbawd10.chm162005124
f1_keywords:
- vbawd10.chm162005124
ms.prod: word
api_name:
- Word.InlineShape.IsPictureBullet
ms.assetid: c53c7269-b6ab-beaa-41d6-105033c077b9
ms.date: 06/08/2017
---


# InlineShape.IsPictureBullet Property (Word)

 **True** indicates that an **InlineShape** object is a picture bullet. Read-only **Boolean** .


## Syntax

 _expression_ . **IsPictureBullet**

 _expression_ An expression that returns a **[InlineShape](inlineshape-object-word.md)** object.


## Remarks

Although picture bullets are considered inline shapes, searching a document's  **InlineShapes** collection will not return picture bullets.


## Example

This example formats the selected list if the list if formatted with a picture bullet. If not, a message is displayed.


```vb
Sub IsSelectionAPictureBullet(shp As InlineShape) 
 On Error GoTo ErrorHandler 
 If shp.IsPictureBullet = True Then 
 shp.Width = InchesToPoints(0.5) 
 shp.Height = InchesToPoints(0.05) 
 End If 
 Exit Sub 
ErrorHandler: 
 MsgBox "The selection is not a list or " &; _ 
 "does not contain picture bullets." 
End Sub
```

Use the following code to call the routine above.




```vb
Sub CallPic() 
 Call IsSelectionAPictureBullet(shp:=Selection _ 
 .Range.ListFormat.ListPictureBullet) 
End Sub
```


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

