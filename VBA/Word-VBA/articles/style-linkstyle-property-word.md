---
title: Style.LinkStyle Property (Word)
keywords: vbawd10.chm153878632
f1_keywords:
- vbawd10.chm153878632
ms.prod: word
api_name:
- Word.Style.LinkStyle
ms.assetid: 3a5c4f41-be1e-9da4-5f94-6d2db00616f5
ms.date: 06/08/2017
---


# Style.LinkStyle Property (Word)

Sets or returns a  **Variant** that represents a link between a paragraph and a character style. Read/write.


## Syntax

 _expression_ . **LinkStyle**

 _expression_ An expression that returns a **[Style](style-object-word.md)** object.


## Remarks

When a character style and a paragraph style are linked, the two styles take on the same character formatting.


## Example

This example creates and formats a new character style, and then it links the character style to the built-in heading style "Heading 1" so that the "Heading 1" style takes on the character formatting of the newly added style.


```vb
Sub LinkHeadStyle() 
 Dim styChar1 As Style 
 
 Set styChar1 = ActiveDocument.Styles.Add _ 
 (Name:="Heading 1 Characters", Type:=wdStyleTypeCharacter) 
 With styChar1 
 .Font.Name = "Verdana" 
 .Font.Bold = True 
 .Font.Shadow = True 
 With .Font.Borders(1) 
 .LineStyle = wdLineStyleDot 
 .LineWidth = wdLineWidth300pt 
 .Color = wdColorDarkRed 
 End With 
 End With 
 ActiveDocument.Styles("Heading 1").LinkStyle = _ 
 ActiveDocument.Styles("Heading 1 Characters") 
 
 With ActiveDocument.Content 
 .InsertParagraphAfter 
 .InsertAfter "New Linked Style" 
 .Select 
 End With 
 
 Selection.Collapse Direction:=wdCollapseEnd 
 Selection.Style = ActiveDocument.Styles("Heading 1") 
 
End Sub
```


## See also


#### Concepts


[Style Object](style-object-word.md)

