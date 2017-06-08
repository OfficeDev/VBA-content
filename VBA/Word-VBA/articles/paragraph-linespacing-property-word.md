---
title: Paragraph.LineSpacing Property (Word)
keywords: vbawd10.chm156696685
f1_keywords:
- vbawd10.chm156696685
ms.prod: word
api_name:
- Word.Paragraph.LineSpacing
ms.assetid: f4ccfe57-4be8-1cdf-3140-45da603fc5ba
ms.date: 06/08/2017
---


# Paragraph.LineSpacing Property (Word)

Returns or sets the line spacing (in points) for the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **LineSpacing**

 _expression_ An expression that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

Use the  **[LinesToPoints](global-linestopoints-method-word.md)** method to convert a number of lines to the corresponding value in points. For example, `LinesToPoints(2)` returns the value 24.

The  **LineSpacing** property can be set after the **[LineSpacingRule](paragraph-linespacingrule-property-word.md)** property has been set to:


-  **wdLineSpaceAtLeast** the line spacing can be greater than or equal to, but never less than, the specified **LineSpacing** value.
    
-  **wdLineSpaceExactly** the line spacing never changes from the specified **LineSpacing** value, even if a larger font is used within the paragraph.
    
-  **wdLineSpaceMultiple** a **LineSpacing** property value must be specified, in points.
    

## Example

This example sets the line spacing for the first paragraph in the active document to always be at least 12 points.


```vb
With ActiveDocument.Paragraphs(1) 
 .LineSpacingRule = wdLineSpaceAtLeast 
 .LineSpacing = 12 
End With
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

