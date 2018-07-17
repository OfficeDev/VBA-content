---
title: Paragraphs.LineSpacing Property (Word)
keywords: vbawd10.chm156762221
f1_keywords:
- vbawd10.chm156762221
ms.prod: word
api_name:
- Word.Paragraphs.LineSpacing
ms.assetid: 3609a32b-3d28-eb9f-4eb9-68a69ed818a2
ms.date: 06/08/2017
---


# Paragraphs.LineSpacing Property (Word)

Returns or sets the line spacing (in points) for the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **LineSpacing**

 _expression_ An expression that represents a **[Paragraphs](paragraphs-object-word.md)** object.


## Remarks

Use the  **[LinesToPoints](global-linestopoints-method-word.md)** method to convert a number of lines to the corresponding value in points. For example, `LinesToPoints(2)` returns the value 24.

The  **LineSpacing** property can be set after the **[LineSpacingRule](paragraphs-linespacingrule-property-word.md)** property has been set to:


-  **wdLineSpaceAtLeast** the line spacing can be greater than or equal to, but never less than, the specified **LineSpacing** value.
    
-  **wdLineSpaceExactly** the line spacing never changes from the specified **LineSpacing** value, even if a larger font is used within the paragraph.
    
-  **wdLineSpaceMultiple** a **LineSpacing** property value must be specified, in points.
    

## Example

This example triple-spaces the lines in the selected paragraphs.


```vb
With Selection.Paragraphs 
 .LineSpacingRule = wdLineSpaceMultiple 
 .LineSpacing = LinesToPoints(3) 
End With 

```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

