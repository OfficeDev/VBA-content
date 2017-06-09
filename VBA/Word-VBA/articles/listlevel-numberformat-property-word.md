---
title: ListLevel.NumberFormat Property (Word)
keywords: vbawd10.chm160235522
f1_keywords:
- vbawd10.chm160235522
ms.prod: word
api_name:
- Word.ListLevel.NumberFormat
ms.assetid: 45305290-e1ca-cd5b-98bd-e60fad989ec5
ms.date: 06/08/2017
---


# ListLevel.NumberFormat Property (Word)

Returns or sets the number format for the specified list level. Read/write  **String** .


## Syntax

 _expression_ . **NumberFormat**

 _expression_ An expression that returns a **[ListLevel](listlevel-object-word.md)** object.


## Remarks

The percent sign (%) followed by any number from 1 through 9 represents the number style from the respective list level. For example, if you wanted the format for the first level to be "Article I," "Article II," and so on, the string for the  **NumberFormat** property would be "Article %1" and the **[NumberStyle](listlevel-numberstyle-property-word.md)** property would be set to **wdListNumberStyleUpperCaseRoman** .

If the  **NumberStyle** property is set to **wdListNumberStyleBullet** , the string for the **NumberFormat** property can only contain one character.


## Example

This example creates a list template that indents each level and formats the level with an Arabic numeral and a period. The new list template is then applied to the selection.


```vb
Set LT = ActiveDocument.ListTemplates.Add(OutlineNumbered:=True) 
For x = 1 To 9 
 With LT.ListLevels(x) 
 .NumberStyle = wdListNumberStyleArabic 
 .NumberPosition = InchesToPoints(0.25 * (x - 1)) 
 .TextPosition = InchesToPoints(0.25 * x) 
 .NumberFormat = "%" &; x &; "." 
 End With 
Next x 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=LT
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

