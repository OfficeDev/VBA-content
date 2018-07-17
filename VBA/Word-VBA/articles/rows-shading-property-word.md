---
title: Rows.Shading Property (Word)
keywords: vbawd10.chm155975782
f1_keywords:
- vbawd10.chm155975782
ms.prod: word
api_name:
- Word.Rows.Shading
ms.assetid: 79c5240c-2845-e038-49cb-8a9b1f8f2a71
ms.date: 06/08/2017
---


# Rows.Shading Property (Word)

Returns a  **[Shading](shading-object-word.md)** object that refers to the shading formatting for the specified object.


## Syntax

 _expression_ . **Shading**

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


## Example

This example applies yellow shading to the first paragraph in the selection.


```vb
With Selection.Paragraphs(1).Shading 
 .Texture = wdTexture12Pt5Percent 
 .BackgroundPatternColorIndex = wdYellow 
 .ForegroundPatternColorIndex = wdBlack 
End With
```

This example applies horizontal line texture to the first row in table one.




```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Rows(1).Shading 
 .Texture = wdTextureHorizontal 
 End With 
End If
```

This example applies 10 percent shading to the first word in the active document.




```vb
ActiveDocument.Words(1).Shading.Texture = wdTexture10Percent
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

