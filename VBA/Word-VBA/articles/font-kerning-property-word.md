---
title: Font.Kerning Property (Word)
keywords: vbawd10.chm156369045
f1_keywords:
- vbawd10.chm156369045
ms.prod: word
api_name:
- Word.Font.Kerning
ms.assetid: 1fddf3d7-6750-dcac-2da6-f9da795a8d64
ms.date: 06/08/2017
---


# Font.Kerning Property (Word)

Returns or sets the minimum font size for which Microsoft Word will adjust kerning automatically. Read/write  **Single** .


## Syntax

 _expression_ . **Kerning**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Example

This example sets the minimum font size for automatic kerning to 12 points or larger in the active document.


```vb
ActiveDocument.Content.Font.Kerning = 12
```

This example displays the minimum font size for which Word will automatically adjust kerning in the selected text.




```vb
If Selection.Type = wdSelectionNormal Then 
 MsgBox Selection.Font.Kerning 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

