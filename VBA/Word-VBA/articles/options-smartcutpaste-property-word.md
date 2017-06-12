---
title: Options.SmartCutPaste Property (Word)
keywords: vbawd10.chm162988103
f1_keywords:
- vbawd10.chm162988103
ms.prod: word
api_name:
- Word.Options.SmartCutPaste
ms.assetid: 57e481b6-f3c4-8da4-2580-4abbbf21a95e
ms.date: 06/08/2017
---


# Options.SmartCutPaste Property (Word)

 **True** if Microsoft Word automatically adjusts the spacing between words and punctuation when cutting and pasting occurs. Read/write **Boolean** .


## Syntax

 _expression_ . **SmartCutPaste**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example sets Word to automatically adjust the spacing between words and punctuation when cutting and pasting occurs, and then it deletes and pastes some text in a newly created document. If the  **SmartCutPaste** property were set to **False** , the second and third words would run together.


```vb
Options.SmartCutPaste = True 
Set myDoc = Documents.Add 
With myDoc 
 .Content.InsertAfter("The brown quick fox") 
 .Words(2).Cut 
 .Characters(10).Paste 
End With
```

This example returns the status of the  **Smart cut and paste** option on the **Edit** tab in the **Options** dialog box ( **Tools** menu).




```
temp = Options.SmartCutPaste
```


## See also


#### Concepts


[Options Object](options-object-word.md)

