---
title: Options.PasteSmartCutPaste Property (Word)
keywords: vbawd10.chm162988470
f1_keywords:
- vbawd10.chm162988470
ms.prod: word
api_name:
- Word.Options.PasteSmartCutPaste
ms.assetid: d25143d6-2c83-ce37-3f8e-3177af0eccdd
ms.date: 06/08/2017
---


# Options.PasteSmartCutPaste Property (Word)

 **True** if Microsoft Word intelligently pastes selections into a document. Read/write **Boolean** .


## Syntax

 _expression_ . **PasteSmartCutPaste**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets Word to enable intelligent selection pasting if the option has been disabled.


```vb
Sub EnableSmartCutPaste() 
 If Options.PasteSmartCutPaste = False Then 
 Options.PasteSmartCutPaste = True 
 End If 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

