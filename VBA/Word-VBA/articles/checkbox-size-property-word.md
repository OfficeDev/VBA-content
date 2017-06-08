---
title: CheckBox.Size Property (Word)
keywords: vbawd10.chm153485314
f1_keywords:
- vbawd10.chm153485314
ms.prod: word
api_name:
- Word.CheckBox.Size
ms.assetid: 1e7fe0d6-7dd9-c19b-a5b4-f60f99ee6bae
ms.date: 06/08/2017
---


# CheckBox.Size Property (Word)

Returns or sets the size of a check box, in points. Read/write  **Single** .


## Syntax

 _expression_ . **Size**

 _expression_ A variable that represents a **[CheckBox](checkbox-object-word.md)** object.


## Example

This example sets the size of the check box named "Check1" in the active document to 14 points and then sets the check box as selected.


```vb
With ActiveDocument.FormFields("Check1").CheckBox 
 .AutoSize = False 
 .Size = 14 
 .Value = True 
End With
```


## See also


#### Concepts


[CheckBox Object](checkbox-object-word.md)

