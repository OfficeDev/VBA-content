---
title: Options.AutoWordSelection Property (Word)
keywords: vbawd10.chm162988101
f1_keywords:
- vbawd10.chm162988101
ms.prod: word
api_name:
- Word.Options.AutoWordSelection
ms.assetid: 44b3a688-b5ef-6145-de33-00f0cf77409d
ms.date: 06/08/2017
---


# Options.AutoWordSelection Property (Word)

 **True** if dragging selects one word at a time instead of one character at a time. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoWordSelection**

 _expression_ A variable that represents an **[Options](options-object-word.md)** object.


## Example

This example sets Word to select individual characters instead of entire words when you select by dragging.


```vb
Options.AutoWordSelection = False
```

This example returns the status of the  **When selecting, automatically select entire word** option on the **Edit** tab in the **Options** dialog box.




```vb
Dim blnAutoSelect as Boolean 
 
blnAutoSelect = Options.AutoWordSelection
```


## See also


#### Concepts


[Options Object](options-object-word.md)

