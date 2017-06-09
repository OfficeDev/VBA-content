---
title: Options.ReplaceSelection Property (Word)
keywords: vbawd10.chm162988099
f1_keywords:
- vbawd10.chm162988099
ms.prod: word
api_name:
- Word.Options.ReplaceSelection
ms.assetid: d1bef8ec-02e0-5f69-13af-0fdd758b3f0c
ms.date: 06/08/2017
---


# Options.ReplaceSelection Property (Word)

 **True** if the result of typing or pasting replaces the selection. Read/write **Boolean** .


## Syntax

 _expression_ . **ReplaceSelection**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

 **False** if the result of typing or pasting is added before the selection, leaving the selection intact.


## Example

This example sets Microsoft Word to add the result of typing or pasting before the selection, leaving the selection intact.


```vb
Options.ReplaceSelection = False
```

This example returns the status of the  **Typing replaces selection** option on the **Edit** tab in the **Options** dialog box ( **Tools** menu).




```
temp = Options.ReplaceSelection
```


## See also


#### Concepts


[Options Object](options-object-word.md)

