---
title: Options.SaveNormalPrompt Property (Word)
keywords: vbawd10.chm162988076
f1_keywords:
- vbawd10.chm162988076
ms.prod: word
api_name:
- Word.Options.SaveNormalPrompt
ms.assetid: bc58327f-d35e-70ae-ae53-0c312d3bbc0b
ms.date: 06/08/2017
---


# Options.SaveNormalPrompt Property (Word)

 **True** if Microsoft Word prompts the user for confirmation to save changes to the Normal template before it closes. Read/write **Boolean** .


## Syntax

 _expression_ . **SaveNormalPrompt**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

 **False** if Word automatically saves changes to the Normal template before it closes.


## Example

This example sets Word to save the Normal template automatically before closing, and then it quits.


```vb
Options.SaveNormalPrompt = False 
Application.Quit
```

This example returns the current status of the  **Prompt to save Normal template** option on the **Save** tab in the **Options** dialog box ( **Tools** menu).




```
temp = Options.SaveNormalPrompt
```


## See also


#### Concepts


[Options Object](options-object-word.md)

