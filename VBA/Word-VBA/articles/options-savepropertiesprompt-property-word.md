---
title: Options.SavePropertiesPrompt Property (Word)
keywords: vbawd10.chm162988075
f1_keywords:
- vbawd10.chm162988075
ms.prod: word
api_name:
- Word.Options.SavePropertiesPrompt
ms.assetid: da2bbc7d-920d-2442-25d3-c6ee11316097
ms.date: 06/08/2017
---


# Options.SavePropertiesPrompt Property (Word)

 **True** if Microsoft Word prompts for document property information when saving a new document. Read/write **Boolean** .


## Syntax

 _expression_ . **SavePropertiesPrompt**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example causes Word to prompt for document property information when saving a new document.


```vb
Options.SavePropertiesPrompt = True
```

This example returns the current status of the  **Prompt for document properties** option on the **Save** tab in the **Options** dialog box ( **Tools** menu).




```
temp = Options.SavePropertiesPrompt
```


## See also


#### Concepts


[Options Object](options-object-word.md)

