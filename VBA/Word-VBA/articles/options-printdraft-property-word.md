---
title: Options.PrintDraft Property (Word)
keywords: vbawd10.chm162988319
f1_keywords:
- vbawd10.chm162988319
ms.prod: word
api_name:
- Word.Options.PrintDraft
ms.assetid: 23be1e0a-784b-5b0f-107c-78e200e31159
ms.date: 06/08/2017
---


# Options.PrintDraft Property (Word)

 **True** if Microsoft Word prints using minimal formatting. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintDraft**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

Not all printers support draft printing.


## Example

This example sets Word to use draft printing and then prints the active document.


```vb
Options.PrintDraft = True 
ActiveDocument.PrintOut
```

This example returns the current status of the  **Draft output** option on the **Print** tab in the **Options** dialog box ( **Tools** menu).




```
temp = Options.PrintDraft
```


## See also


#### Concepts


[Options Object](options-object-word.md)

