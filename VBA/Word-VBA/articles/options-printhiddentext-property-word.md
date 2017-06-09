---
title: Options.PrintHiddenText Property (Word)
keywords: vbawd10.chm162988066
f1_keywords:
- vbawd10.chm162988066
ms.prod: word
api_name:
- Word.Options.PrintHiddenText
ms.assetid: 4f047b82-884e-5109-b931-838f3742094d
ms.date: 06/08/2017
---


# Options.PrintHiddenText Property (Word)

 **True** if hidden text is printed. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintHiddenText**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

Setting the  **PrintHiddenText** property to **False** automatically sets the **[PrintComments](options-printcomments-property-word.md)** property to **False** . However, setting the **PrintHiddenText** property to **True** has no effect on the setting of the **PrintComments** property.


## Example

This example sets Word to print hidden text, and then it prints the active document.


```vb
Options.PrintHiddenText = True 
ActiveDocument.PrintOut
```

This example returns the current status of the  **Hidden text** option on the **Print** tab in the **Options** dialog box.




```
temp = Options.PrintHiddenText
```


## See also


#### Concepts


[Options Object](options-object-word.md)

