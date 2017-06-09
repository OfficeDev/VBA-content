---
title: Options.PrintComments Property (Word)
keywords: vbawd10.chm162988065
f1_keywords:
- vbawd10.chm162988065
ms.prod: word
api_name:
- Word.Options.PrintComments
ms.assetid: 8c65566a-a6e8-5c38-9ef5-23591086bb68
ms.date: 06/08/2017
---


# Options.PrintComments Property (Word)

 **True** if Microsoft Word prints comments, starting on a new page at the end of the document. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintComments**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

Setting the  **PrintComments** property to **True** automatically sets the **[PrintHiddenText](options-printhiddentext-property-word.md)** property to **True** . However, setting the **PrintComments** property to **False** has no effect on the setting of the **PrintHiddenText** property.


## Example

This example sets Word to print comments and then prints the active document.


```vb
Options.PrintComments = True 
ActiveDocument.PrintOut
```


## See also


#### Concepts


[Options Object](options-object-word.md)

