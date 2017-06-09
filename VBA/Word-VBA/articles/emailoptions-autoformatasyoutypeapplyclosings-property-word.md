---
title: EmailOptions.AutoFormatAsYouTypeApplyClosings Property (Word)
keywords: vbawd10.chm165347627
f1_keywords:
- vbawd10.chm165347627
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeApplyClosings
ms.assetid: b5be989e-09ff-455f-5d8a-638016512e3d
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeApplyClosings Property (Word)

 **True** for Microsoft Word to automatically apply the Closing style to letter closings as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyClosings**

 _expression_ Required. A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically apply the Closing style to letter closings as you type.


```vb
Sub AutoClosings() 
 Options.AutoFormatAsYouTypeApplyClosings = True 
End Sub
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

