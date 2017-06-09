---
title: EmailOptions.AutoFormatAsYouTypeInsertClosings Property (Word)
keywords: vbawd10.chm165347631
f1_keywords:
- vbawd10.chm165347631
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeInsertClosings
ms.assetid: f08ab03c-bcc1-0fd2-c752-5476ba641504
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeInsertClosings Property (Word)

 **True** for Microsoft Word to automatically insert the corresponding memo closing when the user enters a memo heading. Read/write.


## Syntax

 _expression_ . **AutoFormatAsYouTypeInsertClosings**

 _expression_ Required. A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically insert the corresponding memo closing when the user enters a memo heading.


```vb
Sub AutoInsertClosings() 
 Options.AutoFormatAsYouTypeInsertClosings = True 
End Sub
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

