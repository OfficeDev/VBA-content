---
title: Options.AutoFormatAsYouTypeInsertClosings Property (Word)
keywords: vbawd10.chm162988335
f1_keywords:
- vbawd10.chm162988335
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeInsertClosings
ms.assetid: 8e51f053-03df-84c3-cd08-d53281602646
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeInsertClosings Property (Word)

 **True** for Microsoft Word to automatically insert the corresponding memo closing when the user enters a memo heading. Read/write.


## Syntax

 _expression_ . **AutoFormatAsYouTypeInsertClosings**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically insert the corresponding memo closing when the user enters a memo heading.


```vb
Sub AutoInsertClosings() 
 Options.AutoFormatAsYouTypeInsertClosings = True 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

