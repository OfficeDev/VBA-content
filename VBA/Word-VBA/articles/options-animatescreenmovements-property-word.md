---
title: Options.AnimateScreenMovements Property (Word)
keywords: vbawd10.chm162988106
f1_keywords:
- vbawd10.chm162988106
ms.prod: word
api_name:
- Word.Options.AnimateScreenMovements
ms.assetid: 8f4a7986-887e-8752-4d77-6db54db58da6
ms.date: 06/08/2017
---


# Options.AnimateScreenMovements Property (Word)

 **True** if Word animates mouse movements, uses animated cursors, and animates actions such as background saving and find and replace operations. Read/write **Boolean** .


## Syntax

 _expression_ . **AnimateScreenMovements**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets Word to animate movements on the screen.


```vb
Options.AnimateScreenMovements = True
```

This example returns the current status of the Provide feedback with animation option on the General tab in the Options dialog box (Tools menu).




```vb
Dim blnAnimation as Boolean blnAnimation = Options.AnimateScreenMovements
```


## See also


#### Concepts


[Options Object](options-object-word.md)

