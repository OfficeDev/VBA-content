---
title: Options.PrintEvenPagesInAscendingOrder Property (Word)
keywords: vbawd10.chm162988363
f1_keywords:
- vbawd10.chm162988363
ms.prod: word
api_name:
- Word.Options.PrintEvenPagesInAscendingOrder
ms.assetid: 355f973c-d60f-5953-8b0d-0b8c5798dce1
ms.date: 06/08/2017
---


# Options.PrintEvenPagesInAscendingOrder Property (Word)

 **True** if Microsoft Word prints even pages in ascending order during manual duplex printing. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintEvenPagesInAscendingOrder**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

If the ManualDuplexPrint argument of the  **[PrintOut](application-printout-method-word.md)** method is **False** , this property is ignored.


## Example

This example sets Word to print odd pages in ascending order and even pages in descending order during manual duplex printing, and then it prints the active document.


```vb
Options.PrintOddPagesInAscendingOrder = True 
Options.PrintEvenPagesInAscendingOrder = False 
ActiveDocument.PrintOut ManualDuplexPrint:=True
```


## See also


#### Concepts


[Options Object](options-object-word.md)

