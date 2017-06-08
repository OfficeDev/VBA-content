---
title: Options.AutoHyphenate Property (Publisher)
keywords: vbapb10.chm1048580
f1_keywords:
- vbapb10.chm1048580
ms.prod: publisher
api_name:
- Publisher.Options.AutoHyphenate
ms.assetid: 821d0540-80ec-9f9d-777e-4d2596baf7d7
ms.date: 06/08/2017
---


# Options.AutoHyphenate Property (Publisher)

 **True** (default) for Microsoft Publisher to automatically hyphenate text in text frames. Read/write **Boolean**.


## Syntax

 _expression_. **AutoHyphenate**

 _expression_A variable that represents an  **Options** object.


### Return Value

Boolean


## Example

This example turns on automatic hyphenation for Publisher and sets the amount of space from the right margin to use when hyphenating words to one inch (72 points).


```vb
Sub SetHyphenationZone() 
 With Options 
 .AutoHyphenate = True 
 .HyphenationZone = 72 
 End With 
End Sub
```


