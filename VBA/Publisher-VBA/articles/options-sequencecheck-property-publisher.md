---
title: Options.SequenceCheck Property (Publisher)
keywords: vbapb10.chm1048625
f1_keywords:
- vbapb10.chm1048625
ms.prod: publisher
api_name:
- Publisher.Options.SequenceCheck
ms.assetid: a2801af8-5c89-9256-80a6-d9dac17b6066
ms.date: 06/08/2017
---


# Options.SequenceCheck Property (Publisher)

 **True** to check the sequence of independent characters for Asian text. Read/write **Boolean**.


## Syntax

 _expression_. **SequenceCheck**

 _expression_A variable that represents a  **Options** object.


### Return Value

Boolean


## Example

This example enables sequence checking, allowing the user to input a valid sequence of independent characters to form valid character cells in South Asian text.


```vb
Sub CheckSequence() 
 Options.SequenceCheck = True 
End Sub
```


