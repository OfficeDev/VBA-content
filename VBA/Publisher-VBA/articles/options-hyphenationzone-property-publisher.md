---
title: Options.HyphenationZone Property (Publisher)
keywords: vbapb10.chm1048593
f1_keywords:
- vbapb10.chm1048593
ms.prod: publisher
api_name:
- Publisher.Options.HyphenationZone
ms.assetid: ed0e90de-4a2a-3c8a-27f1-e8c7c1f0e174
ms.date: 06/08/2017
---


# Options.HyphenationZone Property (Publisher)

Returns or sets a  **Variant** that represents the maximum amount of space that Microsoft Publisher leaves between the end of the last word in a line and the right margin. Read/write.


## Syntax

 _expression_. **HyphenationZone**

 _expression_A variable that represents a  **Options** object.


### Return Value

Variant


## Example

This example turns on automatic hyphenation and specifies the maximum amount of space between the end of the last word and the right margin equal to one inch (72 points).


```vb
Sub SetHyphenationZone() 
 With Options 
 .AutoHyphenate = True 
 .HyphenationZone = 72 
 End With 
End Sub
```


