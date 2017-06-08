---
title: OMath.Type Property (Word)
keywords: vbawd10.chm176357481
f1_keywords:
- vbawd10.chm176357481
ms.prod: word
api_name:
- Word.OMath.Type
ms.assetid: d546f006-dc43-343b-808e-b0230d3f3391
ms.date: 06/08/2017
---


# OMath.Type Property (Word)

Returns or sets a  **WdOMathType** constant that represents whether an equation is displayed inline with the text around it or displayed on its own line. Read/write.


## Syntax

 _expression_ . **Type**

 _expression_ An expression that returns an **OMath** object.


## Example

The following example creates a new equation and sets it to display inline with the text.


```vb
Dim objRange As Range 
Dim objEq As OMath 
 
Set objRange = Selection.Range 
objRange.Text = "Celsius = (5/9)(Fahrenheit - 32)" 
Set objRange = Selection.OMaths.Add(objRange) 
Set objEq = objRange.OMaths(1) 
objEq.Type = wdOMathInline
```


## See also


#### Concepts


[OMath Object](omath-object-word.md)

