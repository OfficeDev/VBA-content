---
title: FindReplace.MatchCase Property (Publisher)
keywords: vbapb10.chm8323080
f1_keywords:
- vbapb10.chm8323080
ms.prod: publisher
api_name:
- Publisher.FindReplace.MatchCase
ms.assetid: 4fabf2f8-f1e4-bc70-e8e6-96dd09cd23d8
ms.date: 06/08/2017
---


# FindReplace.MatchCase Property (Publisher)

Sets or returns a  **Boolean** that represents the case sensitivity of the search operation. Read/write.


## Syntax

 _expression_. **MatchCase**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

Boolean


## Remarks

The default value for  **MatchCase** is **False**.


## Example

This example will select the first occurrence of the word "factory" regardless of case.


```vb
With ActiveDocument.Find 
 .Clear 
 .MatchCase = False 
 .FindText = "factory" 
 .Execute 
End With 

```


