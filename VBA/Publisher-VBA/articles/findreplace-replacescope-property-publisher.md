---
title: FindReplace.ReplaceScope Property (Publisher)
keywords: vbapb10.chm8323085
f1_keywords:
- vbapb10.chm8323085
ms.prod: publisher
api_name:
- Publisher.FindReplace.ReplaceScope
ms.assetid: 555fe65b-9edb-8888-03f0-15ce34813d5f
ms.date: 06/08/2017
---


# FindReplace.ReplaceScope Property (Publisher)

Specifies how many scope replacements are to be made: one, all, or none. Read/write.


## Syntax

 _expression_. **ReplaceScope**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

PbReplaceScope


## Remarks

The  **ReplaceScope** property value can be one of the **[PbReplaceScope](pbreplacescope-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.

The default setting of the  **ReplaceScope** property is **pbReplaceScopeNone**.


## Example

The following example replaces all occurrences of the word "hi" with "hello" in the active document.


```vb
With ActiveDocument.Find 
 .Clear 
 .FindText = "hi" 
 .ReplaceWithText = "hello" 
 .MatchWholeWord = True 
 .ReplaceScope = pbReplaceScopeAll 
 .Execute 
End With
```


