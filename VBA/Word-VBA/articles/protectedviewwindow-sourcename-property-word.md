---
title: ProtectedViewWindow.SourceName Property (Word)
keywords: vbawd10.chm231735306
f1_keywords:
- vbawd10.chm231735306
ms.prod: word
api_name:
- Word.ProtectedViewWindow.SourceName
ms.assetid: 744639ae-dd9f-cf85-f15f-f2c753fc9d9d
ms.date: 06/08/2017
---


# ProtectedViewWindow.SourceName Property (Word)

Returns the name of the source file for the specified protected view window. Read-only  **String** .


## Syntax

 _expression_ . **SourceName**

 _expression_ An expression that returns a **[ProtectedViewWindow](protectedviewwindow-object-word.md)** object.


## Remarks

This property does not return the path for the source file.


## Example

The following code example returns the path and name of the document associated with the specified protected view window.


```vb
MsgBox ActiveProtectedViewWindow.SourcePath &; "\" _ 
 &; ActiveProtectedViewWindow.SourceName 

```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-word.md)

