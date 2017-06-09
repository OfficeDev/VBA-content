---
title: ProtectedViewWindow.Activate Method (Word)
keywords: vbawd10.chm231735396
f1_keywords:
- vbawd10.chm231735396
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Activate
ms.assetid: a784fceb-38b9-2fc4-6c71-fcfb17b53dfe
ms.date: 06/08/2017
---


# ProtectedViewWindow.Activate Method (Word)

Activates the specified protected view window.


## Syntax

 _expression_ . **Activate**

 _expression_ An expression that returns a **[ProtectedViewWindow Object](protectedviewwindow-object-word.md)** object.


### Return Value

Nothing


## Example

The following code example activates the next protected view window in the [ProtectedViewWindows](protectedviewwindows-object-word.md) collection.


```vb
Dim pvWindow As ProtectedViewWindow 
 
' At least one document must be open in protected view for this statement to execute. 
ProtectedViewWindows(1).Activate
```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-word.md)

