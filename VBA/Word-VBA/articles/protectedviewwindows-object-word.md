---
title: ProtectedViewWindows Object (Word)
ms.prod: word
api_name:
- Word.ProtectedViewWindows
ms.assetid: 62c2f4d5-1080-548e-730b-388308144dfe
ms.date: 06/08/2017
---


# ProtectedViewWindows Object (Word)

A collection of all the [ProtectedViewWindow](protectedviewwindow-object-word.md) objects that are currently open in Word.


## Remarks

Use the  **ProtectedViewWindows** property to return the **ProtectedViewWindows** collection.


## Example

The following code example displays the number of protected view windows that are open.


```vb
MsgBox "There are " &; ProtectedViewWindows.Count &; _ 
 " protected view windows currently open."
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


