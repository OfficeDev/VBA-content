---
title: CoAuthLocks Object (Word)
ms.prod: word
api_name:
- Word.CoAuthLocks
ms.assetid: 589763ed-8463-6988-3817-9c2152506d16
ms.date: 06/08/2017
---


# CoAuthLocks Object (Word)

A collection of  **[CoAuthLock](coauthlock-object-word.md)** objects.


## Remarks

Use the  **[Locks](coauthlock-object-word.md)** property to return the **CoAuthLocks** collection.


## Example

The following code example displays the number of locks in the active document.


```vb
MsgBox ActiveDocument.CoAuthoring.Locks.Count
```


## See also


#### Concepts


[CoAuthoring.Locks Property](coauthoring-locks-property-word.md)
#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


