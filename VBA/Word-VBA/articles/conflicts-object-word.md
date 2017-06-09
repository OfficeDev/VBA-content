---
title: Conflicts Object (Word)
ms.prod: word
api_name:
- Word.Conflicts
ms.assetid: 476e8f6d-c93e-b372-2fa7-1c9a4a84a182
ms.date: 06/08/2017
---


# Conflicts Object (Word)

 A collection of[Conflict](conflict-object-word.md) objects that represents the conflicts in a document. The type of a **Conflict** object is specified by the[WdRevisionType](wdrevisiontype-enumeration-word.md) enumeration.


## Remarks

Use the [Conflicts](coauthoring-conflicts-property-word.md) property to return the **Conflicts** collection for a document. Use Conflicts( _Index_ ), where _Index_ is the conflict index number, to return a single **Conflict** object.


## Example

The following code example accepts the first conflict in the active document.


```vb
ActiveDocument.CoAuthoring.Conflicts(1).Accept 

```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

