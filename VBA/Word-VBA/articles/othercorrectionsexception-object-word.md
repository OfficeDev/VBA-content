---
title: OtherCorrectionsException Object (Word)
keywords: vbawd10.chm2529
f1_keywords:
- vbawd10.chm2529
ms.prod: word
api_name:
- Word.OtherCorrectionsException
ms.assetid: f3c92186-0d3a-0585-b545-3a94e27a7d7b
ms.date: 06/08/2017
---


# OtherCorrectionsException Object (Word)

Represents a single AutoCorrect exception. The  **OtherCorrectionsException** object is a member of the **OtherCorrectionsExceptions** collection.


## Remarks

The  **OtherCorrectionsExceptions** collection includes all words that Microsoft Word won't correct automatically. This list corresponds to the list of AutoCorrect exceptions on the **Other Corrections** tab in the **AutoCorrect Exceptions** dialog box.

Use  **OtherCorrectionsExceptions** (Index), where Index is the AutoCorrect exception name or the index number, to return a single **OtherCorrectionsException** object. The following example deletes "WTop" from the list of AutoCorrect exceptions.




```
AutoCorrect.OtherCorrectionsExceptions("WTop").Delete
```

The index number represents the position of the AutoCorrect exception in the  **OtherCorrectionsExceptions** collection. The following example displays the name of the first item in the **OtherCorrectionsExceptions** collection.




```vb
MsgBox AutoCorrect.OtherCorrectionsExceptions(1).Name
```

If the value of the  **OtherCorrectionsAutoAdd** property is **True** , words are automatically added to the list of AutoCorrect exceptions. Use the **Add** method to add an item to the **OtherCorrectionsExceptions** collection. The following example adds "TipTop" to the list of AutoCorrect exceptions.




```
AutoCorrect.OtherCorrectionsExceptions.Add Name:="TipTop"
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


