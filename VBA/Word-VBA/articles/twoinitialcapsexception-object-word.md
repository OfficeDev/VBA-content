---
title: TwoInitialCapsException Object (Word)
keywords: vbawd10.chm2371
f1_keywords:
- vbawd10.chm2371
ms.prod: word
api_name:
- Word.TwoInitialCapsException
ms.assetid: 48e89297-4137-960b-a92a-2a70929e298a
ms.date: 06/08/2017
---


# TwoInitialCapsException Object (Word)

Represents a single initial-capital AutoCorrect exception. The  **TwoInitialCapsException** object is a member of the **[TwoInitialCapsExceptions](twoinitialcapsexceptions-object-word.md)** collection. The **TwoInitialCapsExceptions** collection includes all the items listed in the **Don't correct** box on the **INitial CAps** tab in the **AutoCorrect Exceptions** dialog box.


## Remarks

Use  **TwoInitialCapsExceptions** (Index), where Index is the initial capital exception name or the index number, to return a single **TwoInitialCapsException** object. The following example deletes the initial-capital exception named "KMenu."


```
AutoCorrect.TwoInitialCapsExceptions("KMenu").Delete
```

The index number represents the position of the initial-capital exception in the  **TwoInitialCapsExceptions** collection. The following example displays the name of the first item in the **TwoInitialCapsExceptions** collection.




```vb
MsgBox AutoCorrect.TwoInitialCapsExceptions(1).Name
```

If the  **TwoInitialCapsAutoAdd** property is **True** , words are automatically added to the list of initial-capital exceptions. Use the **Add** method to add an item to the **TwoInitialCapsExceptions** collection. The following example adds "INdustry" to the list of initial-capital exceptions.




```
AutoCorrect.TwoInitialCapsExceptions.Add Name:="INdustry"
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


