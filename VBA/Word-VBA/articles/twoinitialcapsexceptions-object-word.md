---
title: TwoInitialCapsExceptions Object (Word)
keywords: vbawd10.chm2372
f1_keywords:
- vbawd10.chm2372
ms.prod: word
ms.assetid: 21af2d69-8d76-026d-2002-8d69b4ab8aef
ms.date: 06/08/2017
---


# TwoInitialCapsExceptions Object (Word)

A collection of  **[TwoInitialCapsException](twoinitialcapsexception-object-word.md)** objects that represent all the items listed in the **Don't correct** box on the **INitial CAps** tab in the **AutoCorrect Exceptions** dialog box.


## Remarks

Use the  **TwoInitialCapsExceptions** property to return the **TwoInitialCapsExceptions** collection. The following example displays the items in this collection.


```vb
For Each aCap In AutoCorrect.TwoInitialCapsExceptions 
 MsgBox aCap.Name 
Next aCap
```

If the  **TwoInitialCapsAutoAdd** property is **True** , words are automatically added to the list of initial-capital exceptions. Use the **Add** method to add an item to the **TwoInitialCapsExceptions** collection. The following example adds "Industry" to the list of initial-capital exceptions.




```
AutoCorrect.TwoInitialCapsExceptions.Add Name:="INdustry"
```

Use  **TwoInitialCapsExceptions** (Index), where Index is the initial cap name or the index number, to return a single **TwoInitialCapsException** object. The following example deletes the initial-capital item named "KMenu."




```
AutoCorrect.TwoInitialCapsExceptions("KMenu").Delete
```

The index number represents the position of the initial-capital exception in the  **TwoInitialCapsExceptions** collection. The following example displays the name of the first item in the **TwoInitialCapsExceptions** collection.




```vb
MsgBox AutoCorrect.TwoInitialCapsExceptions(1).Name
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


