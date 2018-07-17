---
title: HangulAndAlphabetException Object (Word)
keywords: vbawd10.chm2514
f1_keywords:
- vbawd10.chm2514
ms.prod: word
api_name:
- Word.HangulAndAlphabetException
ms.assetid: f383505b-1f98-117c-e170-606403ad1508
ms.date: 06/08/2017
---


# HangulAndAlphabetException Object (Word)

Represents a single Hangul or alphabet AutoCorrect exception. The  **HangulAndAlphabetException** object is a member of the **HangulAndAlphabetExceptions** collection.


## Remarks

Use  **HangulAndAlphabetExceptions** (Index), where Index is the Hangul or alphabet AutoCorrect exception name or the index number, to return a single **HangulAndAlphabetException** object. The following example deletes the alphabet AutoCorrect exception named "hello."


```
AutoCorrect.HangulAndAlphabetExceptions("hello").Delete
```

The index number represents the position of the Hangul or alphabet AutoCorrect exception in the  **[HangulAndAlphabetExceptions](hangulandalphabetexceptions-object-word.md)** collection. The following example displays the name of the first item in the **[HangulAndAlphabetExceptions](hangulandalphabetexceptions-object-word.md)** collection.




```vb
MsgBox AutoCorrect.HangulAndAlphabetExceptions(1).Name
```

If the value of the  **HangulAndAlphabetAutoAdd** property is **True** , words are automatically added to the list of Hangul and alphabet AutoCorrect exceptions. Use the **Add** method to add an item to the **[HangulAndAlphabetExceptions](hangulandalphabetexceptions-object-word.md)** collection. The following example adds "goodbye" to the list of alphabet AutoCorrect exceptions.




```
AutoCorrect.HangulAndAlphabetExceptions.Add Name:="goodbye"
```


 **Note**  The  **[HangulAndAlphabetExceptions](hangulandalphabetexceptions-object-word.md)** collection includes all Hangul and alphabet AutoCorrect exceptions and corresponds to the items listed on the **Korean** tab in the **AutoCorrect Exceptions** dialog box.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

