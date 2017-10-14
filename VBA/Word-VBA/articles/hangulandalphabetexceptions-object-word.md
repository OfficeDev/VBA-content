---
title: HangulAndAlphabetExceptions Object (Word)
ms.prod: word
ms.assetid: ddb128f0-3752-5d38-e65a-767f17d86294
ms.date: 06/08/2017
---


# HangulAndAlphabetExceptions Object (Word)

A collection of  **HangulAndAlphabetException** objects that represents all Hangul and alphabet AutoCorrect exceptions.


## Remarks

Use the  **HangulAndAlphabetExceptions** property to return the **HangulAndAlphabetExceptions** collection. The following example displays the items in this collection.


```vb
For Each aHan In AutoCorrect.HangulAndAlphabetExceptions 
 MsgBox aHan.Name 
Next aHan
```

If the value of the  **HangulAndAlphabetAutoAdd** property is **True** , words are automatically added to the list of Hangul and alphabet AutoCorrect exceptions. Use the **Add** method to add an item to the **HangulAndAlphabetExceptions** collection. The following example adds "hello" to the list of alphabet AutoCorrect exceptions.




```
AutoCorrect.HangulAndAlphabetExceptions.Add Name:="hello"
```

Use  **HangulAndAlphabetExceptions** (Index), where Index is the Hangul or alphabet AutoCorrect exception name or the index number, to return a single **[HangulAndAlphabetException](hangulandalphabetexception-object-word.md)** object. The following example deletes the alphabet AutoCorrect exception named "goodbye."




```
AutoCorrect.HangulAndAlphabetExceptions("goodbye").Delete
```

The index number represents the position of the hangul or alphabet AutoCorrect exception in the  **HangulAndAlphabetExceptions** collection. The following example displays the name of the first item in the **HangulAndAlphabetExceptions** collection.




```vb
MsgBox AutoCorrect.HangulAndAlphabetExceptions(1).Name
```


 **Note**  The list of Hangul and alphabet AutoCorrect exceptions corresponds to the list of AutoCorrect exceptions on the  **Korean** tab in the **AutoCorrect Exceptions** dialog box.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

