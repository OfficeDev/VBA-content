---
title: FirstLetterException Object (Word)
keywords: vbawd10.chm2373
f1_keywords:
- vbawd10.chm2373
ms.prod: word
api_name:
- Word.FirstLetterException
ms.assetid: e365a683-010a-a074-5563-f0cac1f410b2
ms.date: 06/08/2017
---


# FirstLetterException Object (Word)

Represents an abbreviation excluded from automatic correction. The  **[FirstLetterExceptions](firstletterexceptions-object-word.md)** object is a member of the **FirstLetterExceptions** collection.


## Remarks

The  **FirstLetterExceptions** collection includes all the excluded abbreviations.The first character following a period is automatically capitalized when the **CorrectSentenceCaps** property is set to **True** . The character you type following an item in the **FirstLetterExceptions** collection isn't capitalized.

Use  **FirstLetterExceptions** (Index), where Index is the abbreviation or the index number, to return a single **FirstLetterException** object. The following example deletes the abbreviation "appt." from the **[FirstLetterExceptions](firstletterexceptions-object-word.md)** collection.




```
AutoCorrect.FirstLetterExceptions("appt.").Delete
```

The following example displays the name of the first item in the  **[FirstLetterExceptions](firstletterexceptions-object-word.md)** collection.




```vb
MsgBox AutoCorrect.FirstLetterExceptions(1).Name
```

Use the  **Add** method to add an abbreviation to the list of first-letter exceptions. The following example adds the abbreviation "addr." to this list.




```
AutoCorrect.FirstLetterExceptions.Add Name:="addr."
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

