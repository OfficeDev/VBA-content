---
title: AutoCorrect Object (Word)
keywords: vbawd10.chm2377
f1_keywords:
- vbawd10.chm2377
ms.prod: word
api_name:
- Word.AutoCorrect
ms.assetid: dea9b72c-4378-05ac-ec4b-51cf3af3f2a3
ms.date: 06/08/2017
---


# AutoCorrect Object (Word)

Represents the AutoCorrect functionality in Word.


## Remarks

Use the  **[AutoCorrect](application-autocorrect-property-word.md)** property to return the **AutoCorrect** object. The following example enables the AutoCorrect options and creates an AutoCorrect entry.


```vb
With AutoCorrect 
 .CorrectCapsLock = True 
 .CorrectDays = True 
 .Entries.Add Name:="usualy", Value:="usually" 
End With
```

The  **[Entries](autocorrect-entries-property-word.md)** property returns the **[Entries](autocorrect-entries-property-word.md)** object that represents the AutoCorrect entries in the **AutoCorrect** dialog box.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


