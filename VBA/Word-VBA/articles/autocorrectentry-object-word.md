---
title: AutoCorrectEntry Object (Word)
keywords: vbawd10.chm2375
f1_keywords:
- vbawd10.chm2375
ms.prod: word
api_name:
- Word.AutoCorrectEntry
ms.assetid: 33173958-42eb-00ef-7f37-41f95ed47f87
ms.date: 06/08/2017
---


# AutoCorrectEntry Object (Word)

Represents a single AutoCorrect entry. The  **AutoCorrectEntry** object is a member of the **AutoCorrectEntries** collection. The **[AutoCorrectEntries](autocorrectentries-object-word.md)** collection includes the entries in the **AutoCorrect** dialog box.


## Remarks

Use  **[Entries](autocorrect-entries-property-word.md)** (index), where index is the AutoCorrect entry name or index number, to return a single **AutoCorrectEntry** object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown under **Replace** in the **AutoCorrect** dialog box. The following example sets the value of the AutoCorrect entry named "teh."


```
AutoCorrect.Entries("teh").Value = "the"
```

Use the  **[Apply](autocorrectentry-apply-method-word.md)** method to insert an AutoCorrect entry at the specified range. The following example adds an AutoCorrect entry and then inserts it in place of the selection.




```
AutoCorrect.Entries.Add Name:="hellp", Value:="hello" 
AutoCorrect.Entries("hellp").Apply Range:=Selection.Range
```

Use either the  **[Add](autocorrectentries-add-method-word.md)** or **[AddRichText](autocorrectentries-addrichtext-method-word.md)** method to add an AutoCorrect entry to the list of available entries. The following example adds a plain-text AutoCorrect entry for the misspelling of the word "their.'




```
AutoCorrect.Entries.Add Name:="thier", Value:="their"
```

The following example creates an AutoCorrect entry named "PMO" based on the text and formatting of the selection.




```
AutoCorrect.Entries.AddRichText Name:="PMO", Range:=Selection.Range
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

