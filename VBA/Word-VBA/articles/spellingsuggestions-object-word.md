---
title: SpellingSuggestions Object (Word)
keywords: vbawd10.chm2474
f1_keywords:
- vbawd10.chm2474
ms.prod: word
ms.assetid: 7e0fb008-e43c-c4cb-b7d2-9436d039a070
ms.date: 06/08/2017
---


# SpellingSuggestions Object (Word)

A collection of  **SpellingSuggestion** objects that represent all the suggestions for a specified word or for the first word in the specified range.


## Remarks

Use the  **GetSpellingSuggestions** method to return the **SpellingSuggestions** collection. The **SpellingSuggestions** method, when applied to the **Application** object, must specify the word to be checked. When the **GetSpellingSuggestions** method is applied to a range, the first word in the range is checked. The following example checks to see whether there are any spelling suggestions for any of the words in the active document. If there are, the suggestions are displayed in message boxes.


```vb
For Each wd In ActiveDocument.Words 
 Set sugg = wd.GetSpellingSuggestions 
 If sugg.Count <> 0 Then 
 For Each ss In sugg 
 MsgBox ss.Name 
 Next ss 
 End If 
Next wd
```

You cannot add suggestions to or remove suggestions from the collection of spelling suggestions. Spelling suggestions are derived from main and custom dictionary files.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


