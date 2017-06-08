---
title: ReadabilityStatistics Object (Word)
ms.prod: word
ms.assetid: eabef73c-f837-435a-cfec-b76082cc0f7e
ms.date: 06/08/2017
---


# ReadabilityStatistics Object (Word)

A collection of  **[ReadabilityStatistic](readabilitystatistic-object-word.md)** objects for a document or range.


## Remarks

Use the  **ReadabilityStatistics** property to return the **ReadabilityStatistics** collection. The following example enumerates the readability statistics for the selection and displays each one in a message box.


```vb
For Each rs in Selection.Range.ReadabilityStatistics 
 Msgbox rs.Name &; " - " &; rs.Value 
Next rs
```

Use  **ReadabilityStatistics** (Index), where Index is the index number, to return a single **ReadabilityStatistic** object. The statistics are ordered as follows: Words, Characters, Paragraphs, Sentences, Sentences per Paragraph, Words per Sentence, Characters per Word, Passive Sentences, Flesch Reading Ease, and Flesch-Kincaid Grade Level. The following example returns the word count for the active document.




```vb
Set myRange = ActiveDocument.Content 
wordval = myRange.ReadabilityStatistics(1).Value 
Msgbox wordval
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

