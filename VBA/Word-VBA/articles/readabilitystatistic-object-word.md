---
title: ReadabilityStatistic Object (Word)
keywords: vbawd10.chm2479
f1_keywords:
- vbawd10.chm2479
ms.prod: word
api_name:
- Word.ReadabilityStatistic
ms.assetid: 5e82c44d-fc6d-9586-816b-0c46c4a01f3b
ms.date: 06/08/2017
---


# ReadabilityStatistic Object (Word)

Represents one of the readability statistics for a document or range. The  **ReadabilityStatistic** object is a member of the **ReadabilityStatistics** collection.


## Remarks

Use  **ReadabilityStatistics** (Index), where Index is the index number, to return a single **ReadabilityStatistic** object. The statistics are ordered as follows: Words, Characters, Paragraphs, Sentences, Sentences per Paragraph, Words per Sentence, Characters per Word, Passive Sentences, Flesch Reading Ease, and Flesch-Kincaid Grade Level. The following example returns the character count for the active document.


```
Msgbox ActiveDocument.Content.ReadabilityStatistics(2).Value
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


