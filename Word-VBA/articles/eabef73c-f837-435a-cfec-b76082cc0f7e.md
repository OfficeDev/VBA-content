
# ReadabilityStatistics Object (Word)

 **Last modified:** July 28, 2015

A collection of  ** [ReadabilityStatistic](5e82c44d-fc6d-9586-816b-0c46c4a01f3b.md)**objects for a document or range.

## Remarks

Use the  **ReadabilityStatistics** property to return the **ReadabilityStatistics** collection. The following example enumerates the readability statistics for the selection and displays each one in a message box.


```
For each rs in Selection.Range.ReadabilityStatistics 
 Msgbox rs.Name &amp; " - " &amp; rs.Value 
Next rs
```

Use  **ReadabilityStatistics**(Index), where Index is the index number, to return a single  **ReadabilityStatistic** object. The statistics are ordered as follows: Words, Characters, Paragraphs, Sentences, Sentences per Paragraph, Words per Sentence, Characters per Word, Passive Sentences, Flesch Reading Ease, and Flesch-Kincaid Grade Level. The following example returns the word count for the active document.




```
Set myRange = ActiveDocument.Content 
wordval = myRange.ReadabilityStatistics(1).Value 
Msgbox wordval
```


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [ReadabilityStatistics Object Members](4e7dde67-0de5-89fc-3061-ab67bb2f03ec.md)
