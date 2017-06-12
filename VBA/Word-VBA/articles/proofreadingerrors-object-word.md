---
title: ProofreadingErrors Object (Word)
keywords: vbawd10.chm2491
f1_keywords:
- vbawd10.chm2491
ms.prod: word
ms.assetid: 53fb6382-4c08-83f3-1835-ac2633939758
ms.date: 06/08/2017
---


# ProofreadingErrors Object (Word)

A collection of spelling and grammatical errors for the specified document or range.


## Remarks

Use the  **SpellingErrors** or **GrammaticalErrors** property to return the **ProofreadingErrors** collection. The following example counts the spelling and grammatical errors in the selection and displays the results in a message box.


```vb
Set pr1 = Selection.Range.SpellingErrors 
 sc = pr1.Count 
Set pr2 = Selection.Range.GrammaticalErrors 
 gc = pr2.Count 
Msgbox "Spelling errors: " &; sc &; vbCr _ 
 &; "Grammatical errors: " &; gc
```

Use  **SpellingErrors** (Index), where Index is the index number, to return a single spelling error (represented by a **Range** object). The following example finds the second spelling error in the selection and then selects it.




```vb
Set myRange = Selection.Range.SpellingErrors(2) 
myRange.Select
```

Use  **GrammarErrors** (Index), where Index is the index number, to return a single grammatical error (represented by a **Range** object). The following example returns the sentence that contains the first grammatical error in the selection.




```vb
Set myRange = Selection.Range.GrammaticalErrors(1) 
Msgbox myRange.Text
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object. If all the words in the document or range are spelled correctly and are grammatically correct, the **Count** property for the **ProofreadingErrors** object returns 0 (zero) and the **SpellingChecked** and **GrammarChecked** properties return **True** .


 **Note**  There is no ProofreadingError object; instead, each item in the  **ProofreadingErrors** collection is a **Range** object that represents one spelling or grammatical error.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

