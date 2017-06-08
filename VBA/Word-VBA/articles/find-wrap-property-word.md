---
title: Find.Wrap Property (Word)
keywords: vbawd10.chm162529307
f1_keywords:
- vbawd10.chm162529307
ms.prod: word
api_name:
- Word.Find.Wrap
ms.assetid: 2d6823f3-93aa-383c-af28-d44e6a8a83e2
ms.date: 06/08/2017
---


# Find.Wrap Property (Word)

Returns or sets what happens if the search begins at a point other than the beginning of the document and the end of the document is reached (or vice versa if  **Forward** is set to **False** ) or if the search text isn't found in the specified selection or range. Read/write **WdFindWrap** .


## Syntax

 _expression_ . **Wrap**

 _expression_ Required. A variable that represents a **[Find](find-object-word.md)** object.


## Example

The following example searches forward through the document for the word "aspirin." When the end of the document is reached, the search continues at the beginning of the document. If the word "aspirin" is found, it is selected.


```vb
Sub WordFind() 
 With Selection.Find 
 .Forward = True 
 .ClearFormatting 
 .MatchWholeWord = True 
 .MatchCase = False 
 .Wrap = wdFindContinue 
 .Execute FindText:="aspirin" 
 End With 
End Sub
```


## See also


#### Concepts


[Find Object](find-object-word.md)

