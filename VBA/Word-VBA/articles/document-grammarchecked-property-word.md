---
title: Document.GrammarChecked Property (Word)
keywords: vbawd10.chm158007366
f1_keywords:
- vbawd10.chm158007366
ms.prod: word
api_name:
- Word.Document.GrammarChecked
ms.assetid: 30de1405-196a-e8e0-f5af-710b217ea3fd
ms.date: 06/08/2017
---


# Document.GrammarChecked Property (Word)

 **True** if a grammar check has been run on the specified range or document. Read/write **Boolean** .


## Syntax

 _expression_ . **GrammarChecked**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Returns  **False** if all or some of the specified document has not been checked for grammar. To recheck the grammar in a document, set the **GrammarChecked** property to **False** .


## Example

This example determines whether grammar has been checked in the active document. If it has, the word count is displayed. If grammar has not been checked, a spelling and grammar check is started.


```vb
Set myStat = ActiveDocument.ReadabilityStatistics 
passGram = ActiveDocument.GrammarChecked 
If passGram = True Then 
 Msgbox myStat(1).Name &; " - " &; myStat(1).Value 
Else 
 ActiveDocument.CheckGrammar 
End If
```

This example sets the GrammarChecked property to False for the active document, and then it runs a grammar check again.




```vb
ActiveDocument.GrammarChecked
```




```vb
= False
```




```vb
ActiveDocument.CheckGrammar
```


## See also


#### Concepts


[Document Object](document-object-word.md)

