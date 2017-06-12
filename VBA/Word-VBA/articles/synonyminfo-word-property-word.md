---
title: SynonymInfo.Word Property (Word)
keywords: vbawd10.chm161153025
f1_keywords:
- vbawd10.chm161153025
ms.prod: word
api_name:
- Word.SynonymInfo.Word
ms.assetid: ec019502-6dc7-16f8-b019-957b00a7e3d1
ms.date: 06/08/2017
---


# SynonymInfo.Word Property (Word)

Returns the word or phrase that was looked up by the thesaurus. Read-only  **String** .


## Syntax

 _expression_ . **Word**

 _expression_ An expression that returns a **[SynonymInfo](synonyminfo-object-word.md)** object.


## Remarks

The thesaurus will sometimes look up a shortened version of the string or range used to return the  **SynonymInfo** object. The **Word** property allows you to see the exact string that was used.


## Example

This example returns a list of synonyms for the first meaning of the third word in the active document.


```vb
Sub Syn() 
 Dim mySynObj As Object 
 Dim SList As Variant 
 Dim i As Variant 
 
 Set mySynObj = ActiveDocument.Words(3).SynonymInfo 
 SList = mySynObj.SynonymList(1) 
 For i = 1 To UBound(SList) 
 MsgBox "A synonym for " &; mySynObj.Word _ 
 &; " is " &; SList(i) 
 Next i 
End Sub
```

This example checks to make sure that the word or phrase that was looked up isn't empty. If it is not, the example returns a list of synonyms for the first meaning of the word or phrase.




```vb
Sub SelectWord() 
 Dim mySynObj As Object 
 Dim SList As Variant 
 Dim i As Variant 
 
 Set mySynObj = Selection.Range.SynonymInfo 
 If mySynObj.Word = "" Then 
 MsgBox "Please select a word or phrase" 
 Else 
 SList = mySynObj.SynonymList(1) 
 For i = 1 To UBound(SList) 
 MsgBox "A synonym for " &; mySynObj.Word _ 
 &; " is " &; SList(i) 
 Next i 
 End If 
End Sub
```


## See also


#### Concepts


[SynonymInfo Object](synonyminfo-object-word.md)

