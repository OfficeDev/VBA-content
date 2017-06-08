---
title: SynonymInfo.AntonymList Property (Word)
keywords: vbawd10.chm161153032
f1_keywords:
- vbawd10.chm161153032
ms.prod: word
api_name:
- Word.SynonymInfo.AntonymList
ms.assetid: 4ba1a1b1-79c7-e230-2eae-7b64182fa232
ms.date: 06/08/2017
---


# SynonymInfo.AntonymList Property (Word)

Returns a list of antonyms for the word or phrase. The list is returned as an array of strings. Read-only  **Variant** .


## Syntax

 _expression_ . **AntonymList**

 _expression_ An expression that returns a **[SynonymInfo](synonyminfo-object-word.md)** object.


## Remarks

The AntonymList property is a property of the SynonymInfo object, which can be returned from either a range or the application. If this object is returned from the application, you specify the word to look up and the language to use. When the object is returned from a range, the range is looked up using the language of the range.


## Example

This example returns a list of antonyms for the word "big" in U.S. English.


```vb
Dim arrayAntonyms As Variant 
Dim intLoop As Integer 
 
arrayAntonyms = SynonymInfo(Word:="big", _ 
 LanguageID:=wdEnglishUS).AntonymList 
For intLoop = 1 To UBound(arrayAntonyms) 
 MsgBox arrayAntonyms(intLoop) 
Next intLoop
```

This example returns a list of antonyms for the word or phrase in the selection and displays them in the Immediate window in the Visual Basic Editor.




```vb
Dim arrayAntonyms As Variant 
Dim intLoop As Integer 
 
arrayAntonyms = Selection.Range.SynonymInfo.AntonymList 
If UBound(arrayAntonyms) <> 0 Then 
 For intLoop = 1 To UBound(arrayAntonyms) 
 Debug.Print arrayAntonyms(intLoop) &; Str(intLoop) 
 Next intLoop 
Else 
 MsgBox "No antonyms were found." 
End If
```

This example returns a list of antonyms, if there are any, for the third word in the active document.




```vb
Dim rngTemp As Range 
Dim arrayAntonyms As Variant 
Dim intLoop As Integer 
 
Set rngTemp = ActiveDocument.Words(3) 
 
arrayAntonyms = rngTemp.SynonymInfo.AntonymList 
If UBound(arrayAntonyms) = 0 Then 
 MsgBox "There are no antonyms for the third word." 
Else 
 For intLoop = 1 To UBound(arrayAntonyms) 
 MsgBox arrayAntonyms(intLoop) 
 Next intLoop 
End If
```


## See also


#### Concepts


[SynonymInfo Object](synonyminfo-object-word.md)

