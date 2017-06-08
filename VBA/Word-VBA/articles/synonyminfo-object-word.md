---
title: SynonymInfo Object (Word)
keywords: vbawd10.chm2459
f1_keywords:
- vbawd10.chm2459
ms.prod: word
api_name:
- Word.SynonymInfo
ms.assetid: 0af2d733-a038-1f67-ddca-2b05b3af1b7c
ms.date: 06/08/2017
---


# SynonymInfo Object (Word)

Represents the information about synonyms, antonyms, related words, or related expressions for the specified range or a given string.


## Remarks

Use the  **SynonymInfo** property to return a **SynonymInfo** object. The **SynonymInfo** object can be returned either from a range or from Microsoft Office Word. If it is returned from Word, you specify the lookup word or phrase and a proofing language ID. If it is returned from a range, Word uses the specified range as the lookup word. The following example returns a **SynonymInfo** object from Word.


```
temp = SynonymInfo(Word:="meant", LanguageID:=wdEnglishUS).Found
```

The following example returns a  **SynonymInfo** object from a range.




```
temp = Selection.Range.SynonymInfo.Found
```

The  **Found** property, used in the preceding examples, returns **True** if any information is found in the thesaurus for the specified range or for Word. Note, however, that this property returns **True** not only if synonyms are found but also if related words, related expressions, or antonyms are found.

Many of the properties of the  **SynonymInfo** object return a **Variant** that contains an array of strings. When working with these properties, you can assign the returned array to a variable and then index the variable to see the elements in the array. In the following example, _Slist_ is assigned the synonym list for the first meaning of the selected word or phrase. The **UBound** function finds the upper bound of the array, and then each element is displayed in a message box.




```vb
Slist = Selection.Range.SynonymInfo.SynonymList(1) 
For i = 1 To UBound(Slist) 
 Msgbox Slist(i) 
Next i
```

You can check the value of the  **MeaningCount** property to prevent potential errors in your code. The following example returns a list of synonyms for the second meaning for the word or phrase in the selection and displays these synonyms in the **Immediate** pane.




```vb
Set synInfo = Selection.Range.SynonymInfo 
If synInfo.MeaningCount >= 2 Then 
 synList = synInfo.SynonymList(2) 
 For i = 1 To UBound(synList) 
 Debug.Print synList(i) 
 Next i 
Else 
 MsgBox "There is no second meaning for the selection." 
End If
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


