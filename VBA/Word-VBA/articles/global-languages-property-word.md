---
title: Global.Languages Property (Word)
keywords: vbawd10.chm163119118
f1_keywords:
- vbawd10.chm163119118
ms.prod: word
api_name:
- Word.Global.Languages
ms.assetid: 6f0d87f8-f0f8-5865-3ba5-2a383c212998
ms.date: 06/08/2017
---


# Global.Languages Property (Word)

Returns a  **Languages** collection that represents the proofing languages listed in the **Language** dialog box.


## Syntax

 _expression_ . **Languages**

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example returns the full path and file name of the active spelling dictionary.


```vb
Dim dicSpell As Dictionary 
 
Set dicSpell = _ 
 Languages(Selection.LanguageID).ActiveSpellingDictionary 
 
MsgBox dicSpell.Path &; Application.PathSeparator &; dicSpell.Name
```

This example uses the  _aLang()_ array to store the proofing language names.




```vb
Dim intCount As Integer 
Dim langLoop As Language 
Dim aLang(Languages.Count - 1) As String 
 
intCount = 0 
For Each langLoop In Languages 
 aLang(intCount) = langLoop.NameLocal 
 intCount = intCount + 1 
Next langLoop
```


## See also


#### Concepts


[Global Object](global-object-word.md)

