---
title: Application.Languages Property (Word)
keywords: vbawd10.chm158334990
f1_keywords:
- vbawd10.chm158334990
ms.prod: word
api_name:
- Word.Application.Languages
ms.assetid: f81cfcb6-33e2-bb8e-2ac4-b4f9df833946
ms.date: 06/08/2017
---


# Application.Languages Property (Word)

Returns a  **[Languages](languages-object-word.md)** collection that represents the proofing languages listed in the **Language** dialog box.


## Syntax

 _expression_ . **Languages**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


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

This example uses the  `aLang()` array to store the proofing language names.




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


[Application Object](application-object-word.md)

