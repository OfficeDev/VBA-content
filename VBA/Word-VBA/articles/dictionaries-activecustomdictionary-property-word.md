---
title: Dictionaries.ActiveCustomDictionary Property (Word)
keywords: vbawd10.chm162267139
f1_keywords:
- vbawd10.chm162267139
ms.prod: word
api_name:
- Word.Dictionaries.ActiveCustomDictionary
ms.assetid: 2f6a3fa0-4e7d-a7f0-20ce-49688dab7191
ms.date: 06/08/2017
---


# Dictionaries.ActiveCustomDictionary Property (Word)

Returns or sets a  **[Dictionary](dictionary-object-word.md)** object that represents the custom dictionary to which words will be added. Read/write.


## Syntax

 _expression_ . **ActiveCustomDictionary**

 _expression_ A variable that represents a **[Dictionaries](dictionaries-object-word.md)** collection.


## Example

This example displays the full path and file name of the active custom dictionary.


```vb
Set dicCustom = Application.CustomDictionaries.ActiveCustomDictionary 
MsgBox dicCustom.Path &; Application.PathSeparator &; dicCustom.Name
```

This example clears all existing custom dictionaries, adds a custom dictionary named "Home.dic," and then loads the new dictionary.




```vb
Dim dicCustom As Dictionary 
 
Application.CustomDictionaries.ClearAll 
 
Set dicCustom = Application.CustomDictionaries _ 
 .Add(FileName:="C:\Program Files" _ 
 &; "\Microsoft Office\Office\Home.dic") 
Application.CustomDictionaries.ActiveCustomDictionary = dicCustom
```


## See also


#### Concepts


[Dictionaries Collection Object](dictionaries-object-word.md)

