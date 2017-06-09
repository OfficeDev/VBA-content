---
title: Application.CustomDictionaries Property (Word)
keywords: vbawd10.chm158335071
f1_keywords:
- vbawd10.chm158335071
ms.prod: word
api_name:
- Word.Application.CustomDictionaries
ms.assetid: 1c6dca90-70f0-6b52-72d1-debda33d2ba0
ms.date: 06/08/2017
---


# Application.CustomDictionaries Property (Word)

Returns a  **[Dictionaries](dictionaries-object-word.md)** object that represents the collection of active custom dictionaries. Read-only.


## Syntax

 _expression_ . **CustomDictionaries**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

Active custom dictionaries are marked with a check in the  **Custom Dictionaries** dialog box. For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds a new, blank custom dictionary to the collection. The path and file name of the new custom dictionary are then displayed in a message box.


```vb
Dim dicHome As Dictionary 
Set dicHome = CustomDictionaries.Add(Filename:="Home.dic") 
Msgbox dicHome.Path &; Application.PathSeparator &; dicHome.Name
```

This example removes all custom dictionaries so that no custom dictionaries are active. The custom dictionary files aren't deleted, though.




```
CustomDictionaries.ClearAll
```

This example displays the name of each custom dictionary in the collection.




```vb
Dim dicLoop As Dictionary 
 
For Each dicLoop In CustomDictionaries 
 MsgBox dicLoop.Name 
Next dicLoop
```


## See also


#### Concepts


[Application Object](application-object-word.md)

