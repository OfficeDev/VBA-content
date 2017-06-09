---
title: Application.HangulHanjaDictionaries Property (Word)
keywords: vbawd10.chm158335086
f1_keywords:
- vbawd10.chm158335086
ms.prod: word
api_name:
- Word.Application.HangulHanjaDictionaries
ms.assetid: 453e2a77-f363-5afc-d9a3-26f8b6516b4c
ms.date: 06/08/2017
---


# Application.HangulHanjaDictionaries Property (Word)

Returns a  **[HangulHanjaConversionDictionaries](hangulhanjaconversiondictionaries-object-word.md)** collection that represents all the active custom conversion dictionaries.


## Syntax

 _expression_ . **HangulHanjaDictionaries**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

Active custom conversion dictionaries are marked with a check in the  **Custom Dictionaries** dialog box. Click **Options**, click the  **Spelling &; Grammar** tab, and then click the **Custom Dictionaries** button.

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds a new, blank custom dictionary to the collection. The path and file name of the new custom dictionary are then displayed in a message box.


```vb
Set myHome = _ 
 HangulHanjaDictionaries.Add(Filename:="Home.hhd") 
Msgbox myHome.Path &; Application.PathSeparator _ 
 &; myHome.Name
```

This example deactivates all custom dictionaries but does not delete the custom dictionary files.




```
HangulHanjaDictionaries.ClearAll
```

This example displays the name of each custom dictionary in the collection.




```vb
For Each di In HangulHanjaDictionaries 
 MsgBox di.Name 
Next di
```


## See also


#### Concepts


[Application Object](application-object-word.md)

