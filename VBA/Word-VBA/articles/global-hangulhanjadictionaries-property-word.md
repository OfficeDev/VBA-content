---
title: Global.HangulHanjaDictionaries Property (Word)
keywords: vbawd10.chm163119214
f1_keywords:
- vbawd10.chm163119214
ms.prod: word
api_name:
- Word.Global.HangulHanjaDictionaries
ms.assetid: 46a86461-960b-1ce2-9c86-624cdfd130c9
ms.date: 06/08/2017
---


# Global.HangulHanjaDictionaries Property (Word)

Returns a  **[HangulHanjaConversionDictionaries](hangulhanjaconversiondictionaries-object-word.md)** collection that represents all the active custom conversion dictionaries.


## Syntax

 _expression_ . **HangulHanjaDictionaries**

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

Active custom conversion dictionaries are marked with a check in the  **Custom Dictionaries** dialog box (on the **Tools** menu, click **Options**, then click the  **Spelling &; Grammar** tab, and then click the **Custom Dictionaries** button).

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


[Global Object](global-object-word.md)

