---
title: Application.AddIns Property (Word)
keywords: vbawd10.chm158334998
f1_keywords:
- vbawd10.chm158334998
ms.prod: word
api_name:
- Word.Application.AddIns
ms.assetid: 8e464524-1304-6a4a-f2f0-5f652d5c8123
ms.date: 06/08/2017
---


# Application.AddIns Property (Word)

Returns an  **[AddIns](addins-object-word.md)** collection that represents all available add-ins, regardless of whether they're currently loaded. Read-only.


## Syntax

 _expression_ . **AddIns**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

The  **AddIns** collection includes the global templates and Word add-in libraries (WLLs) listed in the **Templates and Add-ins** dialog box. For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example returns the total number of add-ins.


```vb
Dim intAddIns as Integer 
 
intAddIns = AddIns.Count
```

This example displays the name of each add-in in the Addins collection.




```vb
Dim addinLoop as AddIn 
 
For Each addinLoop In AddIns 
 MsgBox addinLoop.Name 
Next addinLoop
```


## See also


#### Concepts


[Application Object](application-object-word.md)

