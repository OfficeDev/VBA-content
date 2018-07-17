---
title: Application.Dialogs Property (Word)
keywords: vbawd10.chm158334995
f1_keywords:
- vbawd10.chm158334995
ms.prod: word
api_name:
- Word.Application.Dialogs
ms.assetid: 17acdfab-32d2-ddb8-04aa-692f9ffb20b8
ms.date: 06/08/2017
---


# Application.Dialogs Property (Word)

Returns a  **[Dialogs](dialogs-object-word.md)** collection that represents all the built-in dialog boxes in Word. Read-only.


## Syntax

 _expression_ . **Dialogs**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](returning-an-object-from-a-collection-word.md). 

For a list of built-in dialog boxes, see the **[WdWordDialog](wdworddialog-enumeration-word.md)** enumeration.


## Example

This example displays the built-in  **Find** dialog box, with "Hello" in the **Find What** box.


```vb
Dim dlgFind As Dialog 
 
Set dlgFind = Dialogs(wdDialogEditFind) 
 
With dlgFind 
 .Find = "Hello" 
 .Show 
End With
```

This example displays the built-in  **Open** dialog box showing all file types.




```vb
With Dialogs(wdDialogFileOpen) 
 .Name = "*.*" 
 .Show 
End With
```

This example prints the active document, using the settings from the  **Print** dialog box.




```
Dialogs(wdDialogFilePrint).Execute
```


## See also


#### Concepts


[Application Object](application-object-word.md)

