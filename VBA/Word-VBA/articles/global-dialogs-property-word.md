---
title: Global.Dialogs Property (Word)
keywords: vbawd10.chm163119123
f1_keywords:
- vbawd10.chm163119123
ms.prod: word
api_name:
- Word.Global.Dialogs
ms.assetid: 7eea3680-b232-c18a-d99a-d7c2a5b29cd4
ms.date: 06/08/2017
---


# Global.Dialogs Property (Word)

Returns a  **[Dialogs](dialogs-object-word.md)** collection that represents all the built-in dialog boxes in Word. Read-only.


## Syntax

 _expression_ . **Dialogs**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the built-in Find dialog box, with "Hello" in the Find What box.


```vb
Dim dlgFind As Dialog 
 
Set dlgFind = Dialogs(wdDialogEditFind) 
 
With dlgFind 
 .Find = "Hello" 
 .Show 
End With
```

This example displays the built-in Open dialog box showing all file types.




```vb
With Dialogs(wdDialogFileOpen) 
 .Name = "*.*" 
 .Show 
End With
```

This example prints the active document, using the settings from the Print dialog box.




```
Dialogs(wdDialogFilePrint).Execute
```


## See also


#### Concepts


[Global Object](global-object-word.md)

