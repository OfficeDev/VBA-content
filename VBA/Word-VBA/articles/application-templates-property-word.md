---
title: Application.Templates Property (Word)
keywords: vbawd10.chm158335043
f1_keywords:
- vbawd10.chm158335043
ms.prod: word
api_name:
- Word.Application.Templates
ms.assetid: 816e50d1-32b9-c8ff-6d2c-ad1113c952fc
ms.date: 06/08/2017
---


# Application.Templates Property (Word)

Returns a  **[Templates](templates-object-word.md)** collection that represents all the available templates—global templates and those attached to open documents.


## Syntax

 _expression_ . **Templates**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the name of each template in the Templates collection.


```vb
Count = 1 
For Each aTemplate In Templates 
 MsgBox aTemplate.Name &; " is template number " &; Count 
 Count = Count + 1 
Next aTemplate
```

In this example, if template one is a global template, its path is stored in  `thePath`. The  **ChDir** statement is used to make the folder with the path stored in `thePath` the current folder. When this change is made, the **Open** dialog box is displayed.




```vb
If Templates(1).Type = wdGlobalTemplate Then 
 thePath = Templates(1).Path 
 If thePath <> "" Then ChDir thePath 
 Dialogs(wdDialogFileOpen).Show 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

