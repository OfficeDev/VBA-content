---
title: Global.Templates Property (Word)
keywords: vbawd10.chm163119171
f1_keywords:
- vbawd10.chm163119171
ms.prod: word
api_name:
- Word.Global.Templates
ms.assetid: 4aa67807-023a-2b52-4773-114d86e340e3
ms.date: 06/08/2017
---


# Global.Templates Property (Word)

Returns a  **Templates** collection that represents all the available templates—global templates and those attached to open documents.


## Syntax

 _expression_ . **Templates**

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the name of each template in the  **Templates** collection.


```vb
Count = 1 
For Each aTemplate In Templates 
 MsgBox aTemplate.Name &; " is template number " &; Count 
 Count = Count + 1 
Next aTemplate
```

In this example, if template one is a global template, its path is stored in  _thePath_ . The **ChDir** statement is used to make the folder with the path stored in _thePath_ the current folder. When this change is made, the **Open** dialog box is displayed.




```vb
If Templates(1).Type = wdGlobalTemplate Then 
 thePath = Templates(1).Path 
 If thePath <> "" Then ChDir thePath 
 Dialogs(wdDialogFileOpen).Show 
End If
```


## See also


#### Concepts


[Global Object](global-object-word.md)

