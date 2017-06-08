---
title: Window.Panes Property (Word)
keywords: vbawd10.chm157417475
f1_keywords:
- vbawd10.chm157417475
ms.prod: word
api_name:
- Word.Window.Panes
ms.assetid: d75cc2ab-940f-9e2b-81d5-bbbfdb0f4c6c
ms.date: 06/08/2017
---


# Window.Panes Property (Word)

Returns a  **[Panes](panes-object-word.md)** collection that represents all the window panes for the specified window.


## Syntax

 _expression_ . **Panes**

 _expression_ An expression that returns a **[Window](window-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example splits the active window in half.


```vb
If ActiveDocument.ActiveWindow.Panes.Count = 1 Then _ 
 ActiveDocument.ActiveWindow.Panes.Add
```

This example activates the first pane in the window for Document2.




```
Windows("Document2").Panes(1).Activate
```


## See also


#### Concepts


[Window Object](window-object-word.md)

