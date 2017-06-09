---
title: Global.NewWindow Method (Word)
keywords: vbawd10.chm163119449
f1_keywords:
- vbawd10.chm163119449
ms.prod: word
api_name:
- Word.Global.NewWindow
ms.assetid: bf84590f-3a09-1f4f-3957-70a8af99686a
ms.date: 06/08/2017
---


# Global.NewWindow Method (Word)

Opens a new window with the same document as the specified window. Returns a  **Window** object.


## Syntax

 _expression_ . **NewWindow**

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


### Return Value

Window


## Remarks

If the  **NewWindow** method is used with the **Application** object, a new window is opened for the active window. The following two instructions are functionally equivalent.


```vb
Set myWindow = ActiveDocument.ActiveWindow.NewWindow 
Set myWindow = NewWindow
```


## Example

This example posts a message that indicates the number of windows that exist before and after you open a new window for Document1.


```vb
MsgBox Windows.Count &; " windows open" 
Windows("Document1").NewWindow 
MsgBox Windows.Count &; " windows open"
```

This example opens a new window, arranges all the open windows, closes the new window, and then rearranges the open windows.




```vb
Set myWindow = NewWindow 
Windows.Arrange ArrangeStyle:=wdTiled 
myWindow.Close 
Windows.Arrange ArrangeStyle:=wdTiled
```


## See also


#### Concepts


[Global Object](global-object-word.md)

