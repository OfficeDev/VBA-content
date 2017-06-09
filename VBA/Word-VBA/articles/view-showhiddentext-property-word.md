---
title: View.ShowHiddenText Property (Word)
keywords: vbawd10.chm161808403
f1_keywords:
- vbawd10.chm161808403
ms.prod: word
api_name:
- Word.View.ShowHiddenText
ms.assetid: e4f58049-1fb9-5c70-0786-5f4c8c54f3ba
ms.date: 06/08/2017
---


# View.ShowHiddenText Property (Word)

 **True** if text formatted as hidden text is displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowHiddenText**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example hides text formatted as hidden text in each window.


```vb
For Each myWindow In Windows 
 myWindow.View.ShowHiddenText = False 
Next myWindow
```

This example toggles the display of hidden text.




```vb
ActiveDocument.ActiveWindow.View.ShowHiddenText = _ 
 Not ActiveDocument.ActiveWindow.View.ShowHiddenText
```


## See also


#### Concepts


[View Object](view-object-word.md)

