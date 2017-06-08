---
title: View.TableGridlines Property (Word)
keywords: vbawd10.chm161808409
f1_keywords:
- vbawd10.chm161808409
ms.prod: word
api_name:
- Word.View.TableGridlines
ms.assetid: 02ef1d7b-185b-ed17-e811-a752faa11b3f
ms.date: 06/08/2017
---


# View.TableGridlines Property (Word)

 **True** if table gridlines are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **TableGridlines**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example displays table gridlines in the active window.


```vb
ActiveDocument.ActiveWindow.View.TableGridlines = True
```

This example shows table gridlines for the panes associated with window one in the Windows collection.




```vb
For Each myPane In Windows(1).Panes 
 myPane.View.TableGridlines = True 
Next myPane
```


## See also


#### Concepts


[View Object](view-object-word.md)

