---
title: OLEObject.LostFocus Event (Excel)
keywords: vbaxl10.chm501074
f1_keywords:
- vbaxl10.chm501074
ms.prod: excel
api_name:
- Excel.OLEObject.LostFocus
ms.assetid: 9d8004be-97f5-54d2-3826-210f7cf0569f
ms.date: 06/08/2017
---


# OLEObject.LostFocus Event (Excel)

Occurs when an ActiveX control loses input focus.


## Syntax

 _expression_ . **LostFocus**

 _expression_ A variable that represents an **OLEObject** object.


### Return Value

Nothing


## Example

This example runs when ListBox1 loses the focus.


```vb
Private Sub ListBox1_LostFocus() 
 ' runs when list box loses the focus 
End Sub
```


## See also


#### Concepts


[OLEObject Object](oleobject-object-excel.md)

