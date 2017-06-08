---
title: OLEObject.GotFocus Event (Excel)
keywords: vbaxl10.chm501073
f1_keywords:
- vbaxl10.chm501073
ms.prod: excel
api_name:
- Excel.OLEObject.GotFocus
ms.assetid: 2bd9a3d8-9305-2354-5ddd-262f4720b444
ms.date: 06/08/2017
---


# OLEObject.GotFocus Event (Excel)

Occurs when an ActiveX control gets input focus.


## Syntax

 _expression_ . **GotFocus**

 _expression_ A variable that represents an **OLEObject** object.


### Return Value

Nothing


## Example

This example runs when ListBox1 gets the focus.


```vb
Private Sub ListBox1_GotFocus() 
 ' runs when list box gets the focus 
End Sub
```


## See also


#### Concepts


[OLEObject Object](oleobject-object-excel.md)

