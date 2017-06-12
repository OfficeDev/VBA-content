---
title: Axes.Parent Property (Excel)
keywords: vbaxl10.chm571075
f1_keywords:
- vbaxl10.chm571075
ms.prod: excel
api_name:
- Excel.Axes.Parent
ms.assetid: d5cd5daf-7579-4df3-8dad-b3daf3e5b5ae
ms.date: 06/08/2017
---


# Axes.Parent Property (Excel)

Returns the parent object for the specified object. Read-only.


## Syntax

 _expression_ . **Parent**

 _expression_ A variable that represents an **Axes** object.


## Example

This example displays the name of the chart that contains  `myAxis`.


```vb
Sub DisplayParentName() 
 
 Set myAxis = Charts(1).Axes(xlValue) 
 MsgBox myAxis.Parent.Name 
 
End Sub
```


## See also


#### Concepts


[Axes Collection](axes-object-excel.md)

