---
title: Shape.Connector Property (Excel)
keywords: vbaxl10.chm636094
f1_keywords:
- vbaxl10.chm636094
ms.prod: excel
api_name:
- Excel.Shape.Connector
ms.assetid: 757505bd-4c45-9d54-a5ac-94e251b351be
ms.date: 06/08/2017
---


# Shape.Connector Property (Excel)

 **True** if the specified shape is a connector. Read-only **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **Connector**

 _expression_ An expression that returns a **Shape** object.


## Example

This example deletes all connectors on  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
    For i = .Count To 1 Step -1 
        With .Item(i) 
            If .Connector Then .Delete 
        End With 
    Next 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

