---
title: Shape.Connector Property (Publisher)
keywords: vbapb10.chm2228277
f1_keywords:
- vbapb10.chm2228277
ms.prod: publisher
api_name:
- Publisher.Shape.Connector
ms.assetid: 6cdff1e7-59b0-9905-96f8-99b79db1acd5
ms.date: 06/08/2017
---


# Shape.Connector Property (Publisher)

Returns an  **MsoTriState**value indicating whether the specified shape is a connector. Read-only.


## Syntax

 _expression_. **Connector**

 _expression_A variable that represents a  **Shape** object.


## Remarks

The  **Connector** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The shape is not a connector.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified shape.|
| **msoTrue**|The shape is a connector.|

## Example

This example deletes all connectors on page one of the active publication.


```vb
Dim i As Integer 
 
With ActiveDocument.Pages(1).Shapes 
 For i = .Count To 1 Step -1 
 With .Item(i) 
 If .Connector Then .Delete 
 End With 
 Next 
End With
```


