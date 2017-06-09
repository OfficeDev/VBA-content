---
title: OLEFormat.ObjectVerbs Property (Publisher)
keywords: vbapb10.chm4456453
f1_keywords:
- vbapb10.chm4456453
ms.prod: publisher
api_name:
- Publisher.OLEFormat.ObjectVerbs
ms.assetid: 887070e6-7f7d-4f65-290e-3d46bfd91d34
ms.date: 06/08/2017
---


# OLEFormat.ObjectVerbs Property (Publisher)

Returns an  **[ObjectVerbs](objectverbs-object-publisher.md)** collection that contains all the OLE verbs for the specified OLE object. Read-only.


## Syntax

 _expression_. **ObjectVerbs**

 _expression_A variable that represents an  **OLEFormat** object.


### Return Value

ObjectVerbs


## Example

This example displays all the available verbs for the OLE object contained in shape one on page two in the active publication. For this example to work, shape one must be a shape that represents an OLE object.


```vb
Dim v As String 
 
With ActiveDocument.Pages(2).Shapes(1).OLEFormat 
 For Each v In .ObjectVerbs 
 MsgBox v 
 Next 
End With
```


