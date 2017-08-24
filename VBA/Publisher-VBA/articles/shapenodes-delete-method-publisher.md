---
title: ShapeNodes.Delete Method (Publisher)
keywords: vbapb10.chm3473425
f1_keywords:
- vbapb10.chm3473425
ms.prod: publisher
api_name:
- Publisher.ShapeNodes.Delete
ms.assetid: 09f7a8ef-cefd-5a68-f0a6-e99c2f111ea6
ms.date: 06/08/2017
---


# ShapeNodes.Delete Method (Publisher)

Deletes the specified shape node object.


## Syntax

 _expression_. **Delete**( **_Index_**)

 _expression_A variable that represents a  **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **[INT]**| **Long**. The number of the shape node to delete.|

## Example

This example deletes the first node in the first shape in the active publication.


```vb
Sub DeleteNode() 
 ActiveDocument.Pages(1).Shapes(1).Nodes.Delete Index:=1 
End Sub
```


