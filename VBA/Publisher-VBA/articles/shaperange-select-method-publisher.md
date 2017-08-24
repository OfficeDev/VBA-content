---
title: ShapeRange.Select Method (Publisher)
keywords: vbapb10.chm2293799
f1_keywords:
- vbapb10.chm2293799
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Select
ms.assetid: 3252ba74-d051-8c28-a9ed-c6f5ca711dec
ms.date: 06/08/2017
---


# ShapeRange.Select Method (Publisher)

Selects the specified object.


## Syntax

 _expression_. **Select**( **_Replace_**)

 _expression_A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Replace|Optional| **Variant**|Specifies whether the selection replaces any previous selection.  **True** to replace the previous selection with the new selection; **False** to add the new selection to the previous selection. Default is **True**.|

## Example

This example selects shapes one and three on page one in the active publication.


```vb
ActiveDocument.Pages(1).Shapes.Range(Array(1, 3)).Select
```

This example adds shapes two and four on page one in the active publication to the previous selection.




```vb
ActiveDocument.Pages(1).Shapes.Range(Array(2, 4)) _ 
 .Select Replace:=False
```


