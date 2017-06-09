---
title: Shape.RowsCellCount Property (Visio)
keywords: vis_sdr.chm11214265
f1_keywords:
- vis_sdr.chm11214265
ms.prod: visio
api_name:
- Visio.Shape.RowsCellCount
ms.assetid: bb9c1990-5ead-e56b-7b09-a49a2b7ad111
ms.date: 06/08/2017
---


# Shape.RowsCellCount Property (Visio)

Returns the number of cells in a row of a ShapeSheet section. Read-only.


## Syntax

 _expression_ . **RowsCellCount**( **_Section_** , **_Row_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The index of the section that contains the row.|
| _Row_|Required| **Integer**|The index of the row to count.|

### Return Value

Integer


## Remarks

Use section and row index constants declared by the Visio type library in members  **[VisSectionIndices](vissectionindices-enumeration-visio.md)** and **[VisRowIndices](visrowindices-enumeration-visio.md)** .


