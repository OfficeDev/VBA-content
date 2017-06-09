---
title: Layer.CellsC Property (Visio)
keywords: vis_sdr.chm11813195
f1_keywords:
- vis_sdr.chm11813195
ms.prod: visio
api_name:
- Visio.Layer.CellsC
ms.assetid: da7de08d-e492-a74d-a5de-139a32798deb
ms.date: 06/08/2017
---


# Layer.CellsC Property (Visio)

Returns a  **Cell** object that represents a ShapeSheet cell in a row in the Layers section. Read-only.


## Syntax

 _expression_ . **CellsC**( **_Column_** )

 _expression_ An expression that returns a **Layer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Column_|Required| **Integer**|The cell index of the cell to get.|

### Return Value

Cell


## Remarks

The following constants for the cell index are declared by the Microsoft Visio type library in ** VisCellIndices** .



|**Constant **|**Value **|
|:-----|:-----|
| **visLayerName**|0 |
| **visLayerColor**|2 |
| **visLayerStatus**|3 |
| **visLayerVisible**|4 |
| **visLayerPrint**|5 |
| **visLayerActive**|6 |
| **visLayerLock**|7 |
| **visLayerSnap**|8 |
| **visLayerGlue**|9 |
| **visLayerNameUniv**|10 |
| **visLayerColorTrans**|11|

