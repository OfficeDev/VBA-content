---
title: Cell.InheritedFormulaSource Property (Visio)
keywords: vis_sdr.chm10150680
f1_keywords:
- vis_sdr.chm10150680
ms.prod: visio
api_name:
- Visio.Cell.InheritedFormulaSource
ms.assetid: 62aedef3-06b1-2fc3-5fd2-03f77668548f
ms.date: 06/08/2017
---


# Cell.InheritedFormulaSource Property (Visio)

Returns the cell from which this cell inherited its formula. Read-only.


## Syntax

 _expression_ . **InheritedFormulaSource**

 _expression_ A variable that represents a **Cell** object.


### Return Value

Cell


## Remarks

If the formula in this cell is a local formula, the  **InheritedFormulaSource** property returns itself.


