---
title: Shape.HasChart Property (Word)
keywords: vbawd10.chm161480852
f1_keywords:
- vbawd10.chm161480852
ms.prod: word
api_name:
- Word.Shape.HasChart
ms.assetid: 5fd4bc0b-153a-f30b-dd81-81a4b348770c
ms.date: 06/08/2017
---


# Shape.HasChart Property (Word)

 **True** if the specified shape has a chart. Read-only.


## Syntax

 _expression_ . **HasChart**

 _expression_ An expression that returns a **Shape** object.


## Remarks

This property always returns false for OLE charts. For OLE charts, use  `InlineShape.OLEFormat.ProgID` and check for the following possible values: "Excel.Chart.8", "MSGraph.Chart.8", "Excel.Sheet.8", "Excel.Chart.5", "MSGraph.Chart.5", or "Excel.Sheet.5".


## See also


#### Concepts


[Shape Object](shape-object-word.md)

