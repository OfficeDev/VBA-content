---
title: InlineShape.HasChart Property (Word)
keywords: vbawd10.chm162005140
f1_keywords:
- vbawd10.chm162005140
ms.prod: word
api_name:
- Word.InlineShape.HasChart
ms.assetid: f8b88eef-ec41-fc03-f58b-e346d240a121
ms.date: 06/08/2017
---


# InlineShape.HasChart Property (Word)

 **True** if the specified shape is a chart. Read-only.


## Syntax

 _expression_ . **HasChart**

 _expression_ An expression that returns an **InlineShape** object.


## Remarks

This property always returns false for OLE charts. For OLE charts, use  `InlineShape.OLEFormat.ProgID` and check for the following possible values: "Excel.Chart.8", "MSGraph.Chart.8", "Excel.Sheet.8", "Excel.Chart.5", "MSGraph.Chart.5", or "Excel.Sheet.5".


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

