---
title: Cell.Units Property (Visio)
keywords: vis_sdr.chm10114620
f1_keywords:
- vis_sdr.chm10114620
ms.prod: visio
api_name:
- Visio.Cell.Units
ms.assetid: 075cfda9-8b7a-550b-cf72-b8044c3d461a
ms.date: 06/08/2017
---


# Cell.Units Property (Visio)

Indicates the unit of measure associated with a  **Cell** object. Read-only.


## Syntax

 _expression_ . **Units**

 _expression_ A variable that represents a **Cell** object.


### Return Value

Integer


## Remarks

The  **Units** property can be used to determine the unit of measure currently associated with a cell's value. The various unit codes are declared by the Visio type library in member **[VisUnitCodes ](visunitcodes-enumeration-visio.md)** . For example, a cell's width might be expressed in inches ( **visInches** ) or in centimeters ( **visCentimeters** ). In some cases, a program might behave differently depending on whether a cell's value is in metric or in imperial units.

For a list of valid unit codes, see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).


