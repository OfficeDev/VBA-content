---
title: Designs.Add Method (PowerPoint)
keywords: vbapp10.chm643004
f1_keywords:
- vbapp10.chm643004
ms.prod: powerpoint
api_name:
- PowerPoint.Designs.Add
ms.assetid: 00608390-a12b-d698-36a6-ded2df3cc26a
ms.date: 06/08/2017
---


# Designs.Add Method (PowerPoint)

Returns a  **[Design](design-object-powerpoint.md)** object that represents a new slide design.


## Syntax

 _expression_. **Add**( **_designName_**, **_Index_** )

 _expression_ A variable that represents a **Designs** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _designName_|Required|**String**|The name of the design.|
| _Index_|Optional|**Integer**|The index number of the design in the  **Designs** collection. The default value is -1, which means that if you omit the Index parameter, the new slide design is added at the end of existing slide designs.|

### Return Value

Design


## See also


#### Concepts


[Designs Object](designs-object-powerpoint.md)

