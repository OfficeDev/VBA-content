---
title: RulerGuides.Add Method (Publisher)
keywords: vbapb10.chm720900
f1_keywords:
- vbapb10.chm720900
ms.prod: publisher
api_name:
- Publisher.RulerGuides.Add
ms.assetid: 3986452a-73da-04c2-4e11-8369d61cd974
ms.date: 06/08/2017
---


# RulerGuides.Add Method (Publisher)

Adds a new ruler guide to the specified  **RulerGuides** collection.


## Syntax

 _expression_. **Add**( **_Position_**,  **_Type_**)

 _expression_A variable that represents a  **RulerGuides** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Position|Required| **Variant**|The position relative to the left edge or top edge of the page where the new ruler guide will be added. Numeric values are evaluated in points; strings are evaluated in the units specified and can be in any measurement unit supported by Microsoft Publisher (for example, "2.5 in").|
|Type|Required| **PbRulerGuideType**|The type of ruler guide to add.|

## Remarks

Type can be one of these  **PbRulerGuideType** constants.



| **pbRulerGuideTypeHorizontal**|
| **pbRulerGuideTypeVertical**|

## Example

The following example adds ruler guides to page one that are 0.5 inches from the left and top edges of the page.


```vb
With ActiveDocument.Pages(1).RulerGuides 
 .Add Position:="0.5 in", Type:=pbRulerGuideTypeHorizontal 
 .Add Position:="0.5 in", Type:=pbRulerGuideTypeVertical 
End With
```


