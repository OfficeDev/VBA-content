---
title: Shape.AddGuide Method (Visio)
keywords: vis_sdr.chm11216035
f1_keywords:
- vis_sdr.chm11216035
ms.prod: visio
api_name:
- Visio.Shape.AddGuide
ms.assetid: 1155354e-3855-4def-bafb-0d70c933a57a
ms.date: 06/08/2017
---


# Shape.AddGuide Method (Visio)

Adds a guide to a group shape.


## Syntax

 _expression_ . **AddGuide**( **_Type_** , **_xPos_** , **_yPos_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **Integer**|The type of guide to add.|
| _xPos_|Required| **Double**|The x-coordinate of a point on the guide.|
| _yPos_|Required| **Double**|The y-coordinate of a point on the guide.|

### Return Value

Shape


## Remarks

You can add a guide only to a group shape.

To view guides that you add to a group shape by using the  **AddGuide** method, use the **OpenDrawWindow** method to open the **Group Editing** window.

The following constants declared by the Visio type library are valid values for guide types.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visPoint**|1|Guide point|
| **visHorz**|2|Horizontal guide|
| **visVert**|3|Vertical guide|

