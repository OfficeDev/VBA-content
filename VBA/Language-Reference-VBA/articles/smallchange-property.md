---
title: SmallChange Property
keywords: fm20.chm5225094
f1_keywords:
- fm20.chm5225094
ms.prod: office
api_name:
- Office.SmallChange
ms.assetid: ebe0c130-8c96-77f2-709e-32f8b6d720b5
ms.date: 06/08/2017
---


# SmallChange Property



Specifies the amount of movement that occurs when the user clicks either scroll arrow in a  **ScrollBar** or **SpinButton**.
 **Syntax**
 _object_. **SmallChange** [= _Long_ ]
The  **SmallChange** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. An integer that specifies the amount of change to the  **Value** property.|
 **Remarks**
The  **SmallChange** property does not have units.
Any integer is an acceptable setting for this property. The recommended range of values is from -32,767 to +32,767. The default value is 1.

