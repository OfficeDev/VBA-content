---
title: Shapes.AddShape Method (Word)
keywords: vbawd10.chm161415185
f1_keywords:
- vbawd10.chm161415185
ms.prod: word
api_name:
- Word.Shapes.AddShape
ms.assetid: a0f1ce85-a641-5e9f-eb3c-4ebf01fdc32a
ms.date: 06/08/2017
---


# Shapes.AddShape Method (Word)

Adds an AutoShape to a document. Returns a  **[Shape](shape-object-word.md)** object that represents the AutoShape and adds it to the **[Shapes](shapes-object-word.md)** collection.


## Syntax

 _expression_ . **AddShape**( **_Type_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **Long**|The type of shape to be returned. Can be any  **MsoAutoShapeType** constant.|
| _Left_|Required| **Single**|The position, measured in points, of the left edge of the AutoShape.|
| _Top_|Required| **Single**|The position, measured in points, of the top edge of the AutoShape.|
| _Width_|Required| **Single**|The width, measured in points, of the AutoShape.|
| _Height_|Required| **Single**|The height, measured in points, of the AutoShape.|

### Return Value

 **Shape**


## Remarks

To change the type of an AutoShape that you've added, set the  **AutoShapeType** property.


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

