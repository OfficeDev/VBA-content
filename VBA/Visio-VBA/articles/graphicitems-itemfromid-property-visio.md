---
title: GraphicItems.ItemFromID Property (Visio)
keywords: vis_sdr.chm16813775
f1_keywords:
- vis_sdr.chm16813775
ms.prod: visio
api_name:
- Visio.GraphicItems.ItemFromID
ms.assetid: 2d74816f-b667-25f7-7647-ae14e4b8fcad
ms.date: 06/08/2017
---


# GraphicItems.ItemFromID Property (Visio)

Returns a  **GraphicItem** object from the **GraphicItems** collection by using the unique ID of the object. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **ItemFromID**( **_ObjectID_** )

 _expression_ A variable that represents a **GraphicItems** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectID_|Required| **Long**|The unique ID of the  **GraphicItem** object to retrieve.|

### Return Value

GraphicItem


## Remarks

You can get the ID of a  **GraphicItem** object by getting the value of the **GraphicItem.ID** property.


