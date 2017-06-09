---
title: GraphicItem.Tag Property (Visio)
keywords: vis_sdr.chm16960435
f1_keywords:
- vis_sdr.chm16960435
ms.prod: visio
api_name:
- Visio.GraphicItem.Tag
ms.assetid: 1f355106-eb71-0bab-cd6b-497bda447ccc
ms.date: 06/08/2017
---


# GraphicItem.Tag Property (Visio)

Gets or sets a user-defined string expression that can store extra data related to your program. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Tag**

 _expression_ An expression that returns a **GraphicItem** object.


### Return Value

String


## Remarks

Microsoft Visio makes no use of the  **Tag** property, nor does its value appear anywhere in the user interface. The **Tag** property is intended to be used by developers to store additional information about a graphic item or a data graphic.

For example, you can assign a  **Tag** property value to a particular graphic item as an identifier. Then, by writing a procedure that iterates through the **GraphicItems** collection and looks for a graphic item that has that specific **Tag** value, you can find the graphic item.


 **Note**  Before you can set any property of a graphic item, you must use the  **[Master.Open](master-open-method-visio.md)** method to open a copy of the data graphic master that contains the graphic item for editing. When you are finished setting properties, use the **Master.Close** method to commit changes.


