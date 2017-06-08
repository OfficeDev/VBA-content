---
title: GraphicItem.Type Property (Visio)
keywords: vis_sdr.chm16914595
f1_keywords:
- vis_sdr.chm16914595
ms.prod: visio
api_name:
- Visio.GraphicItem.Type
ms.assetid: 36af507e-270b-e2e6-97b9-c5e02ffe1b96
ms.date: 06/08/2017
---


# GraphicItem.Type Property (Visio)

Returns the type of the graphic item. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Type**

 _expression_ A variable that represents a **GraphicItem** object.


### Return Value

VisGraphicItemTypes


## Remarks

The following possible values for the  **Type** property are from the **VisGraphicItemTypes** enumeration. which is declared in the Visio type library. These values correspond to the graphic item types listed in the **New Item** list in the **New Data Graphic** and **Edit Data Graphic** dialog boxes in the Microsoft Visio user interface.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visTypeIconSet**|2|Represents an  **Icon Set** graphic item.|
| **visTypeTextCallout**|3|Represents a  **Text** graphic item.|
| **visTypeDataBar**|4|Represents a  **Data Bar** graphic item.|
| **visTypeColorByValue**|5|Represents a  **Color by Value** graphic item.|
| **visTypeHeading**|6|Represents a  **Text** graphic item that has a **Callout** type of **Heading x**.|

