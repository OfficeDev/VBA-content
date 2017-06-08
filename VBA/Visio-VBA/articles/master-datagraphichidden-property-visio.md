---
title: Master.DataGraphicHidden Property (Visio)
keywords: vis_sdr.chm10760100
f1_keywords:
- vis_sdr.chm10760100
ms.prod: visio
api_name:
- Visio.Master.DataGraphicHidden
ms.assetid: adcf1867-8541-785b-d8ad-dd44583473b9
ms.date: 06/08/2017
---


# Master.DataGraphicHidden Property (Visio)

Hides or displays a data graphic in the  **Data Graphics** task pane in the Microsoft Visio user interface. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataGraphicHidden**

 _expression_ An expression that returns a **Master** object.


### Return Value

Boolean


## Remarks

Set the  **DataGraphicHidden** property to **True** (-1) to hide a data graphic. Set the property to **False** (0) to display the data graphic. The default is to display the data graphic ( **False** ).

A data graphic is a  **Master** object of type **visTypeDataGraphic** . Before you can set any property of a data graphic master, you must use the **[Master.Open](master-open-method-visio.md)** method to open a copy of the data graphic master for editing. When you are finished setting properties, use the **Master.Close** method to commit changes.


