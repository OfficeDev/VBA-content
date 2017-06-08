---
title: Master.DataGraphicHidesText Property (Visio)
keywords: vis_sdr.chm10760115
f1_keywords:
- vis_sdr.chm10760115
ms.prod: visio
api_name:
- Visio.Master.DataGraphicHidesText
ms.assetid: c1a08780-0873-3d8b-1872-edc8a6515840
ms.date: 06/08/2017
---


# Master.DataGraphicHidesText Property (Visio)

Displays or hides the text of a shape or of the primary shape in a selection when a data graphic is applied to the shape or to the selection. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataGraphicHidesText**

 _expression_ An expression that returns a **Master** object.


### Return Value

Boolean


## Remarks

Setting the  **DataGraphicHidesText** property to **False** (0) displays the shape text. Setting the property to **True** (-1) hides the text. The default is to display the text ( **False** ).

The value of the  **DataGraphicHidesText** property corresponds to the setting of the **HIde shape text when data graphic is applied** box under **Display options** in the **New Data Graphic** or **Edit Data Graphic** dialog box.

A data graphic is a  **Master** object of type **visTypeDataGraphic** . Before you can set any property of a data graphic master, you must use the **[Master.Open](master-open-method-visio.md)** method to open a copy of the data graphic master for editing. When you are finished setting properties, use the **Master.Close** method to commit changes.


