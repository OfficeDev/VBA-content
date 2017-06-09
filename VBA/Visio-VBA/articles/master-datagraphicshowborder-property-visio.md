---
title: Master.DataGraphicShowBorder Property (Visio)
keywords: vis_sdr.chm10760120
f1_keywords:
- vis_sdr.chm10760120
ms.prod: visio
api_name:
- Visio.Master.DataGraphicShowBorder
ms.assetid: 203d631c-d838-ea0a-f67a-39de513e738e
ms.date: 06/08/2017
---


# Master.DataGraphicShowBorder Property (Visio)

Gets or sets whether a border is displayed around the graphic items contained in the data graphic that are in default positions. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataGraphicShowBorder**

 _expression_ An expression that returns a **Master** object.


### Return Value

Boolean


## Remarks

Setting the  **DataGraphicShowBorder** property to **False** (0) hides the border around graphic items contained in the data graphic that are at the default position. Setting the property to **True** (-1) displays the border. The default is to hide the border ( **False** ).

The value of the  **DataGraphicShowBorder** property corresponds to the setting of the **Show border around items at default position** box under **Display options** in the **New Data Graphic** or **Edit Data Graphic** dialog box.

A data graphic is a  **Master** object of type **visTypeDataGraphic** . Before you can set any property of a data graphic master, you must use the **[Master.Open](master-open-method-visio.md)** method to open a copy of the data graphic master for editing. When you are finished setting properties, use the **Master.Close** method to commit changes.


