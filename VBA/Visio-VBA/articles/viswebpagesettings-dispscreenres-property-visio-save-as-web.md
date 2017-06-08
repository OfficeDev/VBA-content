---
title: VisWebPageSettings.DispScreenRes Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.DispScreenRes
ms.assetid: ec62976a-4a92-f904-b7de-1e9470dc5411
ms.date: 06/08/2017
---


# VisWebPageSettings.DispScreenRes Property (Visio Save As Web)

Specifies the screen resolution for your Web page. Read/write.


## Syntax

 _expression_. **DispScreenRes**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

VISWEB_DISP_RES


## Remarks

The default screen resolution is 1024 x 768 pixels.

The  **DispScreenRes** property affects only how raster-based formats (for example, GIF) fit on the page. For vector-based images (for example, VML), this property is ignored.

The  **DispScreenRes** property value corresponds to the value in the **Target monitor** box on the **Advanced** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, click  **Publish**, and then click  **Advanced**).


