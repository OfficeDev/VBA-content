---
title: VisWebPageSettings.PanAndZoom Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.PanAndZoom
ms.assetid: 83d1ac9d-e489-0656-a573-ebadd6e06156
ms.date: 06/08/2017
---


# VisWebPageSettings.PanAndZoom Property (Visio Save As Web)

Determines whether the  **Pan and Zoom** control for zooming in and out of the page is displayed in a Web page. Read/write.


## Syntax

 _expression_. **PanAndZoom**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **Long**


## Remarks

 **PanAndZoom** returns non-zero ( **True**) if the  **Pan and Zoom** control is displayed after the drawing is exported to a Web page; otherwise, it returns zero ( **False**). The default is  **True**.

Set  **PanAndZoom** to a non-zero value ( **True**) to display the  **Pan and Zoom** control after the drawing is exported to a Web page; otherwise, set it to zero ( **False**).

The  **PanAndZoom** property corresponds to the **Pan and Zoom** check box under **Publishing Options** on the **General** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, and then click  **Publish**).


 **Note**  The  **Pan and Zoom** control is supported for the VML output format in Microsoft Internet Explorer 5 and later. The **Pan and Zoom** control is not available in SVG, JPG, GIF, and PNG output formats.


