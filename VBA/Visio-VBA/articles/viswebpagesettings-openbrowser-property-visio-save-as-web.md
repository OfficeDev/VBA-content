---
title: VisWebPageSettings.OpenBrowser Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.OpenBrowser
ms.assetid: 701defdf-9f1c-b136-0af5-48605d255f88
ms.date: 06/08/2017
---


# VisWebPageSettings.OpenBrowser Property (Visio Save As Web)

Determines whether the Web page opens in the browser after the drawing is exported to a Web page. Read/write.


## Syntax

 _expression_. **OpenBrowser**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **Long**


## Remarks

 **OpenBrowser** returns non-zero ( **True**) if the Web page opens in the browser after the drawing is exported to a Web page; otherwise, it returns zero ( **False**). The default is  **True**.

Set  **OpenBrowser** to a non-zero value ( **True**) to open a Web page in the browser after the drawing is exported to a Web page; otherwise, set it to zero ( **False**).

The  **OpenBrowser** property corresponds to the **Automatically open Web page in browser** check box on the **General** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, and then click  **Publish**).


