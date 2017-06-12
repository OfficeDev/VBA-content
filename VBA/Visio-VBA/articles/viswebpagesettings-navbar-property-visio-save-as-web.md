---
title: VisWebPageSettings.NavBar Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.NavBar
ms.assetid: 5a3245df-d0b6-40c6-5ed9-6d7700e835c8
ms.date: 06/08/2017
---


# VisWebPageSettings.NavBar Property (Visio Save As Web)

Determines whether the  **Go to Page** navigation control is displayed in a Web page. Read/write.


## Syntax

 _expression_. **NavBar**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **Long**


## Remarks

The  **NavBar** property returns non-zero ( **True**) if the  **Go to Page** navigation control is displayed after the drawing is exported to a Web page; otherwise, it returns zero ( **False**). The default is  **True**.

Set  **NavBar** to a non-zero value ( **True**) to display the  **Go to Page** navigation control after the drawing is exported to a Web page; otherwise, set it to zero ( **False**).

This property corresponds to the  **Go to Page (navigation control)** check box under **Publishing Options** on the **General** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, and then click  **Publish**).

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio.SaveAsWeb** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.SaveAsWeb.IVisWebPageSettings.NavBar**
    

