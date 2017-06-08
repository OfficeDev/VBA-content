---
title: VisWebPageSettings.Search Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.Search
ms.assetid: ae7e09e6-7f54-e939-5e5c-12af35c1b303
ms.date: 06/08/2017
---


# VisWebPageSettings.Search Property (Visio Save As Web)

Determines whether the  **Search Pages** control for searching for shapes in a drawing is displayed in a Web page. Read/write.


## Syntax

 _expression_. **Search**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **Long**


## Remarks

Search returns non-zero ( **True**) if the  **Search Pages** control is displayed after the drawing is exported to a Web page; otherwise, it returns zero ( **False**). The default is  **True**.

Set  **Search** to a non-zero value ( **True**) to display the  **Search Pages** control after the drawing is exported to a Web page; otherwise, set it to zero ( **False**). 

This property corresponds to the  **Search Pages** check box under **Publishing Options** on the **General** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, and then click  **Publish**).

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio.SaveAsWeb** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.SaveAsWeb.IVisWebPageSettings.Search**
    

