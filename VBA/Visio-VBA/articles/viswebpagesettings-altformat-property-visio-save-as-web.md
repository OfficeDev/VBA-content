---
title: VisWebPageSettings.AltFormat Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.AltFormat
ms.assetid: 60f9af7d-dc5a-d234-976a-51db21473e28
ms.date: 06/08/2017
---


# VisWebPageSettings.AltFormat Property (Visio Save As Web)

Determines whether a secondary output format for the Web page is defined. Read/write.


## Syntax

 _expression_. **AltFormat**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **Long**


## Remarks

The  **AltFormat** property returns non-zero ( **True**) if a secondary output format for the Web page is defined; otherwise, it returns zero ( **False**). The default is  **True**.

Set the  **AltFormat** property to a non-zero value ( **True**) to enable selection of a secondary output format for the web page; otherwise, set it to zero ( **False**).

The  **AltFormat** property is ignored if the primary output format chosen is supported in all browsers by Microsoft Visio 2010. For more information about primary and secondary output formats, see the **[PriFormat](viswebpagesettings-altformat-property-visio-save-as-web.md)** and **[SecFormat](viswebpagesettings-secformat-property-visio-save-as-web.md)** properties.

The following table shows the compatibility of several browsers with various graphic file types and features.



|**Format type**|**Microsoft Internet Explorer 6 or later**|**Microsoft Internet Explorer 5 or earlier**|**Firefox 3 or later**|
|:-----|:-----|:-----|:-----|
|XAML|Yes with plug-in|No|Yes with plug-in|
|VML|Yes|Varies|No|
|SVG|Yes with plug-in|Yes with plug-in|Partial|
|PNG|Yes|Yes|Yes|
|GIF|Yes|Yes|Yes|
|JPEG|Yes|Yes|Yes|
The  **AltFormat** property corresponds to the **Provide alternate format for older browsers** check box on the **Advanced** tab of the **Save as Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, click  **Publish**, and then click  **Advanced**).


