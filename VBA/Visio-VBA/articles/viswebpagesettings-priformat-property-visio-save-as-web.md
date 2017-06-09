---
title: VisWebPageSettings.PriFormat Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.PriFormat
ms.assetid: 84c7c085-0f12-f25d-bf17-646cc8b7cd97
ms.date: 06/08/2017
---


# VisWebPageSettings.PriFormat Property (Visio Save As Web)

Specifies the primary output format for the Web page. Read/write.


## Syntax

 _expression_. **PriFormat**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **String**


## Remarks

If you select a primary output format that is not supported by all browsers, you should also select a secondary output format for older browsers. To do this, see the  **[SecFormat](viswebpagesettings-secformat-property-visio-save-as-web.md)** property.

For information about which browsers are compatible with selected formats, see the  **[AltFormat](viswebpagesettings-altformat-property-visio-save-as-web.md)** property.

Possible values for the  **PriFormat** property are as follows:


- XAML (Extensible Application Markup Language), the default
    
- SVG (Scalable Vector Graphics)
    
- JPG (JPEG File Interchage Format)
    
- GIF (Graphics Interchange Format)
    
- PNG (Portable Network Graphics)
    
- VML (Vector Markup Language)
    
This value corresponds to the value selected in the  **Output formats** list on the **Advanced** tab of the **Save as Web Page** dialog box (click the **BackstageButton** tab, click **Save As** , in the **Save as type** list, select **Web Page (*.htm;*.html)** , click **Publish** , and then click **Advanced** ).


## Example

The following macro shows how use the  **PriFormat** property to set the primary output format for the Web page to JPG.

Before running this macro, replace  _path\filename.htm_ with a valid target path on your computer and the filename that you want to assign to your Web page.




```vb
Public Sub PriFormat_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsowebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .PriFormat = "JPG" 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```


