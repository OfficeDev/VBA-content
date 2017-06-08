---
title: VisWebPageSettings.SecFormat Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.SecFormat
ms.assetid: 2c6fa96d-8a71-28fb-c8d7-f7ba6772fe43
ms.date: 06/08/2017
---


# VisWebPageSettings.SecFormat Property (Visio Save As Web)

Specifies the secondary output format for the Web page. Read/write.


## Syntax

 _expression_. **SecFormat**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **String**


## Remarks

The secondary output format is used if the browser does not support the primary output format. For example, if the primary format is XAML and the required Silverlight browser plug-in is not installed, the Web page output uses the secondary format.

The primary output format is specified by the  **[PriFormat](viswebpagesettings-altformat-property-visio-save-as-web.md)** property. For information about which browsers are compatible with selected formats, see the **[AltFormat](viswebpagesettings-altformat-property-visio-save-as-web.md)** property.

Possible values for the  **SecFormat** property are as follows:


- PNG (Portable Network Graphics), the default
    
- JPG (JPEG File Interchage Format)
    
- GIF (Graphics Interchange Format)
    
This value corresponds to the value selected in the list below the  **Provide alternate format for older browsers** check box (if it is selected) on the **Advanced** tab of the **Save as Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, click  **Publish**, and then click  **Advanced**).


## Example

The following macro shows how to use the  **SecFormat** property to set the secondary format value to JPG for browsers that do not support the primary format of XAML (the default).

Before running this macro, replace  _path\filename.htm_ with a valid target path on your computer and the filename that you want to assign to your Web page.




```vb
Public Sub SecFormat_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsowebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .AltFormat = True 
 .SecFormat = "JPG" 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```


