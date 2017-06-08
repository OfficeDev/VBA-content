---
title: VisSaveAsWeb.WebPageSettings Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.WebPageSettings
ms.assetid: a026cbcb-1156-89f9-429a-3d1b23c78065
ms.date: 06/08/2017
---


# VisSaveAsWeb.WebPageSettings Property (Visio Save As Web)

Returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object that represents the Web page settings that will be used for the Save as Web Page operation. Read-only.


## Syntax

 _expression_. **WebPageSettings** **VisWebPageSettings**

 _expression_An expression that returns a  ** [VisSaveAsWeb](http://msdn.microsoft.com/library/c4675de8-0f63-179f-f687-8962d54d6b2f%28Office.15%29.aspx)** object.


## Remarks

Use the  **WebPageSettings** property to get a **VisWebPageSettings** object. You can then use the **VisWebPageSettings** object to get and set the properties of your Web page.


## Example

This example shows the simplest way to create a Web page. Because no properties of the  **VisWebPageSettings** object are set (except the **[TargetPath](viswebpagesettings-targetpath-property-visio-save-as-web.md)** property, which is required), all the default settings apply, and the active document is saved.

Before running this macro, replace  _path_ with a valid target path on your computer and replace _filename.htm_ with the file name that you want to assign to your Web page.




```vb
Public Sub WebPageSettings_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsowebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 vsoWebSettings.TargetPath = "path\filename.htm" 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```


