---
title: VisWebPageSettings.QuietMode Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.QuietMode
ms.assetid: 1bdc15d9-a4f3-de94-d6ed-4da508d98581
ms.date: 06/08/2017
---


# VisWebPageSettings.QuietMode Property (Visio Save As Web)

Determines whether dialog boxes are displayed in the Visio user interface when you save a drawing as a Web page. Read/write.


## Syntax

 _expression_. **QuietMode**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **Long**


## Remarks

Set  **QuietMode** to a non-zero value ( **True**) to prevent modal dialog boxes from appearing in the user interface when a drawing is saved as a Web page; set it to zero ( **False**) to display dialog boxes with default settings. The default is  **False**.

Setting the  **QuietMode** property to **True** prevents modal dialog boxes from appearing in the Microsoft Visio user interface; however, the **Save As Web Page** progress bar is displayed while the page is being created.

To prevent the user interface from appearing entirely, use the  **[SilentMode](viswebpagesettings-silentmode-property-visio-save-as-web.md)** property.

If both the  **QuietMode** and **SilentMode** properties are set to **True**, the  **SilentMode** property takes precedence and no user interface is displayed.


## Example

The following macro shows how to set the  **QuietMode** property to **True** before saving the drawing as a Web page. Setting this value to **True** prevents modal dialog boxes from appearing in the user interface—only the progress bar appears. Because the **[OpenBrowser](viswebpagesettings-openbrowser-property-visio-save-as-web.md)** property is set to **True**, the drawing opens in the browser.

Before running this macro, replace  _path\filename.htm_ with a valid target path on your computer and the filename that you want to assign to your Web page.




```vb
Public Sub QuietMode_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsowebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .QuietMode = True 
 .OpenBrowser = True 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```


