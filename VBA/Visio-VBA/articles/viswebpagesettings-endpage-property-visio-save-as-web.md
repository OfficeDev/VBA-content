---
title: VisWebPageSettings.EndPage Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.EndPage
ms.assetid: 4b7ebf2d-b814-8588-b25e-7c54fd0affda
ms.date: 06/08/2017
---


# VisWebPageSettings.EndPage Property (Visio Save As Web)

Specifies the page number of the last page in the range when you save a range of pages as a Web page. Read/write.


## Syntax

 _expression_. **EndPage**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **Long**


## Remarks

The start page number is specified in the  **[StartPage](viswebpagesettings-startpage-property-visio-save-as-web.md)** property.

The  **EndPage** property value corresponds to the value in the **to** box on the **General** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, and then click  **Publish**).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **EndPage** property to save a range of pages in a drawing (in this case, from page 2 to page 3) as a Web page instead of the complete drawing.

This macro assumes that the current Visio drawing contains at least three pages.

Before running this macro, replace  _path\filename.htm_ with a valid target path on your computer and the filename that you want to assign to your Web page.




```vb
Public Sub EndPage_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .StartPage = 2 
 .EndPage = 3 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```


