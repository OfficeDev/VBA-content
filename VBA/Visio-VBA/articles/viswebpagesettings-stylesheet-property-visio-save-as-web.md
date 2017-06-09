---
title: VisWebPageSettings.Stylesheet Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.Stylesheet
ms.assetid: 9b837460-83a6-71f8-b63f-3f251dedc87c
ms.date: 06/08/2017
---


# VisWebPageSettings.Stylesheet Property (Visio Save As Web)

Specifies a cascading stylesheet (CSS) provided by Microsoft Visio, or one that you have created, that is applied to the Web page. Read/write.


## Syntax

 _expression_. **Stylesheet**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **String**


## Remarks

A stylesheet can be one provided by Visio or one that you create yourself. If you store a stylesheet that you create in the following folder, it will appear in the  **Style sheet** drop-down list on the **Advanced** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, click  **Publish**, and then click  **Advanced**): 

\ _your_Visio_path_\ _your_language_ID_\

Visio identifies stylesheets by searching through the folder named for your language ID (for example, 1033) for CSS files.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Stylesheet** property to assign the "Steel" stylesheet (supplied by Visio) to the Web page you are creating.

Before running this macro, replace  _path\filename.htm_ with a valid target path on your computer and the file name that you want to assign to your Web page. Also, replace _your_Visio_path_ and _your_language_ID_ with the path to Visio stylesheets on your computer, for example:

C:\Program Files\Microsoft Office\Visio14\1033...




```vb
Public Sub Stylesheet_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .Stylesheet = "your_Visio_path\your_language_ID\Steel.css" 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```


