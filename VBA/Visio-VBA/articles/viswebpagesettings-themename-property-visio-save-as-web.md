---
title: VisWebPageSettings.ThemeName Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.ThemeName
ms.assetid: 9efd26b1-7426-1ff4-0b51-5463a2beb822
ms.date: 06/08/2017
---


# VisWebPageSettings.ThemeName Property (Visio Save As Web)

Assigns a Web page theme to the page you are creating. Read/write.


## Syntax

 _expression_. **ThemeName**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **String**


## Remarks

You can use themes that are provided by Microsoft Visio or themes that you create yourself. If you want to create your own theme, do the following: 


1. Create an HTM file that contains the following term in an HTML tag: "##VIS_SAW_FILE##"Visio recognizes HTM files that contain this tag as theme files.
    
2. Store the file in the following folder:\  _your_Visio_path_ \ _your_language_ID_ \
    


Your theme file will then appear in the  **Host in Web page** drop-down list in the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, click  **Publish**, and then click  **Advanced**).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ThemeName** property to assign the "Basic" theme (supplied by Visio) to the Web page you are creating.

Before running this macro, replace  _path\filename.htm_ with a valid target path on your computer and the file name that you want to assign to your Web page. Also, replace _your_Visio_path_ and _your_language_ID_ with the path to Microsoft Visio on your computer, for example:

C:\Program Files\Microsoft Office\Visio14\1033\...




```vb
Public Sub ThemeName_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .ThemeName = "your_Visio_path\your_language_ID\Basic.htm" 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```


