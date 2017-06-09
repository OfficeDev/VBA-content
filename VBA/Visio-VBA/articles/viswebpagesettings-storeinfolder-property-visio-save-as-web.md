---
title: VisWebPageSettings.StoreInFolder Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.StoreInFolder
ms.assetid: ed0cf76a-a68d-cfa7-538c-91df5234a0d0
ms.date: 06/08/2017
---


# VisWebPageSettings.StoreInFolder Property (Visio Save As Web)

Determines whether supporting files for the Web page to be created are placed into a subfolder that has the same name as the root HTML file. Read/write.


## Syntax

 _expression_. **StoreInFolder**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **Long**


## Remarks

Set  **StoreInFolder** to a non-zero value ( **True**) to place supporting Web page files in a subfolder that has the same name as the root HTML file; otherwise, set it to zero ( **False**). 

If you set the  **StoreInFolder** property to **True** (non-zero), Microsoft Visio places the supporting files in a subfolder prefixed with the same name as the .htm file. If either the .htm file or the subfolder is moved or deleted, its corresponding subfolder or .htm file is also moved or deleted.

If you set the  **StoreInFolder** property to **False** (0), Visio places all supporting files in the same folder as the .htm file.

Setting the  **StoreInFolder** property to `True` is the equivalent of selecting the **Organize supporting files in a folder** check box on the **General** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, and then click  **Publish**).


## Example

The following macro shows how to set the  **StoreInFolder** property so that a subfolder that contains all a Web page's supporting files and has the same name as the .htm file is created.

Before running this macro, replace  _path\filename.htm_ with a valid target path on your computer and the filename that you want to assign to your Web page.




```vb
Public Sub StoreInFolder_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsowebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .StoreInFolder = True 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
 End Sub
```


