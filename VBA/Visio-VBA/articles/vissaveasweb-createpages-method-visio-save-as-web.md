---
title: VisSaveAsWeb.CreatePages Method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.CreatePages
ms.assetid: 48094af2-55fb-9732-19bf-8a73827d1afb
ms.date: 06/08/2017
---


# VisSaveAsWeb.CreatePages Method (Visio Save As Web)

Initiates Web page creation.


## Syntax

 _expression_. **CreatePages**

 _expression_An expression that returns a  ** [VisSaveAsWeb](http://msdn.microsoft.com/library/c4675de8-0f63-179f-f687-8962d54d6b2f%28Office.15%29.aspx)** object.


### Return Value

 **Nothing**


## Remarks

Because the  **VisSaveAsWeb** object uses the settings in its ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object to create the Web page, you should call the **CreatePages** method after you make any required changes to the **VisWebPageSettings** object.

To specify which document to save as a Web page, use the  **[AttachToVisioDoc](vissaveasweb-attachtovisiodoc-method-visio-save-as-web.md)** method. If no document is specified, Microsoft Visio saves the active document by default.


## Example

The following example shows how to open an existing file and save it as a Web page by using the Save as Web Page feature's default settings and the  **AttachToVisioDoc** and **CreatePages** methods. Before running this example, replace _path\filename_ with a valid path and file name for a Visio document to pass to the **Open** method. In addition, replace _targetpath\filename_ with a valid target path and a file name for the Web page project files.


```vb
Public Sub CreatePages_Example () 
    Dim vsoSaveAsWeb As VisSaveAsWeb 
    Dim vsoWebSettings As VisWebPageSettings 
    Dim vsoDocument As Visio.Document
 
    Set vsoDocument = Application.Documents.Open("path\filename") 
    Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
    Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings
 
    vsoWebSettings.TargetPath = "targetpath\filename"
    
    With vsoSaveAsWeb
        .AttachToVisioDoc vsoDocument
        .CreatePages 
    End With
End Sub
```


