---
title: VisSaveAsWeb Object (Visio Save As Web)
ms.prod: visio
ms.assetid: 48e19e11-9b41-42ec-84e9-c4aab7f08784
ms.date: 06/08/2017
---


# VisSaveAsWeb Object (Visio Save As Web)

Contains the web page property settings and methods used when a Visio drawing is saved as a web page. 


## Remarks

The  **VisSaveAsWeb** object contains the methods and property settings that are used when a selected Visio drawing is saved as a web page. The web page project includes the following files:


- An HTML version of the drawing (including shape data, formerly called custom properties, and multiple drawing pages, if applicable)
    
- The supporting files associated with the project, for example, the graphics files (GIFs and JPGs), script files, data (XML) files, and cascading style sheet (CSS) files.
    
To set the properties for your web page, use the  **[WebPageSettings](vissaveasweb-webpagesettings-property-visio-save-as-web.md)** property of the **VisSaveAsWeb** object to get a ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object. After the properties are set, perform the following steps.


1. Call the  **[AttachToVisioDoc](vissaveasweb-attachtovisiodoc-method-visio-save-as-web.md)** method to specify the drawing to be saved as a web page. For example:
    
```
  vsoSaveAsWeb.AttachToVisioDoc _ 
Application.Documents.Open("drive:\folder\drawingname.vdx")
```


    If you don't call this method, Visio creates the page from the active document by default.
    
2. Call the  **[CreatePages](vissaveasweb-createpages-method-visio-save-as-web.md)** method to create the web page. For example:
    
```
  vsoSaveAsWeb.CreatePages vsoSaveAsWeb.CreatePages
```

You can control certain user interface behavior during page creation by using the  **[SilentMode](viswebpagesettings-silentmode-property-visio-save-as-web.md)** property or the **[QuietMode](viswebpagesettings-quietmode-property-visio-save-as-web.md)** property of the **VisWebPageSettings** object.

The files created by the Save as Web Page feature are placed into the target path you specify, or a location you specify in the  **[TargetPath](viswebpagesettings-targetpath-property-visio-save-as-web.md)** property of the **VisWebPageSettings** object.


 **Note**  You must specify a target path, or Visio will generate an error.

They can be organized as flat files or in a subfolder that has the same name as the drawing (see the  **[StoreInFolder](viswebpagesettings-storeinfolder-property-visio-save-as-web.md)** property of the **VisWebPageSettings** object).


 **Note**  To view the  **VisSaveAsWeb** class in the Object Browser, make sure that you have a reference to the Save As Web Page DLL in your project (in the Visual Basic Editor window, click **References**, on the  **Tools** menu, and then select the **Microsoft Visio 15.0 SaveAsWeb Type Library** check box in the **Available References** list).


