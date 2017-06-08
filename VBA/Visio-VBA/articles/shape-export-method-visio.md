---
title: Shape.Export Method (Visio)
keywords: vis_sdr.chm11216265
f1_keywords:
- vis_sdr.chm11216265
ms.prod: visio
api_name:
- Visio.Shape.Export
ms.assetid: f4051560-8719-ea9c-30eb-33230c95786c
ms.date: 06/08/2017
---


# Shape.Export Method (Visio)

Exports an object from Microsoft Visio to a file format such as .bmp, .dib, .dwg, .dxf, .emf, .emz, .gif, .htm, .jpg, .png, .svg, .svgz, .tif, or .wmf.


## Syntax

 _expression_ . **Export**( **_FileName_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The fully qualified path and name of the file to receive the exported object.|

### Return Value

Nothing


## Remarks

The file name extension indicates which export filter to use. If the filter is not installed, the  **Export** method returns a compiler error in your Visual Basic or VBA project. The **Export** method uses the default preference settings for the specified filter and does not prompt the user for non-default arguments.

The  **Export** method of a **Page** object supports saving to HTML file format using the extension .htm or .html. When pages are exported, Visio uses the settings that were last selected in the **Save As** dialog box.

If the specified file already exists, Visio replaces it without prompting the user.

Starting with Visio, you can use various properties and methods of the  **[ApplicationSettings](applicationsettings-object-visio.md)** object that relate to raster images to configure settings for export to .bmp, .gif, .jpg, .png, and .tif file types.


## Example

This example shows how to export a Visio page as a bitmap (.bmp) file. It assumes a drawing page is open and active in Microsoft Visio.


```vb
Public Sub Export_Example() 
 Dim vsoPage As Visio.Page 
 Set vsoPage = ActivePage 
 vsoPage.Export ("C:\\myExportedPage.bmp") 
End Sub
```


