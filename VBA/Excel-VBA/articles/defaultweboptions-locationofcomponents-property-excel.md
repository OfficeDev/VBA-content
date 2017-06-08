---
title: DefaultWebOptions.LocationOfComponents Property (Excel)
keywords: vbaxl10.chm660085
f1_keywords:
- vbaxl10.chm660085
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.LocationOfComponents
ms.assetid: a3f1571d-9301-4e3f-7467-f7716f26e45f
ms.date: 06/08/2017
---


# DefaultWebOptions.LocationOfComponents Property (Excel)

Returns or sets the central URL (on the intranet or Web) or path (local or network) to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. The default value is the local or network installation path for Microsoft Office. Read/write  **String** .


## Syntax

 _expression_ . **LocationOfComponents**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks

Office Web components are automatically downloaded with the specified Web page if the  **[DownloadComponents](defaultweboptions-downloadcomponents-property-excel.md)** property is is set to **True** , the components are not already installed, the path is valid and points to a location that contains the necessary components, and the user has a valid Microsoft Office license.


## Example

This example sets the path to the location from which users can download Microsoft Office Web components.


```vb
Application.DefaultWebOptions.DownloadComponents = True 
Application.DefaultWebOptions.LocationOfComponents = _ 
 Application.Path &; Application.PathSeparator &; "foo"
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

