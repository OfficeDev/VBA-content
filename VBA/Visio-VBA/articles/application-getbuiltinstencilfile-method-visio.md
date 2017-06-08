---
title: Application.GetBuiltInStencilFile Method (Visio)
keywords: vis_sdr.chm10062110
f1_keywords:
- vis_sdr.chm10062110
ms.prod: visio
api_name:
- Visio.Application.GetBuiltInStencilFile
ms.assetid: 2ae65aaa-d441-c7e8-3c8c-737bcca84738
ms.date: 06/08/2017
---


# Application.GetBuiltInStencilFile Method (Visio)

Returns the file path to the specified built-in, hidden stencil used to populate certain galleries in the Microsoft Visio user interface.


## Syntax

 _expression_ . **GetBuiltInStencilFile**( **_StencilType_** , **_MeasurementSystem_** )

 _expression_ A variable that represents an **[Application](application-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StencilType_|Required| **[VisBuiltInStencilTypes](visbuiltinstenciltypes-enumeration-visio.md)**|The stencil to retrieve. See Remarks for possible values.|
| _MeasurementSystem_|Required| **[VisMeasurementSystem](vismeasurementsystem-enumeration-visio.md)**|The measurement system for the stencil.|

### Return Value

 **String**


## Remarks

The  _StencilType_ parameter value must be one of the following **VisBuiltInStencilTypes** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visBuiltInStencilBackgrounds**|0|The hidden stencil that contains the shapes displayed in the  **Backgrounds** gallery ( **Design** tab).|
| **visBuiltInStencilBorders**|1|The hidden stencil that contains the shapes displayed in the  **Borders and Titles** gallery ( **Design** tab).|
| **visBuiltInStencilContainers**|2|The hidden stencil that contains the shapes displayed in the  **Container** gallery ( **Insert** tab).|
| **visBuiltInStencilCallouts**|3|The hidden stencil that contains the shapes displayed in the  **Callout** gallery ( **Insert** tab).|
| **visBuiltInStencilLegends**|4|The hidden stencil that contains the shapes displayed in the  **Insert Legend** gallery ( **Data** tab).|

## Example

The following Visual Basic for Applications (VBA) code sample shows how to use the  **GetBuiltInStencilFile** method to open the built-in, hidden, container stencil, and to add one of the containers from that stencil to the active page to contain the selected shape or shapes. Before you run this code, be sure that there is a selected shape (or a selection of shapes) on the active page.


```vb
Public Sub GetBuiltInStencilFile_Example()

    Dim vsoDocument As Visio.Document
    Set vsoDocument = Application.Documents.OpenEx(Application.GetBuiltInStencilFile(visBuiltInStencilContainers, visMSUS), visOpenHidden)
    Application.ActivePage.DropContainer vsoDocument.Masters.ItemU("Container 1"), Application.ActiveWindow.Selection
    vsoDocument.Close

End Sub
```


