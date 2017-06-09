---
title: Viewer.OnLayerChanged Event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.OnLayerChanged
ms.assetid: d0731153-f975-cde1-3649-be34df859168
ms.date: 06/08/2017
---


# Viewer.OnLayerChanged Event (Visio Viewer)

Occurs when a layer is changed in the document open in Microsoft Visio Viewer.


## Syntax

 _expression_. **OnLayerChanged**( **_LayerIndex_**,  **_Visible_**,  **_ColorOverride_**,  **_Color_**,  **_ColorTrans_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|LayerIndex|Required| **Long**|The index of the changed layer.|
|Visible|Required| **Boolean**|Indicates whether the changed layer is visible in the user interface.|
|ColorOverride|Required| **Boolean**|Indicates whether to override the color of shapes on the changed layer.|
|Color|Required| **[OLE_COLOR]**|The color of the changed layer, expressed in RGB values.|
|ColorTrans|Required| **Double**|The transparency percentage of the changed layer.|

## Remarks

You can change a layer either in the  **Layer Properties** dialog box, or programmatically by using the **[LayerColor](viewer-layercolor-property-visio-viewer.md)**,  **[LayerColorOverride](viewer-layercoloroverride-property-visio-viewer.md)**,  **[LayerColorTrans](viewer-layercolortrans-property-visio-viewer.md)**, and  **[LayerVisible](viewer-layervisible-property-visio-viewer.md)** properties.


## Example

The following code shows how to use the  **OnLayerChanged** event to display the new transparency percentage of the changed layer in the **Immediate** window.


```vb
Private Sub vsoViewer_OnLayerChanged(ByVal LayerIndex As Long, ByVal Visible As Boolean, ByVal ColorOverride As Boolean, ByVal Color As stdole.OLE_COLOR, ByVal ColorTrans As Double)

    Debug.Print "The new transparency percentage is"; ColorTrans

End Sub
```


