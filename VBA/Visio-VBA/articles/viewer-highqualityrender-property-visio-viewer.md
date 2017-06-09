---
title: Viewer.HighQualityRender Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.HighQualityRender
ms.assetid: 39f59bc2-36ad-7c74-97de-85a486eb42c3
ms.date: 06/08/2017
---


# Viewer.HighQualityRender Property (Visio Viewer)

Gets or sets a value that indicates whether high-quality rendering is enabled in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **HighQualityRender**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Boolean**


## Remarks

High-quality rendering is slower but produces output that looks better.

The default is for high-quality rendering to be enabled (property value set to  **True**).


## Example

The following code gets a value that indicates whether high-quality rendering is enabled in Visio Viewer.


```vb
Debug.Print vsoViewer.HighQualityRender
```


