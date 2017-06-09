---
title: Viewer.AlertsEnabled Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.AlertsEnabled
ms.assetid: 1bf74608-3652-b015-f862-b503d11e5c77
ms.date: 06/08/2017
---


# Viewer.AlertsEnabled Property (Visio Viewer)

Gets or sets a value that indicates whether warnings and alerts appear when an error occurs in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **AlertsEnabled**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Boolean**


## Remarks

The default is for warnings and alerts to appear ( **True**).


## Example

The following code shows how to determine whether alerts are enabled in Visio Viewer.


```vb
 Debug.Print vsoViewer.AlertsEnabled
```


