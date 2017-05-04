---
title: Viewer.OnDocumentUnloaded Event (Visio Viewer)
ms.prod: VISIO
api_name:
- Visio.OnDocumentUnloaded
ms.assetid: b2f1d5ad-122d-6e55-1cb0-63c78f79bc2b
---


# Viewer.OnDocumentUnloaded Event (Visio Viewer)

Occurs after the current document in Microsoft Visio Viewer is unloaded.


## Syntax

 _expression_. **OnDocumentUnloaded**

 _expression_An expression that returns a  **Viewer** object.


## Remarks

You can unload the current document in Visio Viewer programmatically by using the  **[Unload](viewer-unload-method-visio-viewer.md)** method.


## Example


```vb
Private Sub vsoViewer_OnDocumentUnloaded()

    Debug.Print "Current document unloaded."

End Sub
```


