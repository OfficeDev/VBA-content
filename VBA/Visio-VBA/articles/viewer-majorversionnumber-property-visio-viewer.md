---
title: Viewer.MajorVersionNumber Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.MajorVersionNumber
ms.assetid: 3c96ad5b-811d-88e1-38de-8d8010b83c7e
ms.date: 06/08/2017
---


# Viewer.MajorVersionNumber Property (Visio Viewer)

 Gets the major version number of Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **MajorVersionNumber**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Long**


## Remarks

The major version number of Visio Viewer corresponds to the major version number of Microsoft Office. For the Office 2010 version of Visio Viewer, the major version number is 14. For Visio Viewer 2007, the major version number is 12. For Visio Viewer 2003, the major version number is 11, and for Visio Viewer 2002, the major version number is 10.


## Example

The following code displays the major version number in the  **Immediate** window.


```vb
Debug.Print vsoViewer.MajorVersionNumber
```


