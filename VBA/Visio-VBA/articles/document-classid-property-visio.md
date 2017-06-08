---
title: Document.ClassID Property (Visio)
keywords: vis_sdr.chm10513240
f1_keywords:
- vis_sdr.chm10513240
ms.prod: visio
api_name:
- Visio.Document.ClassID
ms.assetid: 668fec9a-eadf-a496-5db3-b91e30237c11
ms.date: 06/08/2017
---


# Document.ClassID Property (Visio)

Returns the class ID string of the container application in which the document is embedded. Read-only.


## Syntax

 _expression_ . **ClassID**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

If the class ID of the container application in which the document is embedded cannot be retrieved, the  **ClassID** property raises an exception. If the document is not embedded in a container, the **ClassID** property returns an empty string.

 **ClassID** returns a string of the form:




```
    {2287DC42-B167-11CE-88E9-002AFDDD917}
```

This string identifies the application that contains the document. It might, for example, identify the document's container as a Microsoft Excel object.


