---
title: Document.SavePreviewMode Property (Visio)
keywords: vis_sdr.chm10514290
f1_keywords:
- vis_sdr.chm10514290
ms.prod: visio
api_name:
- Visio.Document.SavePreviewMode
ms.assetid: e40f2b06-c9fd-3133-73c9-306f46f21e55
ms.date: 06/08/2017
---


# Document.SavePreviewMode Property (Visio)

Determines whether and how a preview picture is saved in a file. Read/write.


## Syntax

 _expression_ . **SavePreviewMode**

 _expression_ A variable that represents a **Document** object.


### Return Value

VisSavePreviewMode


## Remarks

The value of the  **SavePreviewMode** property corresponds to the **Save preview picture** setting on the **Summary** tab of the **Properties** dialog box. (Click the **File** tab, click **Info**, click  **Properties**, and then click  **Advanced Properties**.) A preview of the first page appears in the  **Open** dialog box. The value of **SavePreviewMode** can be one of the following **VisSavePreviewMode** constants. Selecting the **Save preview mode** checkbox is equivalent to setting the **SavePreviewMode** property to **visSavePreviewDraft1st** , which is the default.



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visSavePreviewNone**| 0| No preview picture.|
| **visSavePreviewDraft1st**| 1| The first page; includes only Visio shapes. Does not include embedded objects, text, or gradient fills.|
| **visSavePreviewDetailed1st**| 2| The first page; includes all objects.|

