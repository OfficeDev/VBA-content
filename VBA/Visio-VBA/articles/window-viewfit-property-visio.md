---
title: Window.ViewFit Property (Visio)
keywords: vis_sdr.chm11614645
f1_keywords:
- vis_sdr.chm11614645
ms.prod: visio
api_name:
- Visio.Window.ViewFit
ms.assetid: 5ee12ad7-4acf-aaf9-a928-93fc473e1c8f
ms.date: 06/08/2017
---


# Window.ViewFit Property (Visio)

Determines which auto-fit mode a window is in, if any. Read/write.


## Syntax

 _expression_ . **ViewFit**

 _expression_ A variable that represents a **Window** object.


### Return Value

Long


## Remarks

The  **ViewFit** property applies to drawing windows only, and can have the following values.



|** Constant**|** Value**|
|:-----|:-----|
| **visFitNone**| 0|
| **visFitPage**| 1|
| **visFitWidth**| 2|
If the value of the window's  **Type** property is not **visDrawing** , the **ViewFit** property returns **visFitNone** . Attempting to set the **ViewFit** property of this type of window raises an exception.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVWindow.ViewFit**
    

