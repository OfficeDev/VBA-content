---
title: Window.NewWindow Method (Visio)
keywords: vis_sdr.chm11651350
f1_keywords:
- vis_sdr.chm11651350
ms.prod: visio
api_name:
- Visio.Window.NewWindow
ms.assetid: 0cca00d4-9cf4-6a30-b9f2-a37fbad69296
ms.date: 06/08/2017
---


# Window.NewWindow Method (Visio)

Opens a new Microsoft Visio window.


## Syntax

 _expression_ . **NewWindow**

 _expression_ A variable that represents a **Window** object.


### Return Value

Window


## Remarks

Calling the  **NewWindow** method is the equivalent of clicking **New Window** in the **Window** group on the **View** tab in the user interface. The **NewWindow** method works only on top-level mutliple-document interface (MDI) windows. The properties and caption of the new window are all determined by Visio, so the caller cannot specify any arguments to the **NewWindow** call. The new window's caption is the same as that of the existing one from which it was opened, but it is distinguished by a colon and a sequential number. For example, if the existing window is a drawing window that is captioned Drawing1, the new window is captioned Drawing1:2. In addition, the existing window's caption is changed to Drawing1:1.

The  **NewWindow** method opens a new window of the same type as the instance of the parent object. In other words, if the parent object is a drawing window, calling **NewWindow** opens another drawing window, while if is a ShapeSheet window, calling **NewWindow** opens another ShapeSheet window.




 **Note**  The  **NewWindow** method is not available for the Microsoft Visio Drawing Control.


