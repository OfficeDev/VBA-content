---
title: Page.ResizeToFitContents Method (Visio)
keywords: vis_sdr.chm10950820
f1_keywords:
- vis_sdr.chm10950820
ms.prod: visio
api_name:
- Visio.Page.ResizeToFitContents
ms.assetid: 26b96288-7d8b-a999-ef45-a586110cc8b9
ms.date: 06/08/2017
---


# Page.ResizeToFitContents Method (Visio)

Resizes the page, or the master's page, to fit tightly around the shapes or master that are on it.


## Syntax

 _expression_ . **ResizeToFitContents**

 _expression_ A variable that represents a **Page** object.


### Return Value

Nothing


## Remarks

After the page is resized, the page height and width and the PinX and PinY values of the shapes or master are typically changed.

Calling the  **ResizeToFitContents** method is the equivalent of selecting **Let Visio expand the page as needed** on the **Page Size** tab in the **Page Setup** dialog box (on the **Design** tab, click **Size**, and then click  **More Page Sizes**).

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVPage.ResizeToFitContents()**
    

