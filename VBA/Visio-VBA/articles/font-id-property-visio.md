---
title: Font.ID Property (Visio)
keywords: vis_sdr.chm12013675
f1_keywords:
- vis_sdr.chm12013675
ms.prod: visio
api_name:
- Visio.Font.ID
ms.assetid: 2ffce82a-7002-584e-3fb2-6482757e33db
ms.date: 06/08/2017
---


# Font.ID Property (Visio)

Gets the ID of an object. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents a **Font** object.


### Return Value

Long


## Remarks

For  **Font** objects, the **ID** property corresponds to the number stored in the Font cell of the row in a shape's Character Properties section. For example, to apply the font named "Arial" to a shape's text, create a **Font** object that represents "Arial," get the ID of that font, and then set the **CharProps** property of the **Shape** object to that ID.

The ID associated with a particular font varies from system to system or as fonts are installed and removed on a given system.


