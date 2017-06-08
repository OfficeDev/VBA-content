---
title: Window.GetViewRect Method (Visio)
keywords: vis_sdr.chm11616325
f1_keywords:
- vis_sdr.chm11616325
ms.prod: visio
api_name:
- Visio.Window.GetViewRect
ms.assetid: 3281d1af-6745-1b74-5071-e388fa1dc32c
ms.date: 06/08/2017
---


# Window.GetViewRect Method (Visio)

Returns the page coordinates of a window's borders.


## Syntax

 _expression_ . **GetViewRect**( **_pdLeft_** , **_pdTop_** , **_pdWidth_** , **_pdHeight_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pdLeft_|Required| **Double**|The coordinate in page units of the left side of the window.|
| _pdTop_|Required| **Double**|The coordinate in page units of the top of the window.|
| _pdWidth_|Required| **Double**|The distance in page units from the left side to the right side of the window.|
| _pdHeight_|Required| **Double**|The distance in page units from the top to the bottom of the window.|

### Return Value

Nothing


## Remarks

The values that the  **GetViewRect** method returns are affected by whether page tabs and rulers are displayed on the Microsoft Visio drawing page.

If the  **Window** object is not a **visDrawing** type, the **GetViewRect** method raises an exception.


