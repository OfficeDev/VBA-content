---
title: Window.SetViewRect Method (Visio)
keywords: vis_sdr.chm11616585
f1_keywords:
- vis_sdr.chm11616585
ms.prod: visio
api_name:
- Visio.Window.SetViewRect
ms.assetid: ab2da074-6e55-3aa7-9c4a-ae299b65a9c9
ms.date: 06/08/2017
---


# Window.SetViewRect Method (Visio)

Sets the page coordinates of a window's borders by adjusting the zoom level and center scroll position.


## Syntax

 _expression_ . **SetViewRect**( **_dLeft_** , **_dTop_** , **_dWidth_** , **_dHeight_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _dLeft_|Required| **Double**|The page coordinate of the left side of the window.|
| _dTop_|Required| **Double**|The page coordinate of the top of the window.|
| _dWidth_|Required| **Double**|The distance in page units from the left side to the right side of the window.|
| _dHeight_|Required| **Double**|The distance in page units from the top to the bottom of the window.|

### Return Value

Nothing


## Remarks

If the  **Window** object is not a **visDrawing** type, the **SetViewRect** method raises an exception.


