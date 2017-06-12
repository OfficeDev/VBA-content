---
title: Window.SubType Property (Visio)
keywords: vis_sdr.chm11614470
f1_keywords:
- vis_sdr.chm11614470
ms.prod: visio
api_name:
- Visio.Window.SubType
ms.assetid: 3e20338f-a63b-462c-731f-4790042b76cb
ms.date: 06/08/2017
---


# Window.SubType Property (Visio)

Returns the subtype of a  **Window** object that represents a drawing window. Read-only.


## Syntax

 _expression_ . **SubType**

 _expression_ A variable that represents a **Window** object.


### Return Value

Integer


## Remarks

If the  **Type** property of a **Window** object returns any value other than **visDrawing** , the **SubType** property returns the same value as the **Type** property. If the **Type** property of a **Window** object returns **visDrawing** , the **SubType** property returns one of the following values.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visPageWin**|128 |A drawing window showing a page. |
| **visPageGroupWin**|160 |A group editing window of a group on a page. |
| **visMasterWin**|64 |A master drawing page window. |
| **visMasterGroupWin**|96 |A group editing window of a group in a master. |

