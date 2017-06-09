---
title: Window.ShowScrollBars Property (Visio)
keywords: vis_sdr.chm11650845
f1_keywords:
- vis_sdr.chm11650845
ms.prod: visio
api_name:
- Visio.Window.ShowScrollBars
ms.assetid: 46be2c47-d9b0-c3d8-6f8b-cc728feb4ccb
ms.date: 06/08/2017
---


# Window.ShowScrollBars Property (Visio)

Determines whether scroll bars are shown in the drawing window. Read/write.


## Syntax

 _expression_ . **ShowScrollBars**

 _expression_ A variable that represents a **Window** object.


### Return Value

Integer


## Remarks

The  **ShowScrollBars** property value can be any combination of the following **VisScrollbarStates** constants, which are declared in the Visio type library. The default value is **visScrollBarBoth** , which shows both horizontal and vertical scroll bars.



|**Constant **|**Value **|
|:-----|:-----|
| **visScrollBarNeither**|&;H0 |
| **visScrollBarHoriz**|&;H1 |
| **visScrollBarVert**|&;H4 |
| **visScrollBarBoth**|&;H5 |

