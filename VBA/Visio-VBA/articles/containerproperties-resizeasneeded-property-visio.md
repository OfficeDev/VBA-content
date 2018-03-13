---
title: ContainerProperties.ResizeAsNeeded Property (Visio)
keywords: vis_sdr.chm17662610
f1_keywords:
- vis_sdr.chm17662610
ms.prod: visio
api_name:
- Visio.ContainerProperties.ResizeAsNeeded
ms.assetid: 13bd0493-95fd-73bf-454c-a39c69589bcd
ms.date: 06/08/2017
---


# ContainerProperties.ResizeAsNeeded Property (Visio)

Determines whether the container boundary resizes automatically to fit its contents. Read/write.


## Syntax

 _expression_ . **ResizeAsNeeded**

 _expression_ An expression that returns a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Return Value

 **[VisContainerAutoResize](viscontainerautoresize-enumeration-visio.md)**


## Remarks

The value of the  **ResizeAsNeeded** property can be one of the following **VisContainerAutoResize** constants.



| <strong>Constant</strong>                             | <strong>Value</strong> | <strong>Description</strong>                                  |
|:------------------------------------------------------|:-----------------------|:--------------------------------------------------------------|
| <strong>visContainerAutoResizeNone</strong>           | 0                      | Do not automatically resize container.                        |
| <strong>visContainerAutoResizeExpand</strong>         | 1                      | Automatically expand the container size, but do not contract. |
| <strong>visContainerAutoResizeExpandContract</strong> | 2                      | Automatically expand and contract the container size.         |

The setting of the  **ResizeAsNeeded** property corresponds to the selection in the **Automatic Resize** drop-down list in the **Size** group on the **Container Tools Format** tab.


