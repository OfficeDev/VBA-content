---
title: Window.AllowEditing Property (Visio)
keywords: vis_sdr.chm11650505
f1_keywords:
- vis_sdr.chm11650505
ms.prod: visio
api_name:
- Visio.Window.AllowEditing
ms.assetid: 805ed8a9-1835-0d7b-9bbe-717ff21af3c9
ms.date: 06/08/2017
---


# Window.AllowEditing Property (Visio)

Determines whether the  **Edit Stencil** command is enabled or disabled in a stencil window. Read/write.


## Syntax

 _expression_ . **AllowEditing**

 _expression_ A variable that represents a **Window** object.


### Return Value

Boolean


## Remarks

Use the  **AllowEditing** property to prevent unintentional editing in stencils. Setting the value of this property for stencils that are already open for editing has no effect. This property has no effect on Visio stencils (stencils that are shipped on the Microsoft Visio CD) or on other stencils that have been published to Visio by using an .msi file.




 **Note**  Only user-created stencils are editable. By default, Visio stencils are not editable. 


