---
title: Window.Master Property (Visio)
keywords: vis_sdr.chm11613870
f1_keywords:
- vis_sdr.chm11613870
ms.prod: visio
api_name:
- Visio.Window.Master
ms.assetid: caf28e17-797a-91b2-c685-27ad0addddfd
ms.date: 06/08/2017
---


# Window.Master Property (Visio)

Gets the master that is displayed in a window. Read-only.


## Syntax

 _expression_ . **Master**

 _expression_ A variable that represents a **Window** object.


### Return Value

Variant


## Remarks

You can use the  **SubType** property of the **Window** object to determine whether the **Window** object shows a master. If the **Window** object does not show a master, the **Master** property raises an exception.

If the  **Window** object shows a master that is open for editing, the master returned is the temporary master being edited, not the original master that was opened for editing.


