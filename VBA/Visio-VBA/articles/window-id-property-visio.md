---
title: Window.ID Property (Visio)
keywords: vis_sdr.chm11613675
f1_keywords:
- vis_sdr.chm11613675
ms.prod: visio
api_name:
- Visio.Window.ID
ms.assetid: bf05dfe0-b6c0-1ea9-7ce4-af2bd98bbecd
ms.date: 06/08/2017
---


# Window.ID Property (Visio)

Gets the ID of an object. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents a **Window** object.


### Return Value

Long


## Remarks

For  **Window** objects, the **ID** property can be used with the **ItemFromID** property of a **Windows** collection to retrieve a **Window** object from the collection without iterating through the collection. A **Window** object whose **Type** property is set to **visAnchorBarBuiltIn** returns an ID of **visWinIDCustProp** , **visWinIDDrawingExplorer** , **visWinIDFormulaTracing** , **visWinIDMasterExplorer** , **visWinIDPanZoom** , **visWinIDSizePos** , or **visWinIDStencilExplorer** . A **Window** object whose **Type** property is set to **visAnchorBarAddon** returns an ID that is unique within its **Windows** collection for the lifetime of that collection. If a **Window** object has an ID of **visInvalWinID** , you cannot use the **ItemFromID** property to retrieve the **Window** object from its collection.


