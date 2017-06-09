---
title: ViewCtl.Selection Property (Outlook View Control)
ms.prod: outlook
ms.assetid: 2f4549bb-a480-bcbb-0fde-90a50460aa92
ms.date: 06/08/2017
---


# ViewCtl.Selection Property (Outlook View Control)

Returns a  [Selection](selection-object-outlook.md) object that consists of one or more items that are selected in the current view. Read-only.


## Syntax

 _expression_. **Selection**

 _expression_A variable that represents a  **ViewCtl** object.


## Remarks

If the current folder is a file system folder, or if Outlook  **Today** or any folder with a current Web view is currently displayed, this property returns an empty collection.


