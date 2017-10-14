---
title: Window.MergeClass Property (Visio)
keywords: vis_sdr.chm11650715
f1_keywords:
- vis_sdr.chm11650715
ms.prod: visio
api_name:
- Visio.Window.MergeClass
ms.assetid: 9ab7b4e7-9be3-9cfe-3a45-57825930ca15
ms.date: 06/08/2017
---


# Window.MergeClass Property (Visio)

Specifies a list of window classes that this anchored window can merge with. Read/write.


## Syntax

 _expression_ . **MergeClass**

 _expression_ A variable that represents a **Window** object.


### Return Value

String


## Remarks

Use semicolons to separate individual items in the list. For example, if the  **MergeClass** property returns a string that contains "123;789", it can merge with any windows that also contain "123" or "789" in its merge class list. Windows that have a merge class list that contains a zero-length string ("") can merge with other windows that contain a zero-length string ("") in their merge class list.

The  **MergeClass** property applies only to anchored windows. If the **Window** object is an MDI frame window, Microsoft Visio raises an exception.

At present, windows of type  **visDocked** can be merged only with other windows of type **visDocked** , and windows of type **visAnchorBar** can be merged only with other windows of type **visAnchorBar.**

Use the  **Type** property to determine window type.


