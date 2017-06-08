---
title: Style.Index Property (Visio)
keywords: vis_sdr.chm11413695
f1_keywords:
- vis_sdr.chm11413695
ms.prod: visio
api_name:
- Visio.Style.Index
ms.assetid: 1a1b0efc-4a66-27f6-9d37-85105987b0b8
ms.date: 06/08/2017
---


# Style.Index Property (Visio)

Gets the ordinal position of a  **Style** object in the **Styles** collection. Read-only.


## Syntax

 _expression_ . **Index**

 _expression_ A variable that represents a **Style** object.


### Return Value

Long


## Remarks

Most collections are indexed starting with 1 rather than zero (0), so the index of the first element is 1, the index of the second element is 2, and so forth. The index of the last element in a collection is the same as the value of that collection's  **Count** property. You can iterate through a collection by using these index values. Adding objects to or deleting objects from a collection can change the index values of other objects in the collection.

There are some exceptions. The  **Colors** collection is indexed starting with 0. This is consistent with the numbering displayed next to the colors that appear in the **Color Palette** dialog box (on the **Tools** menu, click **Color Palette** ).

These collections are also indexed starting with 0:  **AccelItems** , **AccelTables** , **MenuSets** , **MenuItems** , **Menus** , **ToolbarItems** , **Toolbars** , and **ToolbarSets** .


