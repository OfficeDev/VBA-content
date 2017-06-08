---
title: DrawingControl.PageSizingBehavior Property (Visio)
keywords: vis_sdr.chm60138
f1_keywords:
- vis_sdr.chm60138
ms.prod: visio
api_name:
- Visio.PageSizingBehavior
ms.assetid: a16e860c-f60a-73b6-c978-7a5d70ccaa25
ms.date: 06/08/2017
---


# DrawingControl.PageSizingBehavior Property (Visio)

Specifies how drawing pages and shapes in the Microsoft Visio Drawing Control react when the control is resized, usually when an existing drawing file is loaded into the control by means of the  **Src** property. Read/write.


## Syntax

 _expression_ . **PageSizingBehavior**

 _expression_ A variable that represents a **DrawingControl** object.


### Return Value

VisPageSizingBehaviors


## Remarks

Possible values for  **PageSizingBehavior** are declared in the Visio type library in **VisPageSizingBehaviors** and shown in the following table.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visNeverResizePages**|0|Does not automatically resize pages under any circumstances. The default.|
| **visResizePages**|1|Automatically resizes all pages when the Visio Drawing Control is resized or when a new document is loaded into it. Leaves shapes unchanged.|
You can set the  **PageSizingBehavior** property either at design time (for example, in the **Properties** window in Microsoft Visual Basic 6.0), or at run time, typically in the **Form_Load()** procedure. It is recommended that you set **PageSizingBehavior** at design time.

If  **PageSizingBehavior** is set to **visResizePages** , when a new document is loaded into the Visio Drawing Control, the pages of that document are resized to match the size of the control itself. However, the shapes on those pages neither change size nor move; they retain their existing sizes and their locations relative to the coordinate system of the Visio page, which has its origin in the bottom left corner of the page.

If  **PageSizingBehavior** is set to the default, **visNeverResizePages** , when a new document is loaded into the Visio Drawing Control, the pages of that document (and the shapes on the page) retain their existing size. In this case, the size of the control in the container application has no relation to the size of the pages it displays; it is simply an open "window" onto the page or pages.


