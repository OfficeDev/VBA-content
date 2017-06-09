---
title: VisFilterActions Enumeration (Visio)
ms.prod: visio
ms.assetid: 1b96bdba-e5e8-0e24-697d-3791c059fa15
ms.date: 06/08/2017
---


# VisFilterActions Enumeration (Visio)

Drag-state extensions of the  **MouseMove** event to filter, corresponding to user mouse actions related to dragging and dropping Microsoft Visio objects. Passed to the **Event.SetFilterActions** method and returned by the **Event.GetFilterActions** method. By filtering, you can specify which mouse actions (event extensions) you want to listen to.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visFilterMouseMoveDragBegin**|1|Filter the  **DragBegin** extension of the **MouseMove** event.|
| **visFilterMouseMoveDragDrop**|5|Filter the  **DragDrop** extension of the **MouseMove** event.|
| **visFilterMouseMoveDragEnter**|2|Filter the  **DragEnter** extension of the **MouseMove** event.|
| **visFilterMouseMoveDragLeave**|4|Filter the  **DragLeave** extension of the **MouseMove** event.|
| **visFilterMouseMoveDragOver**|3|Filter the  **DragOver** extension of the **MouseMove** event.|
| **visFilterMouseMoveNoDrag**|0|Do not filter any extensions of the  **MouseMove** event.|

