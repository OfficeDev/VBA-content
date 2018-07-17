---
title: Pages.BeforeShapeDelete Event (Visio)
keywords: vis_sdr.chm11019065
f1_keywords:
- vis_sdr.chm11019065
ms.prod: visio
api_name:
- Visio.Pages.BeforeShapeDelete
ms.assetid: e83bb4cc-b9a0-1435-507f-149f5a108ab5
ms.date: 06/08/2017
---


# Pages.BeforeShapeDelete Event (Visio)

Occurs before a shape is deleted.


## Syntax

Private Sub  _expression_ _**BeforeShapeDelete**( **_ByVal Shape As [IVSHAPE]_** )

 _expression_ A variable that represents a **Pages** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that is going to be deleted.|

## Remarks

A  **Shape** object can serve as the source object for the **BeforeShapeDelete** event if the shape's **Type** property is **visTypeGroup** (2) or **visTypePage** (1).

The  **BeforeSelectionDelete** and **BeforeShapeDelete** events are similar in that they both fire before shape(s) are deleted. They differ in how they behave when a single operation deletes several shapes. Suppose a **Cut** operation deletes three shapes. The **BeforeShapeDelete** event fires three times and acts on each of the three objects. The **BeforeSelectionDelete** event fires once, and it acts on a **Selection** object in which the three shapes that you want to delete are selected.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).




 **Note**  You can use the VBA  **WithEvents** keyword to sink the **BeforeShapeDelete** event.

For performance considerations, the  **Document** object's event set does not include the **BeforeShapeDelete** event. To sink the **BeforeShapeDelete** event from a **Document** object (and from the **ThisDocument** object in a VBA project), you must use the **AddAdvise** method.


