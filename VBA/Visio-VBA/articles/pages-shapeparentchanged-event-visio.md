---
title: Pages.ShapeParentChanged Event (Visio)
keywords: vis_sdr.chm11019235
f1_keywords:
- vis_sdr.chm11019235
ms.prod: visio
api_name:
- Visio.Pages.ShapeParentChanged
ms.assetid: 9a566e9f-479f-c69d-8831-21fd7694c201
ms.date: 06/08/2017
---


# Pages.ShapeParentChanged Event (Visio)

Occurs after shapes are grouped or a group is ungrouped.


## Syntax

Private Sub  _expression_ _**ShapeParentChanged**( **_ByVal Shape As [IVSHAPE]_** )

 _expression_ A variable that represents a **Pages** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape whose parent changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


