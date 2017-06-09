---
title: Documents.ShapeDataGraphicChanged Event (Visio)
keywords: vis_sdr.chm10662010
f1_keywords:
- vis_sdr.chm10662010
ms.prod: visio
api_name:
- Visio.Documents.ShapeDataGraphicChanged
ms.assetid: 47fa7996-9b1f-d530-adf1-8dbb6d69aa5c
ms.date: 06/08/2017
---


# Documents.ShapeDataGraphicChanged Event (Visio)

Occurs after a data graphic is applied to or deleted from a shape.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

Private Sub  _expression_ _**ShapeDataGraphicChanged**( **_ByVal Shape As IVSHAPE_** )

 _expression_ An expression that returns a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape to which the data graphic was applied or from which it was deleted.|

## Remarks

A data graphic is a  **Master** object of type **visTypeDataGraphic** . When the same master that is already applied to a shape is re-applied to the shape, the **ShapeDataGraphicChanged** event does not fire, even if the master has been modified since it was applied originally. If, however, a different data graphic master is applied to the shape, the event does fire.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


