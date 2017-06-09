---
title: DrawingControl.QueryCancelGroup Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.QueryCancelGroup
ms.assetid: 630abedc-0b1a-8ad4-47d7-51215c1f0c43
ms.date: 06/08/2017
---


# DrawingControl.QueryCancelGroup Event (Visio)

Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _**QueryCancelGroup**( **_ByVal Selection As [IVSELECTION]_** )

 _expression_ A variable that represents a **DrawingControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that is going to be grouped.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelGroup** after the user has directed the instance to group a selection of shapes.




- If any event handler returns  **True** (cancel), the instance fires **GroupCanceled** and does not group the shapes.
    
- If all handlers return  **False** (don't cancel), the grouping is performed.
    


In some cases, such as when a shape that has a  **ForeignType** property of **visTypeMetafile** is grouped, the initial shape will be deleted and replaced with new shapes. In such cases, the Visio instance will subsequently fire **BeforeSelectionDelete** and **BeforeShapeDelete** events before grouping the shapes.

While a Visio instance is firing a query or cancel event, it will respond to inquiries from client code but will refuse to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


