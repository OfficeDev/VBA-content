---
title: DrawingControl.AfterRemoveHiddenInformation Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.AfterRemoveHiddenInformation
ms.assetid: 53c41981-3f24-53e3-dea5-204e0ad6f046
ms.date: 06/08/2017
---


# DrawingControl.AfterRemoveHiddenInformation Event (Visio)

Occurs when hidden information is removed from the document.


## Syntax

 Private Sub _expression_ _**AfterRemoveHiddenInformation**( **_ByVal doc As IVDOCUMENT_** )

 _expression_ An expression that returns a **DrawingControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document from which hidden information has been removed.|

## Remarks

The  **AfterRemoveHiddenInformation** event is one of a group of events for which the **EventInfo** property of the **Application** object contains extra information.

When the  **AfterRemoveHiddenInformation** event is fired, the **EventInfo** property returns a string that contains information about which items were removed from the document, consisting of the sum of applicable constant values from the **[VisRemoveHiddenInfoItems](visremovehiddeninfoitems-enumeration-visio.md)** enumeration.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


