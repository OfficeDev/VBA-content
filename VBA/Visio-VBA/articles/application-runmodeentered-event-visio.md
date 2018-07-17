---
title: Application.RunModeEntered Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.RunModeEntered
ms.assetid: 3a8827d9-ff0c-a1c4-2848-72758277aff4
ms.date: 06/08/2017
---


# Application.RunModeEntered Event (Visio)

Occurs after a document enters run mode.


## Syntax

Private Sub  _expression_ _**RunModeEntered**( **_ByVal doc As [IVDOCUMENT]_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that entered run mode.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


