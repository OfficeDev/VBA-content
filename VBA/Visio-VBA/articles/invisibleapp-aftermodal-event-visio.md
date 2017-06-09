---
title: InvisibleApp.AfterModal Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.AfterModal
ms.assetid: 52c6926d-0b94-8425-729c-9494f5c7879e
ms.date: 06/08/2017
---


# InvisibleApp.AfterModal Event (Visio)

Occurs after the Microsoft Visio instance leaves a modal state.


## Syntax

Private Sub  _expression_ _**AfterModal**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance that is no longer modal.|

## Remarks

Visio becomes modal when it displays a dialog box. A modal instance of Visio does not handle Automation calls. The  **BeforeModal** event indicates that the instance is about to become modal, and the **AfterModal** event indicates that the instance is no longer modal.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


