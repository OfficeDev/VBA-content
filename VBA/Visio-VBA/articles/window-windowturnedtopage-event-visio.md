---
title: Window.WindowTurnedToPage Event (Visio)
keywords: vis_sdr.chm11619285
f1_keywords:
- vis_sdr.chm11619285
ms.prod: visio
api_name:
- Visio.Window.WindowTurnedToPage
ms.assetid: f1f92687-41b3-fc58-d862-93d4343c5808
ms.date: 06/08/2017
---


# Window.WindowTurnedToPage Event (Visio)

Occurs after a window shows a different page.


## Syntax

Private Sub  _expression_ _**WindowTurnedToPage**( **_ByVal Window As [IVWINDOW]_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that shows a different page.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


