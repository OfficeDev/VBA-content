---
title: Windows.WindowActivated Event (Visio)
keywords: vis_sdr.chm11719270
f1_keywords:
- vis_sdr.chm11719270
ms.prod: visio
api_name:
- Visio.Windows.WindowActivated
ms.assetid: 4092a0e7-1529-1af7-fc9b-f8b1b47ad906
ms.date: 06/08/2017
---


# Windows.WindowActivated Event (Visio)

Occurs after the active window changes in a Microsoft Visio instance.


## Syntax

Private Sub  _expression_ _**WindowActivated**( **_ByVal Window As [IVWINDOW]_** )

 _expression_ A variable that represents a **Windows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that was activated.|

## Remarks

The  **WindowActivated** event indicates that the active window has changed in a Visio instance. This event implies that the **ActiveDocument** and **ActivePage** properties of the **Application** object may also have changed; in contrast, any time the **ActiveDocument** or **ActivePage** property changes, a **WindowActivated** event is always generated.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


