---
title: Windows.WindowChanged Event (Visio)
keywords: vis_sdr.chm11719275
f1_keywords:
- vis_sdr.chm11719275
ms.prod: visio
api_name:
- Visio.Windows.WindowChanged
ms.assetid: 02893ec6-2aca-cb70-919d-6ea2d37bb915
ms.date: 06/08/2017
---


# Windows.WindowChanged Event (Visio)

Occurs when the size or position of a window changes.


## Syntax

Private Sub  _expression_ _**WindowChanged**( **_ByVal Window As [IVWINDOW]_** )

 _expression_ A variable that represents a **Windows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window whose size or position has changed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


