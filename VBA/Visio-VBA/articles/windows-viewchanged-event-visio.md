---
title: Windows.ViewChanged Event (Visio)
keywords: vis_sdr.chm11719260
f1_keywords:
- vis_sdr.chm11719260
ms.prod: visio
api_name:
- Visio.Windows.ViewChanged
ms.assetid: 0c504d9d-3664-38fc-33ad-cc7ec41589e2
ms.date: 06/08/2017
---


# Windows.ViewChanged Event (Visio)

Occurs when the zoom level or scroll position of a drawing window changes.


## Syntax

Private Sub  _expression_ _**ViewChanged**( **_ByVal Window As [IVWINDOW]_** )

 _expression_ A variable that represents a **Windows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window whose zoom level or scroll position changed.|

## Remarks

This event fires whenever the zoom level or scroll position of a  **Window** object of the type **visDrawing** changes.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


