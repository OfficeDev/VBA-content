---
title: Page.PageDeleteCanceled Event (Visio)
keywords: vis_sdr.chm10919360
f1_keywords:
- vis_sdr.chm10919360
ms.prod: visio
api_name:
- Visio.Page.PageDeleteCanceled
ms.assetid: 5fa17e8b-5c80-962b-482e-f9c46f543a65
ms.date: 06/08/2017
---


# Page.PageDeleteCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelPageDelete** event.


## Syntax

Private Sub  _expression_ _**PageDeleteCanceled**( **_ByVal Page As [IVPAGE]_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **[IVPAGE]**|The page that was going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


