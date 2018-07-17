---
title: Styles.BeforeStyleDelete Event (Visio)
keywords: vis_sdr.chm11519070
f1_keywords:
- vis_sdr.chm11519070
ms.prod: visio
api_name:
- Visio.Styles.BeforeStyleDelete
ms.assetid: e73533d6-c5ce-739c-f85d-0137794ac953
ms.date: 06/08/2017
---


# Styles.BeforeStyleDelete Event (Visio)

Occurs before a style is deleted.


## Syntax

Private Sub  _expression_ _**BeforeStyleDelete**( **_ByVal Style As [IVSTYLE]_** )

 _expression_ A variable that represents a **Styles** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[IVSTYLE]**|The style that is going to be deleted.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


