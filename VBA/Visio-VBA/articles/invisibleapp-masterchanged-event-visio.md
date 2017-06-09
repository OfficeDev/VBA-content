---
title: InvisibleApp.MasterChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.MasterChanged
ms.assetid: bb64bee2-9562-8083-d993-d42c00ed889f
ms.date: 06/08/2017
---


# InvisibleApp.MasterChanged Event (Visio)

Occurs after properties of a master are changed and propagated to its instances.


## Syntax

Private Sub  _expression_ _**MasterChanged**( **_ByVal Master As [IVMASTER]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master whose properties changed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


