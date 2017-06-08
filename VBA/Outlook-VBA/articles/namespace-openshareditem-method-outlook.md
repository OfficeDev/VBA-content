---
title: NameSpace.OpenSharedItem Method (Outlook)
keywords: vbaol11.chm789
f1_keywords:
- vbaol11.chm789
ms.prod: outlook
api_name:
- Outlook.NameSpace.OpenSharedItem
ms.assetid: ebfed85c-0af5-eb72-7a58-ae9e8b655347
ms.date: 06/08/2017
---


# NameSpace.OpenSharedItem Method (Outlook)

Opens a shared item from a specified path or URL.


## Syntax

 _expression_ . **OpenSharedItem**( **_Path_** )

 _expression_ An expression that returns a **[NameSpace](namespace-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path or URL of the shared item to be opened.|

### Return Value

An  **Object** representing the appropriate Outlook item for the shared item.


## Remarks

This method is used to open iCalendar appointment (.ics) files, vCard (.vcf) files, and Outlook message (.msg) files. The type of object returned by this method depends on the type of shared item opened, as described in the following table.



| **Shared item type**| **Outlook item**|
|iCalendar appointment (.ics) file| **[AppointmentItem](appointmentitem-object-outlook.md)**|
|vCard (.vcf) file| **[ContactItem](contactitem-object-outlook.md)**|
|Outlook message (.msg) file|Type corresponds to the type of the item that was saved as the .msg file|

 **Note**  This method does not support iCalendar calendar (.ics) files. To open iCalendar calendar files, you can use the  **[OpenSharedFolder](namespace-opensharedfolder-method-outlook.md)** method of the **NameSpace** object.


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

