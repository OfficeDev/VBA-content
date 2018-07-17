---
title: InvisibleApp.BeforeQuit Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.BeforeQuit
ms.assetid: b2554719-ada7-9bed-3ace-9e430c478e7a
ms.date: 06/08/2017
---


# InvisibleApp.BeforeQuit Event (Visio)

Occurs before a Microsoft Visio instance terminates.


## Syntax

Private Sub  _expression_ _**BeforeQuit**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that is going to terminate.|

## Remarks

When programming with Microsoft Visual Basic, use the  **BeforeDocumentClose** event instead of the **BeforeQuit** event. The code in a Visual Basic project of a Visio document never has the chance to respond to the **BeforeQuit** event because the project is a property of a document, and all documents are closed before the **BeforeQuit** event notification is sent.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


