---
title: DrawingControl.PageChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.PageChanged
ms.assetid: d0d870b2-d2d1-4b56-39b9-8c4bbc5ea63c
ms.date: 06/08/2017
---


# DrawingControl.PageChanged Event (Visio)

Occurs after the name of a page, the background page associated with a page, or the page type (foreground or background) changes.


## Syntax

Private Sub  _expression_ _**PageChanged**( **_ByVal page As [IVPAGE]_** )

 _expression_ A variable that represents a **DrawingControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **[IVPAGE]**|The page that changed.|

## Remarks

If several pages of a document have default names and any page but the last is deleted, Microsoft Visio will automatically rename the remaining pages to preserve the naming sequence, but the renaming will not trigger the  **PageChanged** event. For example, if a document contains Page-1, Page-2, and Page-3, and then Page-1 is deleted, Page-2 will be renamed Page-1, and Page-3 will be renamed Page-2, but no **PageChanged** event occurs.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


