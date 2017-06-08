---
title: Window.Page Property (Visio)
keywords: vis_sdr.chm11651205
f1_keywords:
- vis_sdr.chm11651205
ms.prod: visio
api_name:
- Visio.Window.Page
ms.assetid: 17474ce8-f2d7-40c7-7882-30257803c81a
ms.date: 06/08/2017
---


# Window.Page Property (Visio)

Gets or sets the page that is displayed in a window. Read/write.


## Syntax

 _expression_ . **Page**

 _expression_ A variable that represents a **Window** object.


### Return Value

Variant


## Remarks

You can set the  **Page** property to a locale-independent page name (a universal name), a locale-specific page name (a local name), or a **Page** object.

If a window is not showing a page (perhaps because it is showing a master), the  **Page** property returns **Nothing** . You can use the **Type** property of the **Window** object to determine whether the **Window** object is showing a page. Otherwise, the returned **Variant** refers to the **Page** object that the window is showing.

Beginning with Visio 5.0b, the  **Page** property no longer returns an exception if a window is not showing a page; it returns **Nothing** . You can use the following code to handle both return values:




```vb
'Close Window(intCounter) if it is showing a page. 
Set vsoWindow = Windows(intCounter) 
On Error Resume Next 
Set vsoPage = vsoWindow.Page 
 
On Error GoTo 0 
 
If Not vsoPage Is Nothing Then 
 vsoWindow.Close 
End If 

```


 **Note**  In versions of Visio through version 4.1, the  **Page** property of a **Window** object returned an **Object** (as opposed to a **Variant** of type **Object** ) and the **Page** property of a **Window** object accepted a **String** (as opposed to a **Variant** of type **String** ). Because of changes in Automation support tools, the property was changed to accept and return a **Variant** . For backward compatibility, the **PageAsObj** and **PageFromName** properties were added. The **PageAsObj** and **PageFromName** properties have the same signatures and occupy the same vtable slots as did the prior version of the **Page** property.


