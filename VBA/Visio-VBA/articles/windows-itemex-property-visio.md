---
title: Windows.ItemEx Property (Visio)
keywords: vis_sdr.chm11751730
f1_keywords:
- vis_sdr.chm11751730
ms.prod: visio
api_name:
- Visio.Windows.ItemEx
ms.assetid: 24adeef0-20ca-4e00-ff39-c49ec5e72f87
ms.date: 06/08/2017
---


# Windows.ItemEx Property (Visio)

Returns a  **Window** object from a collection. Read-only.


## Syntax

 _expression_ . **ItemEx**( **_CaptionOrIndex_** )

 _expression_ A variable that represents a **Windows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CaptionOrIndex_|Required| **Variant**|Contains the caption or index of the window to retrieve. See Remarks for details.|

### Return Value

Window


## Remarks

The  **ItemEx** property is similar to the **Item** property as it applies to the **Windows** collection, except that the first argument can be either the window caption or the index. Beginning with Microsoft Office Visio 2003, all built-in Multiple Document Interface (MDI) windows have unique captions, although there is no guarantee that subwindows have unique captions. If there are multiple subwindows that have the same caption, **ItemEx** returns the one that has the lowest index.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ItemEx** property to make a window active in Visio. It adds a document to the **Documents** collection, thereby creating a new window. Then it gets the index number of the new window (which is equal to the count of window items), uses that index number to get the new window's caption, and then passes the caption to the **ItemEx** property to activate the new window.


```vb
Sub ItemEx_Example() 
 
 Dim intWindowCount As Integer 
 Dim strWindowCaption As String 
 
 'Add a document not based on a template to the collection 
 Application.Documents.Add ("") 
 
 'Get the index number in the Windows collection of the new window 
 intWindowCount = Application.Windows.Count 
 
 'Get the new window's caption 
 strWindowCaption = Application.Windows(intWindowCount) 
 
 'Activate the new window 
 Application.Windows.ItemEx(strWindowCaption).Activate 
 
End Sub
```


