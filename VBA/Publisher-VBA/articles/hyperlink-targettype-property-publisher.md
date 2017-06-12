---
title: Hyperlink.TargetType Property (Publisher)
keywords: vbapb10.chm4587529
f1_keywords:
- vbapb10.chm4587529
ms.prod: publisher
api_name:
- Publisher.Hyperlink.TargetType
ms.assetid: 1cbc8c36-563c-4464-4f0d-2836682ce532
ms.date: 06/08/2017
---


# Hyperlink.TargetType Property (Publisher)

Returns a  **PbHlinkTargetType** constant that represents the type of hyperlink. Read-only.


## Syntax

 _expression_. **TargetType**

 _expression_A variable that represents a  **Hyperlink** object.


### Return Value

PbHlinkTargetType


## Remarks

The  **TargetType** property value can be one of the following **PbHlinkTargetType** constants.



| **pbHlinkTargetTypeEmail**|
| **pbHlinkTargetTypeFirstPage**|
| **pbHlinkTargetTypeLastPage**|
| **pbHlinkTargetTypeNextPage**|
| **pbHlinkTargetTypeNone**|
| **pbHlinkTargetTypePageID**|
| **pbHlinkTargetTypePreviousPage**|
| **pbHlinkTargetTypeURL**|

## Example

This example verifies that the specified hyperlink is a URL and, if it is, sets the hyperlink display text and address. This example assumes there is at least one shape on the first page of the active publication.


```vb
Sub SetHyperlinkTextToDisplay() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Item(1) 
 If .TargetType = pbHlinkTargetTypeURL Then 
 .TextToDisplay = "Tailspin Toys Web Site" 
 .Address = "http://www.tailspintoys.com/" 
 End If 
 End With 
End Sub
```


