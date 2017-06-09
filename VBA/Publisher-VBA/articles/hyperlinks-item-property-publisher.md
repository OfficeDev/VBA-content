---
title: Hyperlinks.Item Property (Publisher)
keywords: vbapb10.chm6881280
f1_keywords:
- vbapb10.chm6881280
ms.prod: publisher
api_name:
- Publisher.Hyperlinks.Item
ms.assetid: 8d288fc6-9ded-5732-b972-6fa366ef31c3
ms.date: 06/08/2017
---


# Hyperlinks.Item Property (Publisher)

Returns an individual object from a specified collection. Read-only.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **Hyperlinks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The number of the object to return.|

## Example

This example displays the address of the first hyperlink in shape one of the active publication.


```vb
MsgBox "Address of first hyperlink: " _ 
 &; ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Item(1).Address
```


