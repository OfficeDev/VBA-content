---
title: DocumentProperty.LinkSource Property (Office)
keywords: vbaof11.chm250009
f1_keywords:
- vbaof11.chm250009
ms.prod: office
api_name:
- Office.DocumentProperty.LinkSource
ms.assetid: 3e3a6ebc-615a-298e-c40f-cbb6d5cf63e3
ms.date: 06/08/2017
---


# DocumentProperty.LinkSource Property (Office)

Gets or sets the source of a linked custom document property. Read/write.


## Syntax

 _expression_. **LinkSource**( **_pbstrSourceRetVal_** )

 _expression_ A variable that represents a **DocumentProperty** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pbstrSourceRetVal_|Required|**String**|Represents the name of the source of the document property.|

## Remarks

This property applies only to custom document properties; you cannot use it with built-in document properties.

The source of the specified link is defined by the container application.

Setting the  **LinkSource** property sets the **LinkToContent** property to **True**.


## Example

This example displays the linked status of a custom document property. For the example to work,  **dp** must be a valid **DocumentProperty** object.


```
Sub DisplayLinkStatus(dp As DocumentProperty) 
 Dim stat As String, tf As String 
 If dp.LinkToContent Then 
 tf = "" 
 Else 
 tf = "not " 
 End If 
 stat = "This property is " &amp; tf &amp; "linked" 
 If dp.LinkToContent Then 
 stat = stat + Chr(13) &amp; "The link source is " &amp; dp.LinkSource 
 End If 
 MsgBox stat 
End Sub
```


## See also


#### Concepts


[DocumentProperty Object](documentproperty-object-office.md)
#### Other resources


[DocumentProperty Object Members](documentproperty-members-office.md)

