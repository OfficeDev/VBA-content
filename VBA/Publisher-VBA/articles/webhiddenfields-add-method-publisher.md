---
title: WebHiddenFields.Add Method (Publisher)
keywords: vbapb10.chm3997700
f1_keywords:
- vbapb10.chm3997700
ms.prod: publisher
api_name:
- Publisher.WebHiddenFields.Add
ms.assetid: c3035138-f369-b561-b1f8-9977bd9e080c
ms.date: 06/08/2017
---


# WebHiddenFields.Add Method (Publisher)

Adds a new hidden field to a Web form and returns a  **Long** indicating the number of the new field in the **WebHiddenFields** collection. New fields are always placed at the end of the current field list.


## Syntax

 _expression_. **Add**( **_Name_**,  **_Value_**)

 _expression_A variable that represents a  **WebHiddenFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|Required| **String**|The name of the new field.|
|Value|Required| **String**|The value of the new field.|

### Return Value

Long


## Example

The following example adds a new hidden field to the specified Web command button control. Shape one on page one of the active publication must be a Web command button control for this example to work.


```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .WebCommandButton.HiddenFields _ 
 .Add Name:="subject", Value:="service request"
```


