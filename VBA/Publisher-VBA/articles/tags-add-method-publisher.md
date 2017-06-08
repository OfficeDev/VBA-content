---
title: Tags.Add Method (Publisher)
keywords: vbapb10.chm4653060
f1_keywords:
- vbapb10.chm4653060
ms.prod: publisher
api_name:
- Publisher.Tags.Add
ms.assetid: 78602ccc-8198-1183-4775-fe626eb8b5af
ms.date: 06/08/2017
---


# Tags.Add Method (Publisher)

Adds a new  **Tag** object to the specified **Tags** object and returns the new **Tag** object.


## Syntax

 _expression_. **Add**( **_Name_**,  **_Value_**)

 _expression_A variable that represents a  **Tags** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|Required| **String**|The name of the tag to add. If a tag already exists with the same name, an error occurs.|
|Value|Required| **Variant**|The value to assign to the tag.|

### Return Value

Tag


## Example

The following example adds a tag to shape one on page one of the active publication.


```vb
Dim tagNew As Tag 
 
Set tagNew = ActiveDocument.Pages(1).Shapes(1).Tags _ 
 .Add(Name:="required", Value:="yes")
```


