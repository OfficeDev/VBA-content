---
title: Tabs.Remove Method (Outlook Forms Script)
keywords: olfm10.chm2000360
f1_keywords:
- olfm10.chm2000360
ms.prod: outlook
ms.assetid: f0fa694c-112a-b85f-b1c8-74b935fe2609
ms.date: 06/08/2017
---


# Tabs.Remove Method (Outlook Forms Script)

Removes a member from a collection.


## Syntax

 _expression_. **Remove**( **_varg_**)

 _expression_A variable that represents a  **Tabs** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|varg|Required| **Variant**|A member's position, or index, within a collection. Numeric as well as string values are acceptable. If the value is a number, the minimum value is zero, and the maximum value is one less than the number of members in the collection. If the value is a string, it must correspond to a valid member name.|

