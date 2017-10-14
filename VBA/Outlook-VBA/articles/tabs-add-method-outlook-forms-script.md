---
title: Tabs.Add Method (Outlook Forms Script)
keywords: olfm10.chm2000250
f1_keywords:
- olfm10.chm2000250
ms.prod: outlook
ms.assetid: dbc72cb8-e37e-ae98-d18c-0042dc6c139f
ms.date: 06/08/2017
---


# Tabs.Add Method (Outlook Forms Script)

Adds a  **[Tab](tab-object-outlook-forms-script.md)** to a **[Tabs](tabs-object-outlook-forms-script.md)** collection.


## Syntax

 _expression_. **Add**( **_bstrName_**,  **_bstrCaption_**,  **_lIndex_**)

 _expression_A variable that represents a  **Tabs** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|bstrName|Optional| **Variant**|Specifies the name of the object being added. If a name is not specified, the system generates a default name based on the rules of the application where the form is used.|
|bstrCaption|Optional| **Variant**|Specifies the caption to appear on a tab. If a caption is not specified, the system generates a default caption based on the rules of the application where the form is used.|
|lIndex|Optional| **Variant**|Identifies the position of a tab within a  **Tabs** collection. If an index is not specified, the system appends the page to the end of the **Tabs** collection and assigns the appropriate index value.|

### Return Value

A  **Tab** object that represents the added tab.


## Remarks

The index value for the first  **Tab** of a collection is 0, the value for the second **Tab** is 1, and so on.

You can change the  **Name** property of the object at run time only if you added that control at run time with the **Add** method.


