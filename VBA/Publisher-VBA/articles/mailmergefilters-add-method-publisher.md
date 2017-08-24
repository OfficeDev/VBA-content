---
title: MailMergeFilters.Add Method (Publisher)
keywords: vbapb10.chm6750212
f1_keywords:
- vbapb10.chm6750212
ms.prod: publisher
api_name:
- Publisher.MailMergeFilters.Add
ms.assetid: ab114dda-d144-7c5f-88b0-930cadcf53db
ms.date: 06/08/2017
---


# MailMergeFilters.Add Method (Publisher)

Adds a new filter criterion to the specified  **MailMergeFilters** object.


## Syntax

 _expression_. **Add**( **_Column_**,  **_Comparison_**,  **_Conjunction_**,  **_bstrCompareTo_**,  **_DeferUpdate_**)

 _expression_A variable that represents a  **MailMergeFilters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Column|Required| **String**|The name of the table in the data source.|
|Comparison|Required| **MsoFilterComparison**|How the data in the table is filtered.|
|Conjunction|Required| **MsoFilterConjunction**| How this filter relates to other filters in the **MailMergeFilters** object.|
|bstrCompareTo|Optional| **String**|If the  **Comparison** argument is something other than **msoFilterComparisonIsBlank** or **msoFilterComparisonIsNotBlank**, a string to which the data in the table is compared.|
|DeferUpdate|Optional| **Boolean**| **True** to queue the filters and apply them when the **ApplyFilter** method is called. **False** to apply the filter condition immediately. Default is **False**.|

## Remarks

Comparison can be one of these  **MsoFilterComparison** constants.



| **msoFilterComparisonContains**|
| **msoFilterComparisonEqual**|
| **msoFilterComparisonGreaterThan**|
| **msoFilterComparisonGreaterThanEqual**|
| **msoFilterComparisonIsBlank**|
| **msoFilterComparisonIsNotBlank**|
| **msoFilterComparisonLessThan**|
| **msoFilterComparisonLessThanEqual**|
| **msoFilterComparisonNotContains**|
| **msoFilterComparisonNotEqual**|
Conjunction can be one of these  **MsoFilterConjunction** constants.



| **msoFilterConjunctionAnd**|
| **msoFilterConjunctionOr**|

