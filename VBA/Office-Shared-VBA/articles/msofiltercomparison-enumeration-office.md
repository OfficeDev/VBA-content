---
title: MsoFilterComparison Enumeration (Office)
ms.prod: office
api_name:
- Office.MsoFilterComparison
ms.assetid: 12650101-777b-2142-e985-cc34d5e2fb16
ms.date: 06/08/2017
---


# MsoFilterComparison Enumeration (Office)

Specifies how the  **Column** and **CompareTo** properties are compared for an **ODSOFilter** object.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**msoFilterComparisonContains**|8|Column matches CompareTo if any part of the CompareTo string is contained in the Column value.|
|**msoFilterComparisonEqual**|0|Column matches CompareTo if the CompareTo value is the same as the Column value.|
|**msoFilterComparisonGreaterThan**|3|Column matches CompareTo if the Column value is greater than the CompareTo value.|
|**msoFilterComparisonGreaterThanEqual**|5|Column matches CompareTo if the Column value is greater than or equal to the CompareTo value.|
|**msoFilterComparisonIsBlank**|6|Column passes filter if Column is blank.|
|**msoFilterComparisonIsNotBlank**|7|Column passes filter if Column is blank.|
|**msoFilterComparisonLessThan**|2|Column matches CompareTo if the Column value is less than the CompareTo value.|
|**msoFilterComparisonLessThanEqual**|4|Column matches CompareTo if the Column value is less than or equal to the CompareTo value.|
|**msoFilterComparisonNotContains**|9|Column matches CompareTo if any part of the CompareTo string is not contained in the Column value.|
|**msoFilterComparisonNotEqual**|1|Column matches CompareTo if the CompareTo value is not equal to the Column value.|

