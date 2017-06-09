---
title: Worksheet.XmlDataQuery Method (Excel)
keywords: vbaxl10.chm175158
f1_keywords:
- vbaxl10.chm175158
ms.prod: excel
api_name:
- Excel.Worksheet.XmlDataQuery
ms.assetid: de728702-962f-a047-a58d-3e2fa9c86acd
ms.date: 06/08/2017
---


# Worksheet.XmlDataQuery Method (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the cells mapped to a particular XPath. Returns **Nothing** if the specified XPath has not been mapped to the worksheet, or if the mapped range is empty.


## Syntax

 _expression_ . **XmlDataQuery**( **_XPath_** , **_SelectionNamespaces_** , **_Map_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPath_|Required| **String**|The XPath to query for.|
| _SelectionNamespaces_|Optional| **Variant**|A space-delimited  **String** that contains the namespaces referenced in the XPath parameter. A run-time error will be generated if one of the specified namespaces cannot be resolved.|
| _Map_|Optional| **Variant**|Specify an  **[XmlMap](xmlmap-object-excel.md)** if you want to query for the XPath within a specific map.|

### Return Value

Range


## Remarks

If the XPath exists within a column in an XML list, the  **Range** object returned does not include the header row.

This method returns  **Nothing** if the XPath location path is not mapped into the grid. Thus, a return of **Nothing** doesn't necessarily mean that the map doesn't exist. It could mean that there is currently no data range available at the specified XPath location. You can use the **[XmlMapQuery](worksheet-xmlmapquery-method-excel.md)** method to check for the existence of a mapped XPath.


 **Note**  The  **XmlDataQuery** method allows you to query for the existence of particular map data. It can not be used to query for a piece of data in a map. For example, it is valid for a mapped range to exist in which the XPath for that range is "/root/People[@Age="23"]/FirstName". An XmlDataQuery query for this XPath location path will returnreturns the correct range. However, a query for "/root/People[FirstName="Joe"]" hoping to find "Joe" within the above mapped range will fail because the XPath definitions for the mapped ranges are different.


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

