---
title: Form.AllowFilters Property (Access)
keywords: vbaac10.chm13350,vbaac10.chm4262
f1_keywords:
- vbaac10.chm13350,vbaac10.chm4262
ms.prod: access
api_name:
- Access.Form.AllowFilters
ms.assetid: ca2998b5-d5e0-f1ba-f9da-d89ef24a3701
ms.date: 06/08/2017
---


# Form.AllowFilters Property (Access)

You can use the  **AllowFilters** property to specify whether records in a form can be filtered. Read/write **Boolean**.


## Syntax

 _expression_. **AllowFilters**

 _expression_ A variable that represents a **Form** object.


## Remarks

Filters are commonly used to view a temporary subset of the records in a database. When you use a filter, you apply criteria to display only records that meet specific conditions. In an Employees form, for example, you can use a filter to display only records of employees with over 5 years of service. You can also use a filter to restrict access to records containing sensitive information, such as financial or medical data.




 **Note**  Setting the  **AllowFilters** property to No does not affect the **[Filter](form-filter-property-access.md)**, **[FilterOn](form-filteron-property-access.md)**, **[ServerFilter](form-serverfilter-property-access.md)**, or **[ServerFilterByForm](form-serverfilterbyform-property-access.md)** properties. You can still use these properties to set and remove filters. You can also still use the following actions or methods to apply and remove filters.



|**Actions**|**Methods**|
|:-----|:-----|
|ApplyFilter|**[ApplyFilter](docmd-applyfilter-method-access.md)**|
|OpenForm|**[OpenForm](docmd-openform-method-access.md)**|
|ShowAllRecords|**[ShowAllRecords](docmd-showallrecords-method-access.md)**|

## See also


#### Concepts


[Form Object](form-object-access.md)

