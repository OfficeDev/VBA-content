---
title: Report.Section Property (Access)
keywords: vbaac10.chm13874
f1_keywords:
- vbaac10.chm13874
ms.prod: access
api_name:
- Access.Report.Section
ms.assetid: 3baad974-8869-30b5-abe3-8cf754a225b3
ms.date: 06/08/2017
---


# Report.Section Property (Access)

You can use the  **Section** property to identify a section of a report and provide access to the properties of that section. Read-only **Section** object.


## Syntax

 _expression_. **Section**( ** _Index_** )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The section number or name.|

## Remarks

The  **Section** property corresponds to a particular section. You can use the following constants listed below. It is recommended that you use the constants to make your code easier to read.



|**Setting**|**Constant**|**Description**|
|:-----|:-----|:-----|
|0|**acDetail**|Report detail section|
|1|**acHeader**|Report header section|
|2|**acFooter**|Report footer section|
|3|**acPageHeader**|Report page header section|
|4|**acPageFooter**|Report page footer section|
|5|**acGroupLevel1Header**|Group-level 1 header section |
|6|**acGroupLevel1Footer**|Group-level 1 footer section|
|7|**acGroupLevel2Header**|Group-level 2 header section|
|8|**acGroupLevel2Footer**|Group-level 2 footer section|

 **Note**  If a report has additional group-level sections, the header/footer pairs are numbered consecutively beginning with 9.

The  **Section** property is an array of all existing sections in the form or report specified by the section number. For example, `Section(0)` refers to a form's detail section and `Section(3)` refers to a form's page header section.

You can also refer to a section by name. The following statements refer to the Detail0 section for the Customers form and are equivalent.




```vb
Forms!Customers.Section(acDetail).Visible
```




```vb
Forms!Customers.Section(0).Visible
```




```vb
Forms!Customers.Detail0.Visible
```

You must combine the  **Section** property with other properties that apply to form or report sections.


## See also


#### Concepts


[Report Object](report-object-access.md)

