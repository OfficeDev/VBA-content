---
title: Form.Section Property (Access)
keywords: vbaac10.chm13631
f1_keywords:
- vbaac10.chm13631
ms.prod: access
api_name:
- Access.Form.Section
ms.assetid: df8d00af-3e1e-86f8-17f4-dd5792193d03
ms.date: 06/08/2017
---


# Form.Section Property (Access)

You can use the  **Section** property to identify a section of a form and provide access to the properties of that section. Read-only **Section** object.


## Syntax

 _expression_. **Section**( ** _Index_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The section number or name.|

## Remarks

The  **Section** property corresponds to a particular section. You can use the following constants listed below. It is recommended that you use the constants to make your code easier to read.



|**Setting**|**Constant**|**Description**|
|:-----|:-----|:-----|
|0|**acDetail**|Form detail section|
|1|**acHeader**|Form header section|
|2|**acFooter**|Form footer section|
|3|**acPageHeader**|Form page header section|
|4|**acPageFooter**|Form page footer section|
For forms and reports, the  **Section** property is an array of all existing sections in the form specified by the section number. For example, `Section(0)` refers to a form's detail section and `Section(3)` refers to a form's page header section.

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

For forms and reports, you must combine the  **Section** property with other properties that apply to form or report sections.


## Example

The following example shows how to refer to the Visible property of the page header section of the Customers form.


```vb
Forms!Customers.Section(acPageHeader).Visible 
Forms!Customers.Section(3).Visible
```


## See also


#### Concepts


[Form Object](form-object-access.md)

