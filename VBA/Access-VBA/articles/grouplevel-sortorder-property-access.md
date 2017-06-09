---
title: GroupLevel.SortOrder Property (Access)
keywords: vbaac10.chm12240
f1_keywords:
- vbaac10.chm12240
ms.prod: access
api_name:
- Access.GroupLevel.SortOrder
ms.assetid: 2c58785c-4ddb-a581-b438-5f6390f544dd
ms.date: 06/08/2017
---


# GroupLevel.SortOrder Property (Access)

You use the  **SortOrder** property to specify the sort order for fields and expressions in a report. For example, if you're printing a list of suppliers, you can sort the records alphabetically by company name. Read/write **Boolean**.


## Syntax

 _expression_. **SortOrder**

 _expression_ A variable that represents a **GroupLevel** object.


## Remarks

The  **SortOrder** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Ascending|**False**|(Default) Sorts values in ascending (A to Z, 0 to 9) order.|
|Descending|**True**|Sorts values in descending (Z to A, 9 to 0) order.|
In Visual Basic, you set the  **SortOrder** property in report Design view or in the **Open** event procedure of a report by using the **[GroupLevel](report-grouplevel-property-access.md)** property.


## Example

The following example sets the sort order to ascending for the first group level in the "Product Summary" report.


```vb
Reports("Product Summary").GroupLevel(0).SortOrder = False 

```


## See also


#### Concepts


[GroupLevel Object](grouplevel-object-access.md)

