---
title: Application.TableCopy Method (Project)
keywords: vbapj.chm400
f1_keywords:
- vbapj.chm400
ms.prod: project-server
api_name:
- Project.Application.TableCopy
ms.assetid: 90e0a546-2802-5ba7-6b49-086b32051451
ms.date: 06/08/2017
---


# Application.TableCopy Method (Project)

Makes a copy of the active table, adds it to the  **Tables** drop-down menu, and sets the view to use the new table.


## Syntax

 _expression_. **TableCopy**( ** _Name_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|Name of the copied table.|

### Return Value

 **Boolean**


## Remarks

The  **Tables** drop-down menu is on the **View** tab on the ribbon. If you run the **TableCopy** method without specifying the _Name_ argument, Project displays the **Save Table** dialog box.


 **Note**  The  **TableCopy** action is not stored in the **Undo** list.

For detailed control of table features when making a copy, see the  **[TableEditEx](application-tableeditex-method-project.md)** method.


## Example

If the active view is the Resource Sheet, the following statement copies the resource Entry table to a table named "Copy of Resource Sheet table" and then sets the Resource Sheet view to use that table.


```
TableCopy Name:="Copy of Resource Sheet table"
```


