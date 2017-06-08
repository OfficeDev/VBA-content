---
title: Application.DateFormat Method (Project)
keywords: vbapj.chm131208
f1_keywords:
- vbapj.chm131208
ms.prod: project-server
api_name:
- Project.Application.DateFormat
ms.assetid: b4fc14a0-5139-b7cf-8d96-443cd23fd8ec
ms.date: 06/08/2017
---


# Application.DateFormat Method (Project)

Returns a date in the specified format.


## Syntax

 _expression_. **DateFormat**( ** _Date_**, ** _Format_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Date_|Required|**Variant**|The date to format.|
| _Format_|Optional|**Long**|The date format. Can be one of the  **[PjDateFormat](pjdateformat-enumeration-project.md)** constants. The default value is **pjDateDefault**.|

### Return Value

 **Variant**


## Example

The following example displays the start of the selected task using the format "1/31/02 12:33 PM."


```vb
Sub OutputDate() 
 MsgBox DateFormat(ActiveCell.Task.Start, pjDate_mm_dd_yy_hh_mmAM) 
End Sub
```


