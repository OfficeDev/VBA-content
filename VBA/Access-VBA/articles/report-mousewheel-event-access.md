---
title: Report.MouseWheel Event (Access)
keywords: vbaac10.chm13900
f1_keywords:
- vbaac10.chm13900
ms.prod: access
api_name:
- Access.Report.MouseWheel
ms.assetid: 9c234923-3459-c45e-8489-146353f59c21
ms.date: 06/08/2017
---


# Report.MouseWheel Event (Access)

Occurs when the user rolls the mouse wheel in Report view or Layout view.


## Syntax

 _expression_. **MouseWheel**( ** _Page_**, ** _Count_** )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Page_|Required|**Boolean**|**True** if the page was changed.|
| _Count_|Required|**Long**|The number of lines by which the view was scrolled with the mouse wheel.|

## See also


#### Concepts


[Report Object](report-object-access.md)

