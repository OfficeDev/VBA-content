---
title: Application.FilePageSetupMargins Method (Project)
keywords: vbapj.chm2356
f1_keywords:
- vbapj.chm2356
ms.prod: project-server
api_name:
- Project.Application.FilePageSetupMargins
ms.assetid: c36099a7-4ed2-0f0c-c3bb-9af35c88eb35
ms.date: 06/08/2017
---


# Application.FilePageSetupMargins Method (Project)

Sets up margins for printing.


## Syntax

 _expression_. **FilePageSetupMargins**( ** _Name_**, ** _Top_**, ** _Bottom_**, ** _Left_**, ** _Right_**, ** _Borders_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the view or report for which to set up margins for printing.|
| _Top_|Optional|**Long**|The size of the top margin in inches or centimeters.|
| _Bottom_|Optional|**Long**| The size of the bottom margin in inches or centimeters.|
| _Left_|Optional|**Long**|The size of the left margin in inches or centimeters.|
| _Right_|Optional|**Long**|The size of the right margin in inches or centimeters.|
| _Borders_|Optional|**Long**|Where to print borders. Can be one of the following  **PjBorder** constants: **pjNoBorder**, **pjAroundEveryPage**, or **pjOutsidePages**.|

### Return Value

 **Boolean**


## Remarks

Using the  **FilePageSetupMargins** method without specifying any arguments displays the **Page Setup** dialog box with the **Margins** tab selected.


