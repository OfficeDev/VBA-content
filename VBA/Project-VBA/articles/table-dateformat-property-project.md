---
title: Table.DateFormat Property (Project)
keywords: vbapj.chm132681
f1_keywords:
- vbapj.chm132681
ms.prod: project-server
api_name:
- Project.Table.DateFormat
ms.assetid: 69e0d08b-698e-8354-a583-b08122762f3f
ms.date: 06/08/2017
---


# Table.DateFormat Property (Project)

Gets or sets the date format of the table. Read/write  **PjDateFormat**.


## Syntax

 _expression_. **DateFormat**

 _expression_ A variable that represents a **Table** object.


## Remarks

The  **DateFormat** property can be one of the following **[PjDateFormat](pjdateformat-enumeration-project.md)** constants.



|**Constant**|**Date format applied to 9/30/02 (12:33 PM)**|
|:-----|:-----|
|**pjDateDefault**|The default format, as specified on the  **General** tab of the **Project Options** dialog box.|
|**pjDate_mm_dd_yy_hh_mmAM**|9/30/02 12:33 PM|
|**pjDate_mm_dd_yy**|9/30/02|
|**pjDate_mm_dd_yyyy**|9/30/2002|
|**pjDate_mmmm_dd_yyyy_hh_mmAM**|September 30, 2002 12:33 PM|
|**pjDate_mmmm_dd_yyyy**|September 30, 2002|
|**pjDate_mmm_dd_hh_mmAM**|Sep 30 12:33 PM|
|**pjDate_mmm_dd_yyy**|Sep 30, '02|
|**pjDate_mmmm_dd**|September 30|
|**pjDate_mmm_dd**|Sep 30|
|**pjDate_ddd_mm_dd_yy_hh_mmAM**|Tue 9/30/02 12:33 PM|
|**pjDate_ddd_mm_dd_yy**|Tue 9/30/02|
|**pjDate_ddd_mmm_dd_yyy**|Tue Sep 30, '02|
|**pjDate_ddd_hh_mmAM**|Tue 12:33 PM|
|**pjDate_mm_dd**|9/30|
|**pjDate_dd**|30|
|**pjDate_hh_mmAM**|12:33 PM|
|**pjDate_ddd_mmm_dd**|Tue Sep 30|
|**pjDate_ddd_mm_dd**|Tue 9/30|
|**pjDate_ddd_dd**|Tue 30|
|**pjDate_Www_dd**|W41/2|
|**pjDate_Www_dd_yy_hh_mmAM**|W41/2/02 12:33 PM|

