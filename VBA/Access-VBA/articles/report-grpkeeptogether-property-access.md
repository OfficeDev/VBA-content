---
title: Report.GrpKeepTogether Property (Access)
keywords: vbaac10.chm13700,vbaac10.chm4372
f1_keywords:
- vbaac10.chm13700,vbaac10.chm4372
ms.prod: access
api_name:
- Access.Report.GrpKeepTogether
ms.assetid: 605e8999-d184-b8d9-3f55-9926cd0ceefd
ms.date: 06/08/2017
---


# Report.GrpKeepTogether Property (Access)

You can use the  **GrpKeepTogether** property to specify whether groups in a multiple column report that have their **[KeepTogether](grouplevel-keeptogether-property-access.md)** property for a group set to Whole Group or With First Detail will be kept together by page or by column. Read/write **Byte**.


## Syntax

 _expression_. **GrpKeepTogether**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **GrpKeepTogether** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Per Page|0|Groups are kept together by page.|
|Per Column|1|(Default) Groups are kept together by column.|
This property can be set only in report Design view.

You can use this property to specify whether all the data for a group will appear in the same column. For example, if you have a list of employees by department in a multiple-column format, you can use this property to keep all members of the same department in the same column.

The  **GrpKeepTogether** property setting has no effect if the **KeepTogether** property for a group is set to No.


## See also


#### Concepts


[Report Object](report-object-access.md)

