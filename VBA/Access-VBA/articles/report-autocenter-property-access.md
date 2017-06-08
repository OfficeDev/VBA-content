---
title: Report.AutoCenter Property (Access)
keywords: vbaac10.chm13797
f1_keywords:
- vbaac10.chm13797
ms.prod: access
api_name:
- Access.Report.AutoCenter
ms.assetid: d4a12dac-1000-38cd-e4ed-4f5879dfe4a0
ms.date: 06/08/2017
---


# Report.AutoCenter Property (Access)

Returns or sets a  **Boolean** indicating whether a report will be centered automatically in the application window when the form is opened. Read/write.


## Syntax

 _expression_. **AutoCenter**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **AutoCenter** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The report will be centered automatically on opening.|
|No|**False**|(Default) The report upper-left corner will be in the same location as when the form was last saved.|
You can set this property only in Design view.

Depending on the size and placement of the application window, reports can appear off to one side of the application window, hiding part of the form or report. Centering the report automatically when it's opened makes it easier to view and use.


## See also


#### Concepts


[Report Object](report-object-access.md)

