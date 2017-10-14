---
title: ProcStartLine Property
keywords: vbob6.chm104016
f1_keywords:
- vbob6.chm104016
ms.prod: office
api_name:
- Office.ProcStartLine
ms.assetid: 1a28f3e2-77a3-709a-5229-7a1a2d5afa48
ms.date: 06/08/2017
---


# ProcStartLine Property



Returns the line at which the specified [procedure](vbe-glossary.md) begins.
 **Syntax**
 _object_**.ProcStartLine(**_procname_, _prockind_**) As Long**
The  **ProcStartLine** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _procname_|Required. A [String](vbe-glossary.md) containing the name of the procedure.|
| _prockind_|Required. Specifies the kind of procedure to locate. Because [property procedures](vbe-glossary.md) can have multiple representations in the[module](vbe-glossary.md), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is,  **Sub** and **Function** procedures) use **vbext_pk_Proc**.|
You can use one of the following [constants](vbe-glossary.md) for the _prockind_[argument](vbe-glossary.md):


|**Constant**|**Description**|
|:-----|:-----|
|**vbext_pk_Get**|Specifies a [procedure](vbe-glossary.md) that returns the value of a property.|
|**vbext_pk_Let**|Specifies a procedure that assigns a value to a property.|
|**vbext_pk_Set**|Specifies a procedure that sets a reference to an object.|
|**vbext_pk_Proc**|Specifies all procedures other than property procedures.|
 **Remarks**
A procedure starts at the first line below the  **End Sub** statement of the preceding procedure. If the procedure is the first procedure, it starts at the end of the general Declarations section.

