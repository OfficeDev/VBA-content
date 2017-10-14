---
title: ProcOfLine Property
keywords: vbob6.chm104019
f1_keywords:
- vbob6.chm104019
ms.prod: office
api_name:
- Office.ProcOfLine
ms.assetid: daf7ffbf-41a8-aacb-e9ef-c576efd3d11c
ms.date: 06/08/2017
---


# ProcOfLine Property



Returns the name of the [procedure](vbe-glossary.md) that the specified line is in.
 **Syntax**
 _object_**.ProcOfLine(**_line_, _prockind_**) As String**
The  **ProcOfLine** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _line_|Required. A [Long](vbe-glossary.md) specifying the line to check.|
| _prockind_|Required. Specifies the kind of procedure to locate. Because [property procedures](vbe-glossary.md) can have multiple representations in the[module](vbe-glossary.md), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is,  **Sub** and **Function** procedures) use **vbext_pk_Proc**.|
You can use one of the following [constants](vbe-glossary.md) for the _prockind_[argument](vbe-glossary.md):


|**Constant**|**Description**|
|:-----|:-----|
|**vbext_pk_Get**|Specifies a procedure that returns the value of a property.|
|**vbext_pk_Let**|Specifies a procedure that assigns a value to a property.|
|**vbext_pk_Set**|Specifies a procedure that sets a reference to an object.|
|**vbext_pk_Proc**|Specifies all procedures other than property procedures.|
 **Remarks**
A line is within a procedure if it's a blank line or comment line preceding the procedure declaration and, if the procedure is the last procedure in a [code module](vbe-glossary.md), a blank line or lines following the procedure.

