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


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                  |
|:----------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.                                                                                                                                                                                                                                                                                         |
| <em>line</em>         | Required. A [Long](vbe-glossary.md) specifying the line to check.                                                                                                                                                                                                                                                                                                                             |
| <em>prockind</em>     | Required. Specifies the kind of procedure to locate. Because [property procedures](vbe-glossary.md) can have multiple representations in the[module](vbe-glossary.md), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is,  <strong>Sub</strong> and <strong>Function</strong> procedures) use <strong>vbext_pk_Proc</strong>. |

You can use one of the following [constants](vbe-glossary.md) for the _prockind_[argument](vbe-glossary.md):


| <strong>Constant</strong>      | <strong>Description</strong>                                |
|:-------------------------------|:------------------------------------------------------------|
| <strong>vbext_pk_Get</strong>  | Specifies a procedure that returns the value of a property. |
| <strong>vbext_pk_Let</strong>  | Specifies a procedure that assigns a value to a property.   |
| <strong>vbext_pk_Set</strong>  | Specifies a procedure that sets a reference to an object.   |
| <strong>vbext_pk_Proc</strong> | Specifies all procedures other than property procedures.    |

 **Remarks**
A line is within a procedure if it's a blank line or comment line preceding the procedure declaration and, if the procedure is the last procedure in a [code module](vbe-glossary.md), a blank line or lines following the procedure.

