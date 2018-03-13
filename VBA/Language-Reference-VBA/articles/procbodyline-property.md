---
title: ProcBodyLine Property
keywords: vbob6.chm104018
f1_keywords:
- vbob6.chm104018
ms.prod: office
api_name:
- Office.ProcBodyLine
ms.assetid: 63169755-41db-fd3a-a3f4-87efa0739d38
ms.date: 06/08/2017
---


# ProcBodyLine Property



Returns the first line of a [procedure](vbe-glossary.md).
 **Syntax**
 _object_**.ProcBodyLine(**_procname_, _prockind_**) As Long**
The  **ProcBodyLine** syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                  |
|:----------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.                                                                                                                                                                                                                                                                                         |
| <em>procname</em>     | Required. A [String](vbe-glossary.md) containing the name of the procedure.                                                                                                                                                                                                                                                                                                                   |
| <em>prockind</em>     | Required. Specifies the kind of procedure to locate. Because [property procedures](vbe-glossary.md) can have multiple representations in the[module](vbe-glossary.md), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is,  <strong>Sub</strong> and <strong>Function</strong> procedures) use <strong>vbext_pk_Proc</strong>. |

You can use one of the following [constants](vbe-glossary.md) for the _prockind_[argument](vbe-glossary.md):


| <strong>Constant</strong>      | <strong>Description</strong>                                |
|:-------------------------------|:------------------------------------------------------------|
| <strong>vbext_pk_Get</strong>  | Specifies a procedure that returns the value of a property. |
| <strong>vbext_pk_Let</strong>  | Specifies a procedure that assigns a value to a property.   |
| <strong>vbext_pk_Set</strong>  | Specifies a procedure that sets a reference to an object.   |
| <strong>vbext_pk_Proc</strong> | Specifies all procedures other than property procedures.    |

 **Remarks**
The first line of a procedure is the line on which the  **Sub**, **Function**, or **Property** statement appears.

