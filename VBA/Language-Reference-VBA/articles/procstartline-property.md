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


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                  |
|:----------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.                                                                                                                                                                                                                                                                                         |
| <em>procname</em>     | Required. A [String](vbe-glossary.md) containing the name of the procedure.                                                                                                                                                                                                                                                                                                                   |
| <em>prockind</em>     | Required. Specifies the kind of procedure to locate. Because [property procedures](vbe-glossary.md) can have multiple representations in the[module](vbe-glossary.md), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is,  <strong>Sub</strong> and <strong>Function</strong> procedures) use <strong>vbext_pk_Proc</strong>. |

You can use one of the following [constants](vbe-glossary.md) for the _prockind_[argument](vbe-glossary.md):


| <strong>Constant</strong>      | <strong>Description</strong>                                                   |
|:-------------------------------|:-------------------------------------------------------------------------------|
| <strong>vbext_pk_Get</strong>  | Specifies a [procedure](vbe-glossary.md) that returns the value of a property. |
| <strong>vbext_pk_Let</strong>  | Specifies a procedure that assigns a value to a property.                      |
| <strong>vbext_pk_Set</strong>  | Specifies a procedure that sets a reference to an object.                      |
| <strong>vbext_pk_Proc</strong> | Specifies all procedures other than property procedures.                       |

 **Remarks**
A procedure starts at the first line below the  **End Sub** statement of the preceding procedure. If the procedure is the first procedure, it starts at the end of the general Declarations section.

