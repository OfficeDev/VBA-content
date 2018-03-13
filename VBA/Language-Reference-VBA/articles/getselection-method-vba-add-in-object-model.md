---
title: GetSelection Method (VBA Add-In Object Model)
keywords: vbob6.chm1098973
f1_keywords:
- vbob6.chm1098973
ms.prod: office
ms.assetid: f7275ba1-85a3-4939-2ab2-f39e750623f0
ms.date: 06/08/2017
---


# GetSelection Method (VBA Add-In Object Model)



Returns the selection in a [code pane](vbe-glossary.md).
 **Syntax**
 _object_**.GetSelection(**_startline_, _startcol_, _endline_, _endcol_**)**
The  **GetSelection** syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                           |
|:----------------------|:-----------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.                  |
| <em>startline</em>    | Required. A [Long](vbe-glossary.md) that returns a value specifying the first line of the selection in the code pane.  |
| <em>startcol</em>     | Required. A  <strong>Long</strong> that returns a value specifying the first column of the selection in the code pane. |
| <em>endline</em>      | Required. A  <strong>Long</strong> that returns a value specifying the last line of the selection in the code pane.    |
| <em>endcol</em>       | Required. A  <strong>Long</strong> that returns a value specifying the last column of the selection in the code pane.  |

 **Remarks**
When you use the  **GetSelection** method, information is returned in output[arguments](vbe-glossary.md). As a result, you must pass in [variables](vbe-glossary.md) because the variables will be modified to contain the information when returned.

