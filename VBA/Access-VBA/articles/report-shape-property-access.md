---
title: Report.Shape Property (Access)
keywords: vbaac10.chm13815
f1_keywords:
- vbaac10.chm13815
ms.prod: access
ms.assetid: 1fd2eb3c-5473-b239-d0c6-4e0ded950df6
ms.date: 06/08/2017
---


# Report.Shape Property (Access)

Returns a  **String** representing the shape command corresponding to the sorting and grouping of the specified report. Read-only.


## Syntax

 _expression_. **Shape**

 _expression_ A variable that represents a **Report** object.


## Remarks

Analyzing the shape command for a particular report can make it easier to create ADO recordsets that can be used with the report.


## Example

The following example reads the shape command of the specified report and stores it to a string variable.


```vb
Dim strShape As String 
 
strShape = Reports(0).Shape

```


## Property value

 **STRING**


