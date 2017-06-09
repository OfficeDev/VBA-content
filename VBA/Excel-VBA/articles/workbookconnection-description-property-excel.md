---
title: WorkbookConnection.Description Property (Excel)
keywords: vbaxl10.chm774074
f1_keywords:
- vbaxl10.chm774074
ms.prod: excel
api_name:
- Excel.WorkbookConnection.Description
ms.assetid: a0ba84a8-6bea-71aa-92be-2d875ec23a42
ms.date: 06/08/2017
---


# WorkbookConnection.Description Property (Excel)

Returns or sets a brief description for a  **WorkbookConnection** object. Read/write **String** .


## Syntax

 _expression_ . **Description**

 _expression_ A variable that represents a **WorkbookConnection** object.


## Remarks

In the  **Connection Properties** dialog box, the user may edit the name of the connection and/or the description. Changing the name and description in this dialog box changes those fields only within the Excel connection object.

The maximum size of a description is 255 characters. If the user specifies a description within a connection file that is longer than 255 characters, the description is truncate to fit the 255 character limit.


## See also


#### Concepts


[WorkbookConnection Object](workbookconnection-object-excel.md)

