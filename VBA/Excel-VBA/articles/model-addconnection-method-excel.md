---
title: Model.AddConnection Method (Excel)
ms.prod: excel
ms.assetid: 58ed2796-9cfa-2737-43c0-f5a5a4badcc3
ms.date: 06/08/2017
---


# Model.AddConnection Method (Excel)

Adds a new Workbook Connection to the model with the same properties as the one supplied as an argument.


## Syntax

 _expression_ . **AddConnection**_(ConnectionToDataSource)_

 _expression_ A variable that represents a object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ConnectionToDataSource_|Required|WORKBOOKCONNECTION|The Workbook connection.|

### Return value

 **WORKBOOKCONNECTION**


## Remarks

This method only works on legacy/non-model external connections and will fail with a run-time error if called with an external model connection as its argument. When calling this method, a new model connection is created and it is named the same as the legacy connection with the existing logic for making the name unique applied (integer at the end).


## See also


#### Other resources


[Model Object Members](http://msdn.microsoft.com/library/b2bd944a-3484-222b-b3d6-acd70a6ac28a%28Office.15%29.aspx)


