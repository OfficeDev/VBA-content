---
title: OfficeDataSourceObject.Move Method (Office)
keywords: vbaof11.chm232006
f1_keywords:
- vbaof11.chm232006
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.Move
ms.assetid: cf732e6c-58b3-94a7-5081-3f1350800fd0
ms.date: 06/08/2017
---


# OfficeDataSourceObject.Move Method (Office)

Moves a record in a return set from a  **OfficeDataSourceObject** object from one position to another.


## Syntax

 _expression_. **Move**( **_MsoMoveRow_**, **_RowNbr_** )

 _expression_ A variable that represents an **OfficeDataSourceObject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MsoMoveRow_|Required|**MsoMoveRow**|A constant specifying which row to move.|
| _RowNbr_|Optional|**Integer**|The number of the destination row.|

### Return Value

Integer


## Example

The following example moves the first row in the set of records to the third row.


```
oOdso.Move(msoMoveRowFirst, 3)
```


## See also


#### Concepts


[OfficeDataSourceObject Object](officedatasourceobject-object-office.md)
#### Other resources


[OfficeDataSourceObject Object Members](officedatasourceobject-members-office.md)

