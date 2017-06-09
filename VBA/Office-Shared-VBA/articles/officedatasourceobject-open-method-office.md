---
title: OfficeDataSourceObject.Open Method (Office)
keywords: vbaof11.chm232007
f1_keywords:
- vbaof11.chm232007
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.Open
ms.assetid: ef01fe38-68ad-6dfb-fcf6-2bd06d308acc
ms.date: 06/08/2017
---


# OfficeDataSourceObject.Open Method (Office)

Opens a table in a  **OfficeDataSourceObject** object.


## Syntax

 _expression_. **Open**( **_bstrSrc_**, **_bstrConnect_**, **_bstrTable_**, **_fOpenExclusive_**, **_fNeverPrompt_** )

 _expression_ A variable that represents an **OfficeDataSourceObject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrSrc_|Optional|**String**|Contains the name of the data source.|
| _bstrConnect_|Optional|**String**|Contains the connection string to the data source.|
| _bstrTable_|Optional|**String**|Specifies which table to open.|
| _fOpenExclusive_|Optional|**Long**|Indicates whether the table should be opened for exclusive access.|
| _fNeverPrompt_|Optional|**Long**|Indicates whether to notify the user if the table can not be opened.|

## See also


#### Concepts


[OfficeDataSourceObject Object](officedatasourceobject-object-office.md)
#### Other resources


[OfficeDataSourceObject Object Members](officedatasourceobject-members-office.md)

