---
title: IDocumentInspector.GetInfo Method (Office)
ms.prod: office
api_name:
- Office.IDocumentInspector.GetInfo
ms.assetid: 7242cce4-1b36-107f-ec7c-2512b2e1fba7
ms.date: 06/08/2017
---


# IDocumentInspector.GetInfo Method (Office)

Gets information about a custom Document Inspector module.


## Syntax

 _expression_. **GetInfo**( **_Name_**, **_Desc_** )

 _expression_ An expression that returns a **IDocumentInspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|Represents the name of the module.|
| _Desc_|Required|**String**|Represents the description of the module.|

### Return Value

[HRESULT]


 **Note**  The  **IDocumentInspector** object is for the exclusive use of custom Document Inspector module authors and cannot be used with Visual Basic for Applications (VBA).


## See also


#### Concepts


[IDocumentInspector Object](idocumentinspector-object-office.md)
#### Other resources


[IDocumentInspector Object Members](idocumentinspector-members-office.md)

