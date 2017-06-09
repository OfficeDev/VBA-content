---
title: Module.AddFromFile Method (Access)
keywords: vbaac10.chm12274
f1_keywords:
- vbaac10.chm12274
ms.prod: access
api_name:
- Access.Module.AddFromFile
ms.assetid: a782b4dc-a4be-5166-3ce3-deb87ed1195b
ms.date: 06/08/2017
---


# Module.AddFromFile Method (Access)

The  **AddFromFile** method adds the contents of a text file to a **Module** object. The **Module** object may represent a standard module or a class module.


## Syntax

 _expression_. **AddFromFile**( ** _FileName_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name and full path of a text (.txt) file or another file that stores text in an ANSI format.|

## Remarks

The  **AddFromFile** method places the contents of the specified text file immediately after the Declarations section and before the first procedure in the module if it contains other procedures.

The  **AddFromFile** method enables you to import code or comments stored in a text file.

In order to add the contents of a file to a form or report module, the form or report must be open in form Design view or report Design view. In order to add the contents of a file to a standard module or class module, the module must be open.


## Example

The following example places the contents of the file "ShippingRoutines.bas" into the module "CalculateShipping" immediately after the Declarations section, but before the first procedure in the module.


```vb
Modules("CalculateShipping").AddFromFile "C:\Shipping\ShippingRoutines.bas" 

```


## See also


#### Concepts


[Module Object](module-object-access.md)

