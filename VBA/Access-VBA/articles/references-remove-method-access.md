---
title: References.Remove Method (Access)
keywords: vbaac10.chm12644
f1_keywords:
- vbaac10.chm12644
ms.prod: access
api_name:
- Access.References.Remove
ms.assetid: ebdc9da2-cc32-6169-994a-1041b1c49031
ms.date: 06/08/2017
---


# References.Remove Method (Access)

The  **Remove** method removes a **[Reference](reference-object-access.md)** object from the **[References](references-object-access.md)** collection.


## Syntax

 _expression_. **Remove**( ** _Reference_** )

 _expression_ A variable that represents a **References** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Reference_|Required|**Reference**|The  **Reference** object that represents the reference you wish to remove.|

## Remarks

To determine the name of the  **Reference** object you wish to remove, check the Project/Library box in the Object Browser. The names of all references that are currently set appear there. These names correspond to the value of the **Name** property of a **Reference** object.


## See also


#### Concepts


[References Collection](references-object-access.md)

