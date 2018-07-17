---
title: AddIn.Registered Property (PowerPoint)
keywords: vbapp10.chm521006
f1_keywords:
- vbapp10.chm521006
ms.prod: powerpoint
api_name:
- PowerPoint.AddIn.Registered
ms.assetid: 693bcb7a-dabc-5933-38df-710172bbce26
ms.date: 06/08/2017
---


# AddIn.Registered Property (PowerPoint)

Determines whether the specified add-in is registered in the Windows registry. Read/write.


## Syntax

 _expression_. **Registered**

 _expression_ A variable that represents a **AddIn** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Registered** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The specified add-in is not registered in the Windows registry.|
|**msoTrue**| The specified add-in is registered in the Windows registry.|

## Example

This example registers the add-in named "MyTools" in the Windows registry.


```vb
Application.Addins("MyTools").Registered = msoTrue
```


## See also


#### Concepts


[AddIn Object](addin-object-powerpoint.md)

