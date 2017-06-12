---
title: AddIn.Loaded Property (PowerPoint)
keywords: vbapp10.chm521008
f1_keywords:
- vbapp10.chm521008
ms.prod: powerpoint
api_name:
- PowerPoint.AddIn.Loaded
ms.assetid: 8becb17d-dbe4-b151-e66b-3463f3a862f5
ms.date: 06/08/2017
---


# AddIn.Loaded Property (PowerPoint)

Determines whether the specified add-in is loaded. Read/write.


## Syntax

 _expression_. **Loaded**

 _expression_ A variable that represents an **AddIn** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Loaded** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|TThe specified add-in is not loaded. |
|**msoTrue**| The specified add-in is loaded.|

## Example

This example adds MyTools.ppa to the list in the  **Add-Ins** tab and then loads it.


```
Addins.Add("c:\my documents\mytools.ppa").Loaded = msoTrue
```

This example unloads the add-in named "MyTools."




```vb
Application.Addins("mytools").Loaded = msoFalse
```


## See also


#### Concepts


[AddIn Object](addin-object-powerpoint.md)

