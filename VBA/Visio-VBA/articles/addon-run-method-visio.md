---
title: Addon.Run Method (Visio)
keywords: vis_sdr.chm12416490
f1_keywords:
- vis_sdr.chm12416490
ms.prod: visio
api_name:
- Visio.Addon.Run
ms.assetid: 223d87ff-8fd6-b68c-a716-3ff30659f898
ms.date: 06/08/2017
---


# Addon.Run Method (Visio)

Runs the add-on represented by an  **Addon** object.


## Syntax

 _expression_ . **Run**( **_ArgString_** )

 _expression_ An expression that returns a **Addon** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ArgString_|Required| **String**|The argument string to pass to the add-on.|

### Return Value

Nothing


## Remarks

If the add-on is implemented by an EXE file, the arguments are passed in the command line string. If the add-on is implemented by a VSL file, the arguments are passed in a member of the argument structure that accompanies the run message sent to the  **VisioLibMain** procedure of the VSL.


 **Security Note**  




