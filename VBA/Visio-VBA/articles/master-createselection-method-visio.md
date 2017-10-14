---
title: Master.CreateSelection Method (Visio)
keywords: vis_sdr.chm10751430
f1_keywords:
- vis_sdr.chm10751430
ms.prod: visio
api_name:
- Visio.Master.CreateSelection
ms.assetid: 52db8b1b-e253-549f-c3ba-d661fa7b675e
ms.date: 06/08/2017
---


# Master.CreateSelection Method (Visio)

Creates various types of  **Selection** objects.


## Syntax

 _expression_ . **CreateSelection**( **_SelType_** , **_IterationMode_** , **_[Data]_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SelType_|Required| **VisSelectionTypes**|The type of selection. See Remarks for possible values.|
| _IterationMode_|Optional| **VisSelectMode**|The selection mode used. See Remarks for possible values.|
| _Data_|Optional| **Variant**|The object type that corresponds to the  _SelType_ argument. See Remarks for possible values.|

### Return Value

Selection


## Remarks

The  **CreateSelection** method makes it possible to create complex selections programmatically. So instead of having to select all shapes on a page, for example, you can select only those on a given layer, or only those based on a certain master.

Calling the  **CreateSelection** method with _SelType_ equal to **visSelTypeByType** or **visSelTypeByLayer** is equivalent to selecting options in the **Select byType** dialog box (click **Select** in the **Editing** group on the **Home** tab, and then click **Select by Type**).

The  _SelType_ argument should be one of the following values, which are declared in **VisSelectionTypes** in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visSelTypeAll**|1|A selection that initially contains all shapes. |
| **visSelTypeByDataGraphic**|6|A selection that initially contains all shapes that have a given type of data graphic appled.|
| **visSelTypeByLayer**|3|A selection that initially contains all the shapes of a given layer. |
| **visSelTypeByMaster**|5|A selection that initially contains all the instantiated shapes of a given master. |
| **visSelTypeByRole**|7|A selection that initially contains all the shapes of a given role.|
| **visSelTypeByType**|4|A selection that initially contains all the shapes of a given type. |
| **visSelTypeEmpty**|0|A selection that initially contains no shapes. |
| **visSelTypeSingle**|2|A selection that initially contains one shape. |
The optional  _IterationMode_ argument should be one of the following values, which are declared in **VisSelectMode** in the Visio type library. The default is **visSelModeSkipSuper** .



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visSelModeOnlySub**|&;H0800|Selection only reports subselected shapes.|
| **visSelModeOnlySuper**|&;H0200|Selection only reports superselected shapes.|
| **visSelModeSkipSub**|&;H0400|Selection does not report subselected shapes.|
| **visSelModeSkipSuper**|&;H0100|Selection does not report superselected shapes.|
The optional  _Data_ argument should be an object that corresponds to the object type specified by _SelType_. For example, if you want to select all the masters of a certain type,  _Data_ should be of type **Master** . And if you want to select all the shapes on a certain layer, _Data_ should be of type **Layer** .

When  _SelType_ is **visSelTypeByRole** , _Data_ should be a member of the **[VisRoleSelectionTypes](visroleselectiontypes-enumeration-visio.md)** enumeration.

When the  _SelType_ argument is **visSelTypeByType** , possible _Data_ values should be one of the following values, which are declared in **VisTypeSelectionTypes** in the Visio type library.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visTypeSelBitmap**|16|A shape that is a bitmap.|
| **visTypeSelGroup**|1|A shape that contains other shapes.|
| **visTypeSelGuide**|4|A shape that is a guide.|
| **visTypeSelInk**|32|A shape that is ink.|
| **visTypeSelMetafile**|8|A shape that is a metafile.|
| **visTypeSelOLE**|64|A shape that is linked, embedded, or a control.|
| **visTypeSelShape**|2|A native Visio shape.|

