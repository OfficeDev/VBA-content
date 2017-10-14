---
title: OLEObjects.Add Method (Excel)
keywords: vbaxl10.chm422087
f1_keywords:
- vbaxl10.chm422087
ms.prod: excel
api_name:
- Excel.OLEObjects.Add
ms.assetid: 2acd369f-6dd6-0e0e-043c-a691796659a9
ms.date: 06/08/2017
---


# OLEObjects.Add Method (Excel)

Adds a new OLE object to a sheet. 


## Syntax

 _expression_ . **Add**( **_ClassType_** , **_FileName_** , **_Link_** , **_DisplayAsIcon_** , **_IconFileName_** , **_IconIndex_** , **_IconLabel_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents an **OLEObjects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ClassType_|Optional| **Variant**|(you must specify either  _ClassType_ or _FileName_). A string that contains the programmatic identifier for the object to be created. If  _ClassType_ is specified, _FileName_ and _Link_ are ignored.|
| _FileName_|Optional| **Variant**|(you must specify either  _ClassType_ or _FileName_). A string that specifies the file to be used to create the OLE object.|
| _Link_|Optional| **Variant**| **True** to have the new OLE object based on _FileName_ be linked to that file. If the object isn't linked, the object is created as a copy of the file. The default value is **False** .|
| _DisplayAsIcon_|Optional| **Variant**| **True** to display the new OLE object either as an icon or as its regular picture. If this argument is **True** , _IconFileName_ and _IconIndex_ can be used to specify an icon.|
| _IconFileName_|Optional| **Variant**|A string that specifies the file that contains the icon to be displayed. This argument is used only if  _DisplayAsIcon_ is **True** . If this argument isn't specified or the file contains no icons, the default icon for the OLE class is used.|
| _IconIndex_|Optional| **Variant**|The number of the icon in the icon file. This is used only if  _DisplayAsIcon_ is **True** and _IconFileName_ refers to a valid file that contains icons. If an icon with the given index number doesn't exist in the file specified by _IconFileName_, the first icon in the file is used.|
| _IconLabel_|Optional| **Variant**|A string that specifies a label to display beneath the icon. This is used only if  _DisplayAsIcon_ is **True** . If this argument is omitted or is an empty string (""), no caption is displayed.|
| _Left_|Optional| **Variant**|The initial coordinates of the new object, in points, relative to the upper-left corner of cell A1 on a worksheet, or to the upper-left corner of a chart.|
| _Width_|Optional| **Variant**|The initial width of the new object, in points.|
| _Height_|Optional| **Variant**|The initial height of the new object, in points.|
| _Top_|Optional| **Variant**|The initial coordinates of the new object in points, relative to the upper-left corner of cell A1 on a worksheet, or to the upper-left corner of a chart.|

### Return Value

An  **[OLEObject](oleobject-object-excel.md)** object that represents the new OLE object.


## Example

This example creates a new Microsoft Word OLE object on Sheet1.


```vb
ActiveWorkbook.Worksheets("Sheet1").OLEObjects.Add _ 
 ClassType:="Word.Document"
```

This example adds a command button to sheet one.




```vb
Worksheets(1).OLEObjects.Add ClassType:="Forms.CommandButton.1", _ 
 Link:=False, DisplayAsIcon:=False, Left:=40, Top:=40, _ 
 Width:=150, Height:=10
```


## See also


#### Concepts


[OLEObjects Object](oleobjects-object-excel.md)

