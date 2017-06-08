---
title: Shapes.AddOLEObject Method (Excel)
keywords: vbaxl10.chm638091
f1_keywords:
- vbaxl10.chm638091
ms.prod: excel
api_name:
- Excel.Shapes.AddOLEObject
ms.assetid: 6e73970f-3c2d-0e4d-8974-14e478bf489a
ms.date: 06/08/2017
---


# Shapes.AddOLEObject Method (Excel)

Creates an OLE object. Returns a  **[Shape](shape-object-excel.md)** object that represents the new OLE object.


## Syntax

 _expression_ . **AddOLEObject**( **_ClassType_** , **_Filename_** , **_Link_** , **_DisplayAsIcon_** , **_IconFileName_** , **_IconIndex_** , **_IconLabel_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ClassType_|Optional| **Variant**|(you must specify either  _ClassType_ or _FileName_). A string that contains the programmatic identifier for the object to be created. If  _ClassType_ is specified, _FileName_ and _Link_ are ignored.|
| _Filename_|Optional| **Variant**| The file from which the object is to be created. If the path isn?t specified, the current working folder is used. You must specify either the _ClassType_ or _FileName_ argument for the object, but not both.|
| _Link_|Optional| **Variant**| **True** to link the OLE object to the file from which it was created. **False** to make the OLE object an independent copy of the file. If you specified a value for _ClassType_, this argument must be  **False** . The default value is **False** .|
| _DisplayAsIcon_|Optional| **Variant**| **True** to display the OLE object as an icon. The default value is **False** .|
| _IconFileName_|Optional| **Variant**| The file that contains the icon to be displayed.|
| _IconIndex_|Optional| **Variant**|The index of the icon within  _IconFileName_. The order of icons in the specified file corresponds to the order in which the icons appear in the  **Change Icon** dialog box (accessed from the **Object** dialog box when the **Display as icon** check box is selected). The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in _IconFileName_, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).|
| _IconLabel_|Optional| **Variant**|A label (caption) to be displayed beneath the icon.|
| _Left_|Optional| **Variant**|The position (in points) of the upper-left corner of the new object relative to the upper-left corner of the document. The default value is 0 (zero).|
| _Top_|Optional| **Variant**|The position (in points) of the upper-left corner of the new object relative to the upper-left corner of the document. The default value is 0 (zero).|
| _Width_|Optional| **Variant**|The initial dimensions of the OLE object, in points.|
| _Height_|Optional| **Variant**|The initial dimensions of the OLE object, in points.|

### Return Value

Shape


## Example

This example adds a linked Word document to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddOLEObject Left:=100, Top:=100, _ 
 Width:=200, Height:=300, _ 
 FileName:="c:\my documents\testing.doc", link:=True
```

This example adds a new command button to  `myDocument`.




```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddOLEObject Left:=100, Top:=100, _ 
 Width:=100, Height:=200, _ 
 ClassType:="Forms.CommandButton.1"
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

