---
title: Shapes.AddOLEObject Method (PowerPoint)
keywords: vbapp10.chm543022
f1_keywords:
- vbapp10.chm543022
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddOLEObject
ms.assetid: 88a5aa63-0531-b9d8-43d2-5a995b91b189
ms.date: 06/08/2017
---


# Shapes.AddOLEObject Method (PowerPoint)

Creates an OLE object. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new OLE object.


## Syntax

 _expression_. **AddOLEObject**( **_Left_**, **_Top_**, **_Width_**, **_Height_**, **_ClassName_**, **_FileName_**, **_DisplayAsIcon_**, **_IconFileName_**, **_IconIndex_**, **_IconLabel_**, **_Link_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Left_|Optional|**Single**|The position (in points) of the upper-left corner of the new object relative to the upper-left corner of the slide. The default value is 0 (zero).|
| _Top_|Optional|**Single**|The position (in points) of the upper-left corner of the new object relative to the upper-left corner of the slide. The default value is 0 (zero).|
| _Width_|Optional|**Single**|The initial width of the OLE object, in points.|
| _Height_|Optional|**Single**|The initial height of the OLE object, in points.|
| _ClassName_|Optional|**String**|The OLE long class name or the ProgID for the object that's to be created. You must specify either the ClassName or FileName argument for the object, but not both.|
| _FileName_|Optional|**String**|The file from which the object is to be created. If the path isn't specified, the current working folder is used. You must specify either the ClassName or FileName argument for the object, but not both.|
| _DisplayAsIcon_|Optional|**[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|Determines whether the OLE object will be displayed as an icon.|
| _IconFileName_|Optional|**String**|The file that contains the icon to be displayed.|
| _IconIndex_|Optional|**Long**|The index of the icon within IconFileName. The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in IconFileName, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).|
| _IconLabel_|Optional|**String**|A label (caption) to be displayed beneath the icon.|
| _Link_|Optional|**[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|Determines whether the OLE object will be linked to the file from which it was created. If you specified a value for ClassName, this argument must be  **msoFalse**.|

### Return Value

Shape


## Example

This example adds a linked Word document to  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.AddOLEObject Left:=100, Top:=100, _ 
    Width:=200, Height:=300, _ 
    FileName:="c:\my documents\testing.doc", Link:=msoTrue
```

This example adds a new Microsoft Excel worksheet to  `myDocument`. The worksheet will be displayed as an icon.




```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.AddOLEObject Left:=100, Top:=100, _ 
    Width:=200, Height:=300, _ 
    ClassName:="Excel.Sheet", DisplayAsIcon:=True
```

This example adds a command button to  `myDocument`.




```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.AddOLEObject Left:=100, Top:=100, _ 
    Width:=150, Height:=50, ClassName:="Forms.CommandButton.1"
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

