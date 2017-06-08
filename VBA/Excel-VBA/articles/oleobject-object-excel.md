---
title: OLEObject Object (Excel)
keywords: vbaxl10.chm414072
f1_keywords:
- vbaxl10.chm414072
ms.prod: excel
api_name:
- Excel.OLEObject
ms.assetid: bc3ef12d-1531-6c21-71ab-3df6bb851f3b
ms.date: 06/08/2017
---


# OLEObject Object (Excel)

Represents an ActiveX control or a linked or embedded OLE object on a worksheet.


## Remarks

 The **OLEObject** object is a member of the **[OLEObjects](oleobjects-object-excel.md)** collection. The **OLEObjects** collection contains all the OLE objects on a single worksheet.


## Example

Use  **[OLEObjects](worksheet-oleobjects-method-excel.md)** ( _index_ ), where _index_ is the name or number of the object, to return an **OLEObject** object. The following example deletes OLE object one on Sheet1.


```
Worksheets("sheet1").OLEObjects(1).Delete
```

The following example deletes the OLE object named "ListBox1."




```
Worksheets("sheet1").OLEObjects("ListBox1").Delete
```

The properties and methods of the  **OLEObject** object are duplicated on each ActiveX control on a worksheet. This enables Visual Basic code to gain access to these properties by using the control's name. The following example selects the check box control named "MyCheckBox," aligns it with the active cell, and then activates the control.




```
With MyCheckBox 
 .Value = True 
 .Top = ActiveCell.Top 
 .Activate 
End With
```


## Events



|**Name**|
|:-----|
|[GotFocus](oleobject-gotfocus-event-excel.md)|
|[LostFocus](oleobject-lostfocus-event-excel.md)|

## Methods



|**Name**|
|:-----|
|[Activate](oleobject-activate-method-excel.md)|
|[BringToFront](oleobject-bringtofront-method-excel.md)|
|[Copy](oleobject-copy-method-excel.md)|
|[CopyPicture](oleobject-copypicture-method-excel.md)|
|[Cut](oleobject-cut-method-excel.md)|
|[Delete](oleobject-delete-method-excel.md)|
|[Duplicate](oleobject-duplicate-method-excel.md)|
|[Select](oleobject-select-method-excel.md)|
|[SendToBack](oleobject-sendtoback-method-excel.md)|
|[Update](oleobject-update-method-excel.md)|
|[Verb](oleobject-verb-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](oleobject-application-property-excel.md)|
|[AutoLoad](oleobject-autoload-property-excel.md)|
|[AutoUpdate](oleobject-autoupdate-property-excel.md)|
|[Border](oleobject-border-property-excel.md)|
|[BottomRightCell](oleobject-bottomrightcell-property-excel.md)|
|[Creator](oleobject-creator-property-excel.md)|
|[Enabled](oleobject-enabled-property-excel.md)|
|[Height](oleobject-height-property-excel.md)|
|[Index](oleobject-index-property-excel.md)|
|[Interior](oleobject-interior-property-excel.md)|
|[Left](oleobject-left-property-excel.md)|
|[LinkedCell](oleobject-linkedcell-property-excel.md)|
|[ListFillRange](oleobject-listfillrange-property-excel.md)|
|[Locked](oleobject-locked-property-excel.md)|
|[Name](oleobject-name-property-excel.md)|
|[Object](oleobject-object-property-excel.md)|
|[OLEType](oleobject-oletype-property-excel.md)|
|[Parent](oleobject-parent-property-excel.md)|
|[Placement](oleobject-placement-property-excel.md)|
|[PrintObject](oleobject-printobject-property-excel.md)|
|[progID](oleobject-progid-property-excel.md)|
|[Shadow](oleobject-shadow-property-excel.md)|
|[ShapeRange](oleobject-shaperange-property-excel.md)|
|[SourceName](oleobject-sourcename-property-excel.md)|
|[Top](oleobject-top-property-excel.md)|
|[TopLeftCell](oleobject-topleftcell-property-excel.md)|
|[Visible](oleobject-visible-property-excel.md)|
|[Width](oleobject-width-property-excel.md)|
|[ZOrder](oleobject-zorder-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
