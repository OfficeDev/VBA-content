---
title: OLEObjects Object (Excel)
keywords: vbaxl10.chm418072
f1_keywords:
- vbaxl10.chm418072
ms.prod: excel
api_name:
- Excel.OLEObjects
ms.assetid: e3fcf4bd-7c96-ecb3-dc04-551f7f7348f9
ms.date: 06/08/2017
---


# OLEObjects Object (Excel)

A collection of all the  **[OLEObject](oleobject-object-excel.md)** objects on the specified worksheet.


## Remarks

 Each **OLEObject** object represents an ActiveX control or a linked or embedded OLE object.

An ActiveX control on a sheet has two names: the name of the shape that contains the control, which you can see in the  **Name** box when you view the sheet, and the code name for the control, which you can see in the cell to the right of **(Name)** in the **Properties** window. When you first add a control to a sheet, the shape name and code name match. However, if you change either the shape name or code name, the other is not automatically changed to match.


## Example

Use the  **[OLEObjects](worksheet-oleobjects-method-excel.md)** method to return the **OLEObjects** collection. The following example hides all the OLE objects on worksheet one.


```
Worksheets(1).OLEObjects.Visible = False
```

Use the  **[Add](oleobjects-add-method-excel.md)** method to create a new OLE object and add it to the **OLEObjects** collection. The following example creates a new OLE object representing the bitmap file Arcade.bmp and adds it to worksheet one.




```
Worksheets(1).OLEObjects.Add FileName:="arcade.gif"
```

The following example creates a new ActiveX control (a list box) and adds it to worksheet one.




```
Worksheets(1).OLEObjects.Add ClassType:="Forms.ListBox.1"
```

You use the code name of a control in the names of its event procedures. However, when you return a control from the  **[Shapes](shapes-object-excel.md)** or **OLEObjects** collection for a sheet, you must use the shape name, not the code name, to refer to the control by name. For example, assume that you add a check box to a sheet and that both the default shape name and the default code name are CheckBox1. If you then change the control code name by typing chkFinished next to **(Name)** in the **Properties** window, you must use chkFinished in event procedures names, but you still have to use CheckBox1 to return the control from the **Shapes** or **OLEObject** collection, as shown in the following example.




```
Private Sub chkFinished_Click() 
 ActiveSheet.OLEObjects("CheckBox1").Object.Value = 1 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](oleobjects-add-method-excel.md)|
|[BringToFront](oleobjects-bringtofront-method-excel.md)|
|[Copy](oleobjects-copy-method-excel.md)|
|[CopyPicture](oleobjects-copypicture-method-excel.md)|
|[Cut](oleobjects-cut-method-excel.md)|
|[Delete](oleobjects-delete-method-excel.md)|
|[Duplicate](oleobjects-duplicate-method-excel.md)|
|[Item](oleobjects-item-method-excel.md)|
|[Select](oleobjects-select-method-excel.md)|
|[SendToBack](oleobjects-sendtoback-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](oleobjects-application-property-excel.md)|
|[AutoLoad](oleobjects-autoload-property-excel.md)|
|[Border](oleobjects-border-property-excel.md)|
|[Count](oleobjects-count-property-excel.md)|
|[Creator](oleobjects-creator-property-excel.md)|
|[Enabled](oleobjects-enabled-property-excel.md)|
|[Height](oleobjects-height-property-excel.md)|
|[Interior](oleobjects-interior-property-excel.md)|
|[Left](oleobjects-left-property-excel.md)|
|[Locked](oleobjects-locked-property-excel.md)|
|[Parent](oleobjects-parent-property-excel.md)|
|[Placement](oleobjects-placement-property-excel.md)|
|[PrintObject](oleobjects-printobject-property-excel.md)|
|[Shadow](oleobjects-shadow-property-excel.md)|
|[ShapeRange](oleobjects-shaperange-property-excel.md)|
|[SourceName](oleobjects-sourcename-property-excel.md)|
|[Top](oleobjects-top-property-excel.md)|
|[Visible](oleobjects-visible-property-excel.md)|
|[Width](oleobjects-width-property-excel.md)|
|[ZOrder](oleobjects-zorder-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
