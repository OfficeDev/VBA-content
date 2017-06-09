---
title: Using ActiveX Controls on Sheets
keywords: vbaxl10.chm5205775
f1_keywords:
- vbaxl10.chm5205775
ms.prod: excel
ms.assetid: eef29794-5bc3-aecb-5ed2-e078c28851b4
ms.date: 06/08/2017
---


# Using ActiveX Controls on Sheets

This topic covers specific information about using ActiveX controls on worksheets and chart sheets. For general information on adding and working with controls, see  [Using ActiveX Controls on a Document](using-activex-controls-on-a-document.md) and [Creating a Custom Dialog Box](create-a-custom-dialog-box.md).

Keep the following points in mind when you are working with controls on sheets:

- In addition to the standard properties available for ActiveX controls, the following properties can be used with ActiveX controls in Microsoft Excel:  **[BottomRightCell](oleobject-bottomrightcell-property-excel.md)**,  **[LinkedCell](oleobject-linkedcell-property-excel.md)**,  **[ListFillRange](oleobject-listfillrange-property-excel.md)**,  **[Placement](oleobject-placement-property-excel.md)**,  **[PrintObject](oleobject-printobject-property-excel.md)**,  **[TopLeftCell](oleobject-topleftcell-property-excel.md)**, and  **[ZOrder](oleobject-zorder-property-excel.md)**.
    
    
These properties can be set and returned using the ActiveX control name. The following example scrolls the workbook window so CommandButton1 is in the upper-left corner.
    


```vb
  Set t = Sheet1.CommandButton1.TopLeftCell
With ActiveWindow
    .ScrollRow = t.Row
    .ScrollColumn = t.Column
End With

```


- Some Microsoft Excel Visual Basic methods and properties are disabled when an ActiveX control is activated. For example, the  **Sort** method cannot be used when a control is active, so the following code fails in a button click event procedure (because the control is still active after the user clicks it).
    
```vb
  Private Sub CommandButton1.Click 
    Range("a1:a10").Sort Key1:=Range("a1") 
End Sub 
```


    You can work around this problem by activating some other element on the sheet before you use the property or method that failed. For example, the following code sorts the range:
    


```vb
  Private Sub CommandButton1.Click 
    Range("a1").Activate 
    Range("a1:a10").Sort Key1:=Range("a1") 
    CommandButton1.Activate 
End Sub
```


- Controls on a Microsoft Excel workbook embedded in a document in another application will not work if the user double-clicks the workbook to edit it. The controls will work if the user right-clicks the workbook and selects the  **Open** command from the shortcut menu.
    
- When a Microsoft Excel workbook is saved using the Microsoft Excel 5.0/95 Workbook file format, ActiveX control information is lost.
    
- The  **Me** keyword in an event procedure for an ActiveX control on a sheet refers to the sheet, not to the control.
    

## Adding Controls with Visual Basic

In Microsoft Excel, ActiveX controls are represented by  **OLEObject** objects in the **OLEObjects** collection (all **OLEObject** objects are also in the **Shapes** collection). To programmatically add an ActiveX control to a sheet, use the **Add** method of the **OLEObjects** collection. The following example adds a command button to worksheet 1.


```vb
Worksheets(1).OLEObjects.Add "Forms.CommandButton.1", _ 
    Left:=10, Top:=10, Height:=20, Width:=100
```


## Using Control Properties with Visual Basic

Most often, your Visual Basic code will refer to ActiveX controls by name. The following example changes the caption on the control named "CommandButton1."


```vb
Sheet1.CommandButton1.Caption = "Run"
```

Note that when you use a control name outside the class module for the sheet containing the control, you must qualify the control name with the sheet name.

To change the control name you use in Visual Basic code, select the control and set the  **(Name)** property in the Properties window.

Because ActiveX controls are also represented by  **OLEObject** objects in the **OLEObjects** collection, you can set control properties using the objects in the collection. The following example sets the left position of the control named "CommandButton1."




```vb
Worksheets(1).OLEObjects("CommandButton1").Left = 10
```

Control properties that are not shown as properties of the  **OLEObject** object can be set by returning the actual control object using the **Object** property. The following example sets the caption for CommandButton1.




```vb
Worksheets(1).OLEObjects("CommandButton1"). _ 
    Object.Caption = "run me"
```

Because all OLE objects are also members of the  **Shapes** collection, you can use the collection to set properties for several controls. The following example aligns the left edge of all controls on worksheet 1.




```vb
For Each s In Worksheets(1).Shapes 
    If s.Type = msoOLEControlObject Then s.Left = 10 
Next
```


## Using Control Names with the Shapes and OLEObjects Collections

An ActiveX control on a sheet has two names: the name of the shape that contains the control, which you can see in the  **Name** box when you view the sheet, and the code name for the control, which you can see in the cell to the right of **(Name)** in the Properties window. When you first add a control to a sheet, the shape name and code name match. However, if you change either the shape name or code name, the other is not automatically changed to match.

You use the code name of a control in the names of its event procedures. However, when you return a control from the  **Shapes** or **OLEObjects** collection for a sheet, you must use the shape name, not the code name, to refer to the control by name. For example, assume that you add a check box to a sheet and that both the default shape name and the default code name are CheckBox1. If you then change the control code name by typing **chkFinished** next to **(Name)** in the Properties window, you must use chkFinished in event procedure names, but you still have to use CheckBox1 to return the control from the **Shapes** or **OLEObject** collection, as shown in the following example.




```vb
Private Sub chkFinished_Click() 
    ActiveSheet.OLEObjects("CheckBox1").Object.Value = 1 
End Sub
```


