---
title: CustomControl.Class Property (Access)
keywords: vbaac10.chm12012
f1_keywords:
- vbaac10.chm12012
ms.prod: access
api_name:
- Access.CustomControl.Class
ms.assetid: c745856b-c447-af0a-ed9e-9945d3917d10
ms.date: 06/08/2017
---


# CustomControl.Class Property (Access)

You can use the  **Class** property to specify or determine the class name of an embeddedOLE object. Read/write **String**.


## Syntax

 _expression_. **Class**

 _expression_ A variable that represents a **CustomControl** object.


## Remarks

The  **Class** property setting is a string expression supplied by you or Microsoft Access when you create or paste an OLE object.

A class name defines the type of OLE object. For example, Microsoft Excel supports several types of OLE objects, including worksheets and charts. Their class names are "Excel.Sheet" and "Excel.Chart" respectively.


 **Note**  To determine the class name of an OLE object, see the documentation for the application supplying the object.

The  **Class** property setting is updated when you copy an object from the Clipboard. For example, if you paste a Microsoft Excel chart from the Clipboard into an OLE object that previously contained a Microsoft Excel worksheet, the **Class** property setting changes from "Excel.Sheet" to "Excel.Chart". You can paste an object from the Clipboard by using Visual Basic to set the control's **Action** property to **acOLEPaste** or **acOLEPasteSpecialDlg**.


 **Note**  The  **OLEClass** property and the **Class** property are similar but not identical. The **OLEClass** property setting is a general description of the OLE object whereas the **Class** property setting is the name used to refer to the OLE object in Visual Basic. Examples of **OLEClass** property settings are Microsoft Excel Chart, Microsoft Word Document, and Paintbrush Picture.


## Example

The following example creates a linked OLE object using an unbound object frame named  `OLE1` and sizes the control to display the object's entire contents when the user clicks a command button.


```vb
Sub Command1_Click 
 OLE1.Class = "Excel.Sheet" ' Set class name. 
 ' Specify type of object. 
 OLE1.OLETypeAllowed = acOLELinked 
 ' Specify source file. 
 OLE1.SourceDoc = "C:\Excel\Oletext.xls" 
 ' Specify data to create link to. 
 OLE1.SourceItem = "R1C1:R5C5" 
 ' Create linked object. 
 OLE1.Action = acOLECreateLink 
 ' Adjust control size. 
 OLE1.SizeMode = acOLESizeZoom 
End Sub
```


## See also


#### Concepts


[CustomControl Object](customcontrol-object-access.md)

