---
title: BoundObjectFrame.SourceItem Property (Access)
keywords: vbaac10.chm10914
f1_keywords:
- vbaac10.chm10914
ms.prod: access
api_name:
- Access.BoundObjectFrame.SourceItem
ms.assetid: ab802b9b-d17c-695b-aaf5-4f84d1935615
ms.date: 06/08/2017
---


# BoundObjectFrame.SourceItem Property (Access)

You can use the  **SourceItem** property to specify the data within a file to be linked when you create a linked OLE object. Read/write **String**.


## Syntax

 _expression_. **SourceItem**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

You can set the  **SourceItem** property by specifying data in units recognized by the application supplying the object. For example, when you link to Microsoft Excel, you specify the **SourceItem** property setting by using a cell or cell-range reference such as R1C1 or R3C4:R9C22 or a named range such as Revenues.


 **Note**  To determine the syntax to describe a unit of data for a particular object, see the documentation for the application that was used to create the object.

The control's  **OLETypeAllowed** property must be set to Linked or Either when you use this property. Use the control's **SourceDoc** property to specify the file to link.


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


[BoundObjectFrame Object](boundobjectframe-object-access.md)

