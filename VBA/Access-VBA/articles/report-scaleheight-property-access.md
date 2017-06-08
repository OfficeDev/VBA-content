---
title: Report.ScaleHeight Property (Access)
keywords: vbaac10.chm13743
f1_keywords:
- vbaac10.chm13743
ms.prod: access
api_name:
- Access.Report.ScaleHeight
ms.assetid: b150ece7-b285-669f-8677-f28d6899454b
ms.date: 06/08/2017
---


# Report.ScaleHeight Property (Access)

You can use the  **ScaleHeight** property to specify the number of units for the vertical measurement of the page when the **[Circle](report-circle-method-access.md)**, **[Line](report-line-method-access.md)**, **[Pset](report-pset-method-access.md)**, or **[Print](report-print-method-access.md)** method is used while a report is printed or previewed, or its output is saved to a file. Read/write **Single**.


## Syntax

 _expression_. **ScaleHeight**

 _expression_ A variable that represents a **Report** object.


## Remarks

The default setting is the internal height of a report page in twips.

You can set the  **ScaleHeight** property by using a macro or a[Visual Basic](set-properties-by-using-visual-basic.md)event procedure specified by a section's **OnPrint** property setting.

You can use the  **ScaleHeight** property to create a custom coordinate scale for drawing or printing. For example, the statement `ScaleHeight = 100` defines the internal height of the section as 100 units, or one vertical unit as one one-hundredth of the height.

Use the  **[ScaleMode](report-scalemode-property-access.md)** property to define a scale based on a standard unit of measurement, such as points, pixels, characters, inches, millimeters, or centimeters.

Setting the  **ScaleHeight** property to a positive value makes coordinates increase in value from top to bottom. Setting it to a negative value makes coordinates increase in value from bottom to top.

By using these properties and the related  **ScaleLeft** and **ScaleTop** properties, you can set up a custom coordinate system with both positive and negative coordinates. All four of these Scale properties interact with the **ScaleMode** property in the following ways:


- Setting any other Scale property to any value automatically sets the  **ScaleMode** property to 0.
    
- Setting the  **ScaleMode** property to a number greater than 0 changes the **ScaleHeight** and **ScaleWidth** properties to the new unit of measurement and sets the **ScaleLeft** and **ScaleTop** properties to 0. Also, the **CurrentX** and **CurrentY** property settings change to reflect the new coordinates of the current point.
    
You can also use the  **Scale** method to set the **ScaleHeight**, **ScaleWidth**, **ScaleLeft**, and **ScaleTop** properties in one statement.


 **Note**  The  **ScaleHeight** property isn't the same as the **Height** property.


## Example

The following example uses the  **Print** method to display text on a report named Report1. It uses the **TextWidth** and **TextHeight** methods to center the text vertically and horizontally.


```vb
Private Sub Detail_Format(Cancel As Integer, _ 
 FormatCount As Integer) 
 Dim rpt as Report 
 Dim strMessage As String 
 Dim intHorSize As Integer, intVerSize As Integer 
 
 Set rpt = Me 
 strMessage = "DisplayMessage" 
 With rpt 
 'Set scale to pixels, and set FontName and 
 'FontSize properties. 
 .ScaleMode = 3 
 .FontName = "Courier" 
 .FontSize = 24 
 End With 
 ' Horizontal width. 
 intHorSize = Rpt.TextWidth(strMessage) 
 ' Vertical height. 
 intVerSize = Rpt.TextHeight(strMessage) 
 ' Calculate location of text to be displayed. 
 Rpt.CurrentX = (Rpt.ScaleWidth/2) - (intHorSize/2) 
 Rpt.CurrentY = (Rpt.ScaleHeight/2) - (intVerSize/2) 
 ' Print text on Report object. 
 Rpt.Print strMessage 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

