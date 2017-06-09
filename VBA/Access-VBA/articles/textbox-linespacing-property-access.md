---
title: TextBox.LineSpacing Property (Access)
keywords: vbaac10.chm11140
f1_keywords:
- vbaac10.chm11140
ms.prod: access
api_name:
- Access.TextBox.LineSpacing
ms.assetid: 3ac1c335-4b26-1a14-e4dc-bd5d56f44a2b
ms.date: 06/08/2017
---


# TextBox.LineSpacing Property (Access)

You can use the  **LineSpacing** property to specify or determine the location of information displayed within a label or text box control. Read/write **Integer**.


## Syntax

 _expression_. **LineSpacing**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

A control's displayed information location is the distance measured between each line of the displayed information. To use a unit of measurement different from the setting in the  **Regional Options** dialog box in **Windows Control Panel**, specify the unit, such as cm or in (for example, 3 cm or 2 in).

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips.


## Example

The following example sets the line spacing to 0.25 inches for the text box "PurchaseOrderInformation" on the "Purchase Order Form"


```vb
' 0.25 inches = 360/1440 twips. 
 Forms("Purchase Orders").Controls("PurchaseOrderDescription").LineSpacing = 360
```


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

