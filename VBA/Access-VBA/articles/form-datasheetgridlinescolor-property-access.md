---
title: Form.DatasheetGridlinesColor Property (Access)
keywords: vbaac10.chm13403
f1_keywords:
- vbaac10.chm13403
ms.prod: access
api_name:
- Access.Form.DatasheetGridlinesColor
ms.assetid: 92d07c1c-fc47-0049-7da3-a34ee56fbc83
ms.date: 06/08/2017
---


# Form.DatasheetGridlinesColor Property (Access)

You can use the  **DatasheetGridlinesColor** property to specify the color of gridlines in a datasheet. Read/write **Long**.


## Syntax

 _expression_. **DatasheetGridlinesColor**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **DatasheetGridlinesColor** property applies only to objects in Datasheet view.

This property is only available in [Visual Basic](set-properties-by-using-visual-basic.md)within a Microsoft Access database.

You can also use the  **RGB** or **QBColor** functions to set this property.

This property setting affects the gridline color for the entire datasheet. It's not possible to set the gridline color of individual cells in Datasheet view.

The following table contains the properties that don't exist in the DAO  **Properties** collection of until you set them by using the **Formatting (Datasheet)** toolbar or you can add them in an Access database by using the **CreateProperty** method and append it to the DAO **Properties** collection.


|||
|:-----|:-----|
|**[DatasheetFontItalic](form-datasheetfontitalic-property-access.md)** *|**[DatasheetForeColor](form-datasheetforecolor-property-access.md)** *|
|**[DatasheetFontHeight](form-datasheetfontheight-property-access.md)** *|**[DatasheetBackColor](form-datasheetbackcolor-property-access.md)**|
|**[DatasheetFontName](form-datasheetfontname-property-access.md)** *|**DatasheetGridlinesColor**|
|**[DatasheetFontUnderline](form-datasheetfontunderline-property-access.md)** *|**[DatasheetGridlinesBehavior](form-datasheetgridlinesbehavior-property-access.md)**|
|**[DatasheetFontWeight](form-datasheetfontweight-property-access.md)** *|**[DatasheetCellsEffect](form-datasheetcellseffect-property-access.md)**|

 **Note**  When you add or set any property listed with an asterisk, Microsoft Access automatically adds all the properties listed with an asterisk to the  **Properties** collection in the database.


## See also


#### Concepts


[Form Object](form-object-access.md)

