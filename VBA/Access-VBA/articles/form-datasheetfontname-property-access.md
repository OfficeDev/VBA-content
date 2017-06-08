---
title: Form.DatasheetFontName Property (Access)
keywords: vbaac10.chm13396
f1_keywords:
- vbaac10.chm13396
ms.prod: access
api_name:
- Access.Form.DatasheetFontName
ms.assetid: e6b963ca-7162-912e-e63d-1437904ec8f1
ms.date: 06/08/2017
---


# Form.DatasheetFontName Property (Access)

You can use the  **DatasheetFontName** property to specify the font used to display and print field names and data in Datasheet view. Read/write **String**.


## Syntax

 _expression_. **DatasheetFontName**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **DatasheetFontName** property applies to all fields in Datasheet view and to form controls when the form is in Datasheet view.

This property is only available in [Visual Basic](set-properties-by-using-visual-basic.md)within a Microsoft Access database.

For the  **DatasheetFontName** property, the font names you can specify depend on the fonts installed on your system and for your printer. If you specify a font that your system can't display or that isn't installed, Microsoft Windows will substitute a similar font.

The following table contains the properties that don't exist in the DAO  **Properties** collection of until you set them by using the **Formatting (Datasheet)** toolbar or you can add them in an Access database by using the **CreateProperty** method and append it to the **DAO Properties** collection.


|||
|:-----|:-----|
|**[DatasheetFontItalic](form-datasheetfontitalic-property-access.md)** *|**[DatasheetForeColor](form-datasheetforecolor-property-access.md)** *|
|**[DatasheetFontHeight](form-datasheetfontheight-property-access.md)** *|**[DatasheetBackColor](form-datasheetbackcolor-property-access.md)**|
|**DatasheetFontName** *|**[DatasheetGridlinesColor](form-datasheetgridlinescolor-property-access.md)**|
|**[DatasheetFontUnderline](form-datasheetfontunderline-property-access.md)** *|**[DatasheetGridlinesBehavior](form-datasheetgridlinesbehavior-property-access.md)**|
|**[DatasheetFontWeight](form-datasheetfontweight-property-access.md)** *|**[DatasheetCellsEffect](form-datasheetcellseffect-property-access.md)**|

 **Note**  When you add or set any property listed with an asterisk, Microsoft Access automatically adds all the properties listed with an asterisk to the  **Properties** collection of the database.


## See also


#### Concepts


[Form Object](form-object-access.md)

