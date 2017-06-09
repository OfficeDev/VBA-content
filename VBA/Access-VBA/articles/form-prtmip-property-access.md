---
title: Form.PrtMip Property (Access)
keywords: vbaac10.chm13417
f1_keywords:
- vbaac10.chm13417
ms.prod: access
api_name:
- Access.Form.PrtMip
ms.assetid: 0b87f955-638c-5cd2-95b1-5aec870350ff
ms.date: 06/08/2017
---


# Form.PrtMip Property (Access)

You can use the  **PrtMip** property in Visual Basic to set or return the device mode information specified for a form or report in the **Print** dialog box.


## Syntax

 _expression_. **PrtMip**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **PrtMip** property setting is a 28-byte structure that maps to settings on the **Margins** tab for a form or report in the **Page Setup** dialog box.

The  **PrtMip** property has the following members.



|**Member**|**Description**|
|:-----|:-----|
|LeftMargin, RightMargin, TopMargin, BottomMargin|A  **Long** that specifies the distance between the edge of the page and the item to be printed in twips.|
|DataOnly|A  **Long** that specifies the elements to be printed. When **True**, prints only the data in a table or query in Datasheet view, form, or report, and suppresses labels, control borders, grid lines, and display graphics such as lines and boxes. When **False**, prints data, labels, and graphics.|
|ItemsAcross|A  **Long** that specifies the number of columns across for multiple-column reports or labels. This member is equivalent to the value of the **Number of Columns** box under **Grid Settings** on the **Columns** tab of the **Page Setup** dialog box.|
|RowSpacing|A  **Long** that specifies the horizontal space between detail sections in units of 1/20 of a point.|
|ColumnSpacing|A  **Long** that specifies the vertical space between detail sections in twips.|
|DefaultSize|A  **Long**. When **True**, uses the size of the detail section in Design view. When **False**, uses the values specified by the ItemSizeWidth and ItemSizeHeight members.|
|ItemSizeWidth|A  **Long** that specifies the width of the detail section in twips. This member is equivalent to the value of the **Width** box under **Column Size** on the **Columns** tab of the **Page Setup** dialog box.|
|ItemSizeHeight|A  **Long** that specifies the height of the detail section twips. This member is equivalent to the value of the **Height** box under **Column Size** on the **Columns** tab of the **Page Setup** dialog box.|
|ItemLayout|A  **Long** that specifies horizontal (1953) or vertical (1954) layout of columns. This member is equivalent to **Across, then Down** or **Down, then Across** respectively under **Column Layout** on the **Columns** tab of the **Page Setup** dialog box.|
|FastPrint|Reserved.|
|Datasheet|Reserved.|
The  **PrtMip** property setting is read/write in Design view and read-only in other views.


## See also


#### Concepts


[Form Object](form-object-access.md)

