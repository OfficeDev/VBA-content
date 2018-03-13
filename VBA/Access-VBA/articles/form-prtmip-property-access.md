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



| <strong>Member</strong>                          | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                   |
|:-------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| LeftMargin, RightMargin, TopMargin, BottomMargin | A  <strong>Long</strong> that specifies the distance between the edge of the page and the item to be printed in twips.                                                                                                                                                                                                                         |
| DataOnly                                         | A  <strong>Long</strong> that specifies the elements to be printed. When <strong>True</strong>, prints only the data in a table or query in Datasheet view, form, or report, and suppresses labels, control borders, grid lines, and display graphics such as lines and boxes. When <strong>False</strong>, prints data, labels, and graphics. |
| ItemsAcross                                      | A  <strong>Long</strong> that specifies the number of columns across for multiple-column reports or labels. This member is equivalent to the value of the <strong>Number of Columns</strong> box under <strong>Grid Settings</strong> on the <strong>Columns</strong> tab of the <strong>Page Setup</strong> dialog box.                       |
| RowSpacing                                       | A  <strong>Long</strong> that specifies the horizontal space between detail sections in units of 1/20 of a point.                                                                                                                                                                                                                              |
| ColumnSpacing                                    | A  <strong>Long</strong> that specifies the vertical space between detail sections in twips.                                                                                                                                                                                                                                                   |
| DefaultSize                                      | A  <strong>Long</strong>. When <strong>True</strong>, uses the size of the detail section in Design view. When <strong>False</strong>, uses the values specified by the ItemSizeWidth and ItemSizeHeight members.                                                                                                                              |
| ItemSizeWidth                                    | A  <strong>Long</strong> that specifies the width of the detail section in twips. This member is equivalent to the value of the <strong>Width</strong> box under <strong>Column Size</strong> on the <strong>Columns</strong> tab of the <strong>Page Setup</strong> dialog box.                                                               |
| ItemSizeHeight                                   | A  <strong>Long</strong> that specifies the height of the detail section twips. This member is equivalent to the value of the <strong>Height</strong> box under <strong>Column Size</strong> on the <strong>Columns</strong> tab of the <strong>Page Setup</strong> dialog box.                                                                |
| ItemLayout                                       | A  <strong>Long</strong> that specifies horizontal (1953) or vertical (1954) layout of columns. This member is equivalent to <strong>Across, then Down</strong> or <strong>Down, then Across</strong> respectively under <strong>Column Layout</strong> on the <strong>Columns</strong> tab of the <strong>Page Setup</strong> dialog box.     |
| FastPrint                                        | Reserved.                                                                                                                                                                                                                                                                                                                                      |
| Datasheet                                        | Reserved.                                                                                                                                                                                                                                                                                                                                      |

The  **PrtMip** property setting is read/write in Design view and read-only in other views.


## See also


#### Concepts


[Form Object](form-object-access.md)

