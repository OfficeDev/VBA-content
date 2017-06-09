---
title: Report.PrtMip Property (Access)
keywords: vbaac10.chm13737
f1_keywords:
- vbaac10.chm13737
ms.prod: access
api_name:
- Access.Report.PrtMip
ms.assetid: f2a3eb10-04d5-c1fc-5ca3-0dc588db18ff
ms.date: 06/08/2017
---


# Report.PrtMip Property (Access)

You can use the  **PrtMip** property in Visual Basic to set or return the device mode information specified for a form or report in the **Print** dialog box.


## Syntax

 _expression_. **PrtMip**

 _expression_ A variable that represents a **Report** object.


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


## Example

The following  **PrtMip** property example demonstrates how to set up the report with two horizontal columns.


```vb
Private Type str_PRTMIP 
 strRGB As String * 28 
End Type 
 
Private Type type_PRTMIP 
 xLeftMargin As Long 
 yTopMargin As Long 
 xRightMargin As Long 
 yBotMargin As Long 
 fDataOnly As Long 
 xWidth As Long 
 yHeight As Long 
 fDefaultSize As Long 
 cxColumns As Long 
 yColumnSpacing As Long 
 xRowSpacing As Long 
 rItemLayout As Long 
 fFastPrint As Long 
 fDatasheet As Long 
End Type 
 
Public Sub PrtMipCols(ByVal strName As String) 
 
 Dim PrtMipString As str_PRTMIP 
 Dim PM As type_PRTMIP 
 Dim rpt As Report 
 Const PM_HORIZONTALCOLS = 1953 
 Const PM_VERTICALCOLS = 1954 
 
 ' Open the report. 
 DoCmd.OpenReport strName, acDesign 
 Set rpt = Reports(strName) 
 PrtMipString.strRGB = rpt.PrtMip 
 LSet PM = PrtMipString 
 
 ' Create two columns. 
 PM.cxColumns = 2 
 
 ' Set 0.25 inch between rows. 
 PM.xRowSpacing = 0.25 * 1440 
 
 ' Set 0.5 inch between columns. 
 PM.yColumnSpacing = 0.5 * 1440 
 PM.rItemLayout = PM_HORIZONTALCOLS 
 
 ' Update property. 
 LSet PrtMipString = PM 
 rpt.PrtMip = PrtMipString.strRGB 
 
 Set rpt = Nothing 
 
End Sub
```

The next  **PrtMip** property example shows how to set all margins to 1 inch.




```vb
Public Sub SetMarginsToDefault(ByVal strName As String) 
 
 Dim PrtMipString As str_PRTMIP 
 Dim PM As type_PRTMIP 
 Dim rpt As Report 
 
 ' Open the report. 
 DoCmd.OpenReport strName, acDesign 
 Set rpt = Reports(strName) 
 PrtMipString.strRGB = rpt.PrtMip 
 LSet PM = PrtMipString 
 
 ' Set margins. 
 PM.xLeftMargin = 1 * 1440 
 PM.yTopMargin = 1 * 1440 
 PM.xRightMargin = 1 * 1440 
 PM.yBotMargin = 1 * 1440 
 
 ' Update property. 
 LSet PrtMipString = PM 
 rpt.PrtMip = PrtMipString.strRGB 
 
 Set rpt = Nothing 
 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

