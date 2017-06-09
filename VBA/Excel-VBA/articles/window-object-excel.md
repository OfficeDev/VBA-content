---
title: Window Object (Excel)
keywords: vbaxl10.chm355072
f1_keywords:
- vbaxl10.chm355072
ms.prod: excel
api_name:
- Excel.Window
ms.assetid: 8591b1ad-76f8-14e2-9120-406b65093f5a
ms.date: 06/08/2017
---


# Window Object (Excel)

Represents a window.


## Remarks

 Many worksheet characteristics, such as scroll bars and gridlines, are actually properties of the window. The **Window** object is a member of the **[Windows](windows-object-excel.md)** collection. The **Windows** collection for the **Application** object contains all the windows in the application, whereas the **Windows** collection for the **Workbook** object contains only the windows in the specified workbook.


## Example

Use  **Windows** ( _index_ ), where _index_ is the window name or index number, to return a single **Window** object. The following example maximizes the active window.


```
Windows(1).WindowState = xlMaximized
```

Note that the active window is always  `Windows(1)`.

The window caption is the text shown in the title bar at the top of the window when the window isn't maximized. The caption is also shown in the list of open files on the bottom of the  **Windows** menu. Use the **[Caption](window-caption-property-excel.md)** property to set or return the window caption. Changing the window caption doesn't change the name of the workbook. The following example turns off cell gridlines for the worksheet shown in the Book1.xls:1 window.




```
Windows("book1.xls":1).DisplayGridlines = False
```


## Methods



|**Name**|
|:-----|
|[Activate](window-activate-method-excel.md)|
|[ActivateNext](window-activatenext-method-excel.md)|
|[ActivatePrevious](window-activateprevious-method-excel.md)|
|[Close](window-close-method-excel.md)|
|[LargeScroll](window-largescroll-method-excel.md)|
|[NewWindow](window-newwindow-method-excel.md)|
|[PointsToScreenPixelsX](window-pointstoscreenpixelsx-method-excel.md)|
|[PointsToScreenPixelsY](window-pointstoscreenpixelsy-method-excel.md)|
|[PrintOut](window-printout-method-excel.md)|
|[PrintPreview](window-printpreview-method-excel.md)|
|[RangeFromPoint](window-rangefrompoint-method-excel.md)|
|[ScrollIntoView](window-scrollintoview-method-excel.md)|
|[ScrollWorkbookTabs](window-scrollworkbooktabs-method-excel.md)|
|[SmallScroll](window-smallscroll-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[ActiveCell](window-activecell-property-excel.md)|
|[ActiveChart](window-activechart-property-excel.md)|
|[ActivePane](window-activepane-property-excel.md)|
|[ActiveSheet](window-activesheet-property-excel.md)|
|[ActiveSheetView](window-activesheetview-property-excel.md)|
|[Application](window-application-property-excel.md)|
|[AutoFilterDateGrouping](window-autofilterdategrouping-property-excel.md)|
|[Caption](window-caption-property-excel.md)|
|[Creator](window-creator-property-excel.md)|
|[DisplayFormulas](window-displayformulas-property-excel.md)|
|[DisplayGridlines](window-displaygridlines-property-excel.md)|
|[DisplayHeadings](window-displayheadings-property-excel.md)|
|[DisplayHorizontalScrollBar](window-displayhorizontalscrollbar-property-excel.md)|
|[DisplayOutline](window-displayoutline-property-excel.md)|
|[DisplayRightToLeft](window-displayrighttoleft-property-excel.md)|
|[DisplayRuler](window-displayruler-property-excel.md)|
|[DisplayVerticalScrollBar](window-displayverticalscrollbar-property-excel.md)|
|[DisplayWhitespace](window-displaywhitespace-property-excel.md)|
|[DisplayWorkbookTabs](window-displayworkbooktabs-property-excel.md)|
|[DisplayZeros](window-displayzeros-property-excel.md)|
|[EnableResize](window-enableresize-property-excel.md)|
|[FreezePanes](window-freezepanes-property-excel.md)|
|[GridlineColor](window-gridlinecolor-property-excel.md)|
|[GridlineColorIndex](window-gridlinecolorindex-property-excel.md)|
|[Height](window-height-property-excel.md)|
|[Hwnd](window-hwnd-property-excel.md)|
|[Index](window-index-property-excel.md)|
|[Left](window-left-property-excel.md)|
|[OnWindow](window-onwindow-property-excel.md)|
|[Panes](window-panes-property-excel.md)|
|[Parent](window-parent-property-excel.md)|
|[RangeSelection](window-rangeselection-property-excel.md)|
|[ScrollColumn](window-scrollcolumn-property-excel.md)|
|[ScrollRow](window-scrollrow-property-excel.md)|
|[SelectedSheets](window-selectedsheets-property-excel.md)|
|[Selection](window-selection-property-excel.md)|
|[SheetViews](window-sheetviews-property-excel.md)|
|[Split](window-split-property-excel.md)|
|[SplitColumn](window-splitcolumn-property-excel.md)|
|[SplitHorizontal](window-splithorizontal-property-excel.md)|
|[SplitRow](window-splitrow-property-excel.md)|
|[SplitVertical](window-splitvertical-property-excel.md)|
|[TabRatio](window-tabratio-property-excel.md)|
|[Top](window-top-property-excel.md)|
|[Type](window-type-property-excel.md)|
|[UsableHeight](window-usableheight-property-excel.md)|
|[UsableWidth](window-usablewidth-property-excel.md)|
|[View](window-view-property-excel.md)|
|[Visible](window-visible-property-excel.md)|
|[VisibleRange](window-visiblerange-property-excel.md)|
|[Width](window-width-property-excel.md)|
|[WindowNumber](window-windownumber-property-excel.md)|
|[WindowState](window-windowstate-property-excel.md)|
|[Zoom](window-zoom-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
