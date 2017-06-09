---
title: Window Object (Word)
keywords: vbawd10.chm2402
f1_keywords:
- vbawd10.chm2402
ms.prod: word
api_name:
- Word.Window
ms.assetid: d92f83f9-ae44-56c0-4584-7a9359253c6d
ms.date: 06/08/2017
---


# Window Object (Word)

Represents a window. Many document characteristics, such as scroll bars and rulers, are actually properties of the window.


## Remarks

The  **Window** object is a member of the **[Windows](windows-object-word.md)** collection. The **Windows** collection for the **Application** object contains all the windows in the application, whereas the **Windows** collection for the **Document** object contains only the windows that display the specified document.

Use  **Windows** (Index), where Index is the window name or the index number, to return a single **Window** object. The following example maximizes the Document1 window.




```
Windows("Document1").WindowState = wdWindowStateMaximize
```

The index number is the number to the left of the window name on the  **Window** menu. The following example displays the caption of the first window in the **Windows** collection.




```
MsgBox Windows(1).Caption
```

Use the  **Add** method or the **NewWindow** method to add a new window to the **Windows** collection. Each of the following statements creates a new window for the document in the active window.




```
ActiveDocument.ActiveWindow.NewWindow 
NewWindow 
Windows.Add
```

A colon (:) and a number appear in the window caption when more than one window is open for a document.

When you switch the view to print preview, a new window is created. This window is removed from the  **Windows** collection when you close print preview.


## Methods



|**Name**|
|:-----|
|[Activate](window-activate-method-word.md)|
|[Close](window-close-method-word.md)|
|[GetPoint](window-getpoint-method-word.md)|
|[LargeScroll](window-largescroll-method-word.md)|
|[NewWindow](window-newwindow-method-word.md)|
|[PageScroll](window-pagescroll-method-word.md)|
|[PrintOut](window-printout-method-word.md)|
|[RangeFromPoint](window-rangefrompoint-method-word.md)|
|[ScrollIntoView](window-scrollintoview-method-word.md)|
|[SetFocus](window-setfocus-method-word.md)|
|[SmallScroll](window-smallscroll-method-word.md)|
|[ToggleRibbon](window-toggleribbon-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Active](window-active-property-word.md)|
|[ActivePane](window-activepane-property-word.md)|
|[Application](window-application-property-word.md)|
|[Caption](window-caption-property-word.md)|
|[Creator](window-creator-property-word.md)|
|[DisplayHorizontalScrollBar](window-displayhorizontalscrollbar-property-word.md)|
|[DisplayLeftScrollBar](window-displayleftscrollbar-property-word.md)|
|[DisplayRightRuler](window-displayrightruler-property-word.md)|
|[DisplayRulers](window-displayrulers-property-word.md)|
|[DisplayScreenTips](window-displayscreentips-property-word.md)|
|[DisplayVerticalRuler](window-displayverticalruler-property-word.md)|
|[DisplayVerticalScrollBar](window-displayverticalscrollbar-property-word.md)|
|[Document](window-document-property-word.md)|
|[DocumentMap](window-documentmap-property-word.md)|
|[EnvelopeVisible](window-envelopevisible-property-word.md)|
|[Height](window-height-property-word.md)|
|[HorizontalPercentScrolled](window-horizontalpercentscrolled-property-word.md)|
|[Hwnd](window-hwnd-property-word.md)|
|[IMEMode](window-imemode-property-word.md)|
|[Index](window-index-property-word.md)|
|[Left](window-left-property-word.md)|
|[Next](window-next-property-word.md)|
|[Panes](window-panes-property-word.md)|
|[Parent](window-parent-property-word.md)|
|[Previous](window-previous-property-word.md)|
|[Selection](window-selection-property-word.md)|
|[ShowSourceDocuments](window-showsourcedocuments-property-word.md)|
|[Split](window-split-property-word.md)|
|[SplitVertical](window-splitvertical-property-word.md)|
|[StyleAreaWidth](window-styleareawidth-property-word.md)|
|[Thumbnails](window-thumbnails-property-word.md)|
|[Top](window-top-property-word.md)|
|[Type](window-type-property-word.md)|
|[UsableHeight](window-usableheight-property-word.md)|
|[UsableWidth](window-usablewidth-property-word.md)|
|[VerticalPercentScrolled](window-verticalpercentscrolled-property-word.md)|
|[View](window-view-property-word.md)|
|[Visible](window-visible-property-word.md)|
|[Width](window-width-property-word.md)|
|[WindowNumber](window-windownumber-property-word.md)|
|[WindowState](window-windowstate-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
