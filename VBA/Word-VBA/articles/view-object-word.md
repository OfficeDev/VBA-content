---
title: View Object (Word)
keywords: vbawd10.chm2469
f1_keywords:
- vbawd10.chm2469
ms.prod: word
api_name:
- Word.View
ms.assetid: 8bf5b26b-14c0-1985-65b2-3e034360baeb
ms.date: 08/15/2017
---


# View Object (Word)

Contains the view attributes (such as show all, field shading, and table gridlines) for a window or pane.


## Remarks

Use the  **View** property to return the **View** object. The following example sets view options for the active window.


```
With ActiveDocument.ActiveWindow.View 
 .ShowAll = True 
 .TableGridlines = True 
 .WrapToWindow = False 
End With
```

Use the  **Type** property to change the view. The following example switches the active window to normal view.




```
ActiveDocument.ActiveWindow.View.Type = wdNormalView
```

Use the  **Percentage** property to change the size of the text on-screen. The following example enlarges the on-screen text to 120 percent.




```
ActiveDocument.ActiveWindow.View.Zoom.Percentage = 120
```

Use the  **SeekView** property to view comments, endnotes, footnotes, or the document header or footer. The following example displays the current footer in the active window in print layout view.




```
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekCurrentPageFooter 
End With
```


## Methods



|**Name**|
|:-----|
|[CollapseAllHeadings](view-collapseallheadings-method-word.md)|
|[CollapseOutline](view-collapseoutline-method-word.md)|
|[ExpandAllHeadings](view-expandallheadings-method-word.md)|
|[ExpandOutline](view-expandoutline-method-word.md)|
|[ForceLowresUpdate](http://msdn.microsoft.com/library/85f017eb-8506-53ad-d9f8-beb759572cde%28Office.15%29.aspx)|
|[ForceOffscreenUpdate](http://msdn.microsoft.com/library/d1394841-4cd2-0e3f-b4be-116baf1110b3%28Office.15%29.aspx)|
|[NextHeaderFooter](view-nextheaderfooter-method-word.md)|
|[PreviousHeaderFooter](view-previousheaderfooter-method-word.md)|
|[ShowAllHeadings](view-showallheadings-method-word.md)|
|[ShowHeading](view-showheading-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](view-application-property-word.md)|
|[ColumnWidth](view-columnwidth-property-word.md)|
|[ConflictMode](view-conflictmode-property-word.md)|
|[Creator](view-creator-property-word.md)|
|[DisplayBackgrounds](view-displaybackgrounds-property-word.md)|
|[DisplayPageBoundaries](view-displaypageboundaries-property-word.md)|
|[Draft](view-draft-property-word.md)|
|[FieldShading](view-fieldshading-property-word.md)|
|[FullScreen](view-fullscreen-property-word.md)|
|[Magnifier](view-magnifier-property-word.md)|
|[MailMergeDataView](view-mailmergedataview-property-word.md)|
|[MarkupMode](view-markupmode-property-word.md)|
|[PageColor](view-pagecolor-property-word.md)|
|[PageMovementType](view-pagemovementtype-property-word.md)|
|[Panning](view-panning-property-word.md)|
|[Parent](view-parent-property-word.md)|
|[ReadingLayout](view-readinglayout-property-word.md)|
|[ReadingLayoutActualView](view-readinglayoutactualview-property-word.md)|
|[ReadingLayoutTruncateMargins](view-readinglayouttruncatemargins-property-word.md)|
|[RevisionsBalloonShowConnectingLines](view-revisionsballoonshowconnectinglines-property-word.md)|
|[RevisionsBalloonSide](view-revisionsballoonside-property-word.md)|
|[RevisionsBalloonWidth](view-revisionsballoonwidth-property-word.md)|
|[RevisionsBalloonWidthType](view-revisionsballoonwidthtype-property-word.md)|
|[RevisionsFilter](view-revisionsfilter-property-word.md)|
|[SeekView](view-seekview-property-word.md)|
|[ShadeEditableRanges](view-shadeeditableranges-property-word.md)|
|[ShowAll](view-showall-property-word.md)|
|[ShowBookmarks](view-showbookmarks-property-word.md)|
|[ShowComments](view-showcomments-property-word.md)|
|[ShowCropMarks](view-showcropmarks-property-word.md)|
|[ShowDrawings](view-showdrawings-property-word.md)|
|[ShowFieldCodes](view-showfieldcodes-property-word.md)|
|[ShowFirstLineOnly](view-showfirstlineonly-property-word.md)|
|[ShowFormat](view-showformat-property-word.md)|
|[ShowFormatChanges](view-showformatchanges-property-word.md)|
|[ShowHiddenText](view-showhiddentext-property-word.md)|
|[ShowHighlight](view-showhighlight-property-word.md)|
|[ShowHyphens](view-showhyphens-property-word.md)|
|[ShowInkAnnotations](view-showinkannotations-property-word.md)|
|[ShowInsertionsAndDeletions](view-showinsertionsanddeletions-property-word.md)|
|[ShowMainTextLayer](view-showmaintextlayer-property-word.md)|
|[ShowMarkupAreaHighlight](view-showmarkupareahighlight-property-word.md)|
|[ShowObjectAnchors](view-showobjectanchors-property-word.md)|
|[ShowOptionalBreaks](view-showoptionalbreaks-property-word.md)|
|[ShowOtherAuthors](view-showotherauthors-property-word.md)|
|[ShowParagraphs](view-showparagraphs-property-word.md)|
|[ShowPicturePlaceHolders](view-showpictureplaceholders-property-word.md)|
|[ShowRevisionsAndComments](view-showrevisionsandcomments-property-word.md)|
|[ShowSpaces](view-showspaces-property-word.md)|
|[ShowTabs](view-showtabs-property-word.md)|
|[ShowTextBoundaries](view-showtextboundaries-property-word.md)|
|[ShowXMLMarkup](view-showxmlmarkup-property-word.md)|
|[SplitSpecial](view-splitspecial-property-word.md)|
|[TableGridlines](view-tablegridlines-property-word.md)|
|[Type](view-type-property-word.md)|
|[WrapToWindow](view-wraptowindow-property-word.md)|
|[Zoom](view-zoom-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
