---
title: PageSetup Object (Excel)
keywords: vbaxl10.chm472072
f1_keywords:
- vbaxl10.chm472072
ms.prod: excel
api_name:
- Excel.PageSetup
ms.assetid: 2fd22df9-5987-f723-04a9-9a3f2e84ac81
ms.date: 06/08/2017
---


# PageSetup Object (Excel)

Represents the page setup description.


## Remarks

 The **PageSetup** object contains all page setup attributes (left margin, bottom margin, paper size, and so on) as properties.


## Example

Use the  **[PageSetup](worksheet-pagesetup-property-excel.md)** property to return a **PageSetup** object. The following example sets the orientation to landscape mode and then prints the worksheet.


```
With Worksheets("Sheet1") 
 .PageSetup.Orientation = xlLandscape 
 .PrintOut 
End With
```

The  **With** statement makes it easier and faster to set several properties at the same time. The following example sets all the margins for worksheet one.




```
With Worksheets(1).PageSetup 
 .LeftMargin = Application.InchesToPoints(0.5) 
 .RightMargin = Application.InchesToPoints(0.75) 
 .TopMargin = Application.InchesToPoints(1.5) 
 .BottomMargin = Application.InchesToPoints(1) 
 .HeaderMargin = Application.InchesToPoints(0.5) 
 .FooterMargin = Application.InchesToPoints(0.5) 
End With
```


## Properties



|**Name**|
|:-----|
|[AlignMarginsHeaderFooter](pagesetup-alignmarginsheaderfooter-property-excel.md)|
|[Application](pagesetup-application-property-excel.md)|
|[BlackAndWhite](pagesetup-blackandwhite-property-excel.md)|
|[BottomMargin](pagesetup-bottommargin-property-excel.md)|
|[CenterFooter](pagesetup-centerfooter-property-excel.md)|
|[CenterFooterPicture](pagesetup-centerfooterpicture-property-excel.md)|
|[CenterHeader](pagesetup-centerheader-property-excel.md)|
|[CenterHeaderPicture](pagesetup-centerheaderpicture-property-excel.md)|
|[CenterHorizontally](pagesetup-centerhorizontally-property-excel.md)|
|[CenterVertically](pagesetup-centervertically-property-excel.md)|
|[Creator](pagesetup-creator-property-excel.md)|
|[DifferentFirstPageHeaderFooter](pagesetup-differentfirstpageheaderfooter-property-excel.md)|
|[Draft](pagesetup-draft-property-excel.md)|
|[EvenPage](pagesetup-evenpage-property-excel.md)|
|[FirstPage](pagesetup-firstpage-property-excel.md)|
|[FirstPageNumber](pagesetup-firstpagenumber-property-excel.md)|
|[FitToPagesTall](pagesetup-fittopagestall-property-excel.md)|
|[FitToPagesWide](pagesetup-fittopageswide-property-excel.md)|
|[FooterMargin](pagesetup-footermargin-property-excel.md)|
|[HeaderMargin](pagesetup-headermargin-property-excel.md)|
|[LeftFooter](pagesetup-leftfooter-property-excel.md)|
|[LeftFooterPicture](pagesetup-leftfooterpicture-property-excel.md)|
|[LeftHeader](pagesetup-leftheader-property-excel.md)|
|[LeftHeaderPicture](pagesetup-leftheaderpicture-property-excel.md)|
|[LeftMargin](pagesetup-leftmargin-property-excel.md)|
|[OddAndEvenPagesHeaderFooter](pagesetup-oddandevenpagesheaderfooter-property-excel.md)|
|[Order](pagesetup-order-property-excel.md)|
|[Orientation](pagesetup-orientation-property-excel.md)|
|[Pages](pagesetup-pages-property-excel.md)|
|[PaperSize](pagesetup-papersize-property-excel.md)|
|[Parent](pagesetup-parent-property-excel.md)|
|[PrintArea](pagesetup-printarea-property-excel.md)|
|[PrintComments](pagesetup-printcomments-property-excel.md)|
|[PrintErrors](pagesetup-printerrors-property-excel.md)|
|[PrintGridlines](pagesetup-printgridlines-property-excel.md)|
|[PrintHeadings](pagesetup-printheadings-property-excel.md)|
|[PrintNotes](pagesetup-printnotes-property-excel.md)|
|[PrintQuality](pagesetup-printquality-property-excel.md)|
|[PrintTitleColumns](pagesetup-printtitlecolumns-property-excel.md)|
|[PrintTitleRows](pagesetup-printtitlerows-property-excel.md)|
|[RightFooter](pagesetup-rightfooter-property-excel.md)|
|[RightFooterPicture](pagesetup-rightfooterpicture-property-excel.md)|
|[RightHeader](pagesetup-rightheader-property-excel.md)|
|[RightHeaderPicture](pagesetup-rightheaderpicture-property-excel.md)|
|[RightMargin](pagesetup-rightmargin-property-excel.md)|
|[ScaleWithDocHeaderFooter](pagesetup-scalewithdocheaderfooter-property-excel.md)|
|[TopMargin](pagesetup-topmargin-property-excel.md)|
|[Zoom](pagesetup-zoom-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
