
# SeriesCollection.Add Method (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Adds one or more new series to the  **SeriesCollection** collection.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Add**( **_Source_**,  **_Rowcol_**,  **_SeriesLabels_**,  **_CategoryLabels_**,  **_Replace_**)

 _expression_A variable that represents a  **SeriesCollection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Source|Required| **Variant**|The new data as a  ** [Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)** object.|
|Rowcol|Optional| ** [XlRowCol](78f808d5-e5e4-bee8-93ae-d2589d854fe7.md)**|. Specifies whether the new values are in the rows or columns of the specified range.|
|SeriesLabels|Optional| **Variant**| **True** if the first row or column contains the name of the data series. **False** if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Excel attempts to determine the location of the series name from the contents of the first row or column.|
|CategoryLabels|Optional| **Variant**| **True** if the first row or column contains the name of the category labels. **False** if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Excel attempts to determine the location of the category label from the contents of the first row or column.|
|Replace|Optional| **Variant**|If CategoryLabels is **True** andReplace is **True**, the specified categories replace the categories that currently exist for the series. If Replace is **False**, the existing categories will not be replaced. The default value is  **False**.|

### Return Value

A  ** [Series](c7d34b32-8172-f7a0-0a17-f01d44246b64.md)** object that represents the new series.


## Remarks
<a name="sectionSection1"> </a>

This method does not actually return a  **Series** object as stated in the Object Browser. This method is not available for PivotChart reports.


## Example
<a name="sectionSection2"> </a>

This example creates a new series in Chart1. The data source for the new series is range B1:B10 on Sheet1.


```
Charts("Chart1").SeriesCollection.Add _ 
 Source:=ActiveWorkbook.Worksheets("Sheet1").Range("B1:B10")
```

This example creates a new series on the embedded chart on Sheet1.




```
Worksheets("Sheet1").ChartObjects(1).Activate 
ActiveChart.SeriesCollection.Add _ 
 Source:=Worksheets("Sheet1").Range("B1:B10")
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [SeriesCollection Object](93aa1f0b-4939-8c60-a444-2f791e8ce144.md)
#### Other resources


 [SeriesCollection Object Members](72d02a33-0b2b-1adb-9629-3eb322bed271.md)
