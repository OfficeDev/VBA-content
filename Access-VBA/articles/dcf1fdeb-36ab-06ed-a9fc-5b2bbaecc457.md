
# Window.SmallScroll Method (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Scrolls the contents of the window by rows or columns.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SmallScroll**( **_Down_**,  **_Up_**,  **_ToRight_**,  **_ToLeft_**)

 _expression_A variable that represents a  **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Down|Optional| **Variant**|The number of rows to scroll the contents down.|
|Up|Optional| **Variant**|The number of rows to scroll the contents up.|
|ToRight|Optional| **Variant**|The number of columns to scroll the contents to the right.|
|ToLeft|Optional| **Variant**|The number of columns to scroll the contents to the left.|

### Return Value

Variant


## Remarks
<a name="sectionSection1"> </a>

If Down andUp are both specified, the contents of the window are scrolled by the difference of the arguments. For example, ifDown is 3 andUp is 6, the contents are scrolled up three rows.

If ToLeft andToRight are both specified, the contents of the window are scrolled by the difference of the arguments. For example, ifToLeft is 3 andToRight is 6, the contents are scrolled to the right three columns.

Any of these arguments can be a negative number.


## Example
<a name="sectionSection2"> </a>

This example scrolls the contents of the active window of Sheet1 down three rows.


```
Worksheets("Sheet1").Activate 
ActiveWindow.SmallScroll down:=3
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Window Object](8591b1ad-76f8-14e2-9120-406b65093f5a.md)
#### Other resources


 [Window Object Members](f11db427-24a4-041c-2fd5-03ce73ae6c16.md)
