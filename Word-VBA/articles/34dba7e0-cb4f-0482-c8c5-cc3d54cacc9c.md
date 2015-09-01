
# Panes.Add Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  **Pane** object that represents a new pane to a window.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Add**( **_SplitVertical_**)

 _expression_Required. A variable that represents a  ** [Panes](6ed6353c-9134-f47d-a108-13e84eced8ff.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|SplitVertical|Optional| **Variant**|A number that represents the percentage of the window, from top to bottom, you want to appear above the split.|

### Return Value

Pane


## Remarks
<a name="sectionSection1"> </a>

This method will fail if it is applied to a window that has already been split.


## Example
<a name="sectionSection2"> </a>

The following example splits the active window such that the top pane is 30 percent of the total window size.


```
ActiveDocument.ActiveWindow.Panes.Add SplitVertical:=30
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Panes Collection Object](6ed6353c-9134-f47d-a108-13e84eced8ff.md)
#### Other resources


 [Panes Object Members](22673447-a48d-afea-0642-5eb2a3efd221.md)
