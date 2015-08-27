
# Page.Delete Method (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Deletes a  **Page** object. Can also renumber remaining pages.


## Syntax

 _expression_. **Delete**( **_fRenumberPages_**)

 _expression_A variable that represents a  **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|fRenumberPages|Required| **Integer**|1 ( **True**) to renumber remaining pages; otherwise, 0 ( **False**).|

### Return Value

Nothing


## Remarks

When fRenumberPages is non-zero, the remaining pages' default page names are renumbered after the page is deleted, otherwise, the pages retain their names.

