
# Categories.Item Method (Outlook)

 **Last modified:** July 28, 2015

Returns a  ** [Category](143ef095-54b0-cbe2-e356-632029061ac2.md)** object from the collection.

## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **Categories** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|Either a  **Long** value representing the index number of the object, or a **String** value representing either the ** [Name](b9a711e9-f79d-f4f7-88bb-eaeb61d64089.md)** or ** [CategoryID](e75ed17a-940f-2325-8739-1367329854d2.md)** property value of an object in the collection.|

### Return Value

A  **Category** object that represents the specified object.


## Remarks

If the name of a category is specified in Index, this method returns the first  **Category** object that matches the specified value. If a match cannot be found, the method returns **Null** ( **Nothing** in Visual Basic.)


## See also


#### Concepts


 [Categories Object](319efa26-269d-9f2f-c8ec-33082e80a9e2.md)
#### Other resources


 [Categories Object Members](36fd8906-69fa-5aa8-b026-a2de208ccd56.md)
