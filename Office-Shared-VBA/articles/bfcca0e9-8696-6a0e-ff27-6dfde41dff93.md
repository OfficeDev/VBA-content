
# IRibbonUI.InvalidateControlMso Method (Office)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Used to invalidate a built-in control.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **InvalidateControlMso**( **_ControlID_**)

 _expression_An expression that returns a  **IRibbonUI** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ControlID|Required| **String**||

### Return Value

Nothing


## Remarks
<a name="sectionSection1"> </a>

Invalidating a control repaints the screen and causes any callback procedures associated with that control to execute.


## Example
<a name="sectionSection2"> </a>


```XML
<customUI â€¦ OnLoad="MyAddInInitialize" â€¦>
```


```
Sub MyAddInInitialize(Ribbon As IRibbonUI) 
 Set MyRibbon = Ribbon 
End Sub 
 
Sub myFunction() 
 MyRibbon.InvalidateControlMso("TabInsert") ' Invalidates the Insert control 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [IRibbonUI Object](d323aa21-de74-e821-c914-db71ef3b9c5e.md)
#### Other resources


 [IRibbonUI Object Members](c6f6ec3b-3132-da29-ea08-70f20923d013.md)
