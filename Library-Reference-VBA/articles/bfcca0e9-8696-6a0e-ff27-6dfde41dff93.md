
# IRibbonUI.InvalidateControlMso Method (Office)

Used to invalidate a built-in control.


## Syntax

 _expression_. **InvalidateControlMso**( ** _ControlID_** )

 _expression_ An expression that returns a **IRibbonUI** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ControlID_|Required|**String**||

### Return Value

Nothing


## Remarks

Invalidating a control repaints the screen and causes any callback procedures associated with that control to execute.


## Example


```XML
<customUI … OnLoad="MyAddInInitialize" …>
```


```vb
Sub MyAddInInitialize(Ribbon As IRibbonUI) 
 Set MyRibbon = Ribbon 
End Sub 
 
Sub myFunction() 
 MyRibbon.InvalidateControlMso("TabInsert") ' Invalidates the Insert control 
End Sub
```


## See also


#### Concepts


[IRibbonUI Object](d323aa21-de74-e821-c914-db71ef3b9c5e.md)
#### Other resources


[IRibbonUI Object Members](c6f6ec3b-3132-da29-ea08-70f20923d013.md)