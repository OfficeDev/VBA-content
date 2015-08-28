
# TabControl.MultiRow Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


You can use the  **MultiRow** property to specify or determine whether a tab control can display more than one row of tabs. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **MultiRow**

 _expression_A variable that represents a  **TabControl** object.


## Remarks
<a name="sectionSection1"> </a>

The  **MultiRow** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes| **True**|Multiple rows are allowed.|
|No| **False**|(Default) Multiple rows aren't allowed.|
You can also set the default for this property by setting a control's  **DefaultControl**property in Visual Basic.

When the  **MultiRow** property is set to **True**, the number of rows is determined by the width and number of tabs. The number of rows may change if the control is resized or if additional tabs are added to the control.

When the  **MultiRow** property is set to **False** and the width of the tabs exceeds the width of the control, navigation buttons appear on the right side of the tab control. You can use the navigation buttons to scroll through all the tabs on the tab control.


## Example
<a name="sectionSection2"> </a>

To return the value of the  **MultiRow** property for a tab control named "Details" on the "Order Entry" form, you can use the following:


```
Dim b As Boolean 
b = Forms("Order Entry").Controls("Details").MultiRow
```

To set the value of the  **MultiRow** property, you can use the following:




```
Forms("Order Entry").Controls("Details").MultiRow = True
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [TabControl Object](05f7de7b-8665-df6d-3fbb-47f8547d3baf.md)
#### Other resources


 [TabControl Object Members](d6de9ec4-e7f9-5c26-d750-d7c134ec9fb0.md)
