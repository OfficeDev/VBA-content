
# DoCmd.BrowseTo Method (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


The  **BrowseTo** method performs the BrowseTo action in Visual Basic.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **BrowseTo**( **_ObjectType_**,  **_ObjectName_**,  **_PathtoSubformControl_**,  **_WhereCondition_**,  **_Page_**,  **_DataMode_**)

 _expression_A variable that represents a  **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ObjectType|Required| ** [AcBrowseToObjectType](52388196-e1e3-f199-24e8-04b399d55cfb.md)**|The object type to which to browse.|
|ObjectName|Required| **Variant**|The object that loads inside the subform control referenced by the PathtoSubformControl argument. |
|PathtoSubformControl|Optional| **Variant**|If specified, the path from the main form of the application to the target subform control that loads the object specified by the ObjectName argument.|
|WhereCondition|Optional| **Variant**|If specified, replaces the Where condition of the object record source.|
|Page|Optional| **Variant**|If specified, sets the page of the continuous form that will be made the current page. This argument is Web only.|
|DataMode|Optional| ** [AcFormOpenDataMode](24c39abb-154c-39cd-3097-77be75fe917c.md)**|If specified, the data entry mode of the form.|

## Remarks
<a name="sectionSection1"> </a>

Use the  **BrowseTo** method to navigate between objects in place. You can also change the source object of a subform control by specifying the PathToSubFormControl argument. You can use **BrowseTo** to navigate from form1 to form2 without opening up a new window.

The PathToSubFormControl argument must be specified using the syntax in the following example:

Main Form.SubForm Ctrl 1>Form 2.SubForm Ctrl 2>Form 3.SubFormCtrl3

In this example, the Main Form is the top level form in the Access client application. The PathToSubFormControl argument must alternately specify form and subform control names leading from the main form to the subform control that is the container of the object specified by the ObjectName argument. Each subform control specified must be a control on the form that precedes it. The path must end with a subform control.


## Example
<a name="sectionSection2"> </a>

The following code example opens the "EventDS" form in place in edit mode in the "NavigationSubform" subform control of the "Main" form.


```
DoCmd.BrowseTo ObjectType:=acBrowseToForm, _ 
ObjectName:="EventDS", _ 
PathToSubformControl:="Main.NavigationSubform", _ 
WhereCondition:="", _ 
Page:="", _ 
DataMode:=acFormEdit
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [DoCmd Object](3ce44cca-9979-0a1e-9787-079a52ce528f.md)
#### Other resources


 [DoCmd Object Members](3e7ade9e-86e4-0751-188b-5d31c9101651.md)
