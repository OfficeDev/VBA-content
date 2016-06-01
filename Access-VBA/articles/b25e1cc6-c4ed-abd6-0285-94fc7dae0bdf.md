
# BrowseTo Macro Action

 **Last modified:** July 28, 2015

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Setting](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)
[About the Contributors](#AboutContributors)


You can use the  **BrowseTo** action to navigate between objects in place. You can also change the source object of a subform control by specifying the **Path to Subform Control** argument. Use **BrowseTo** to navigate from form1 to form2 without opening up a new window.

## Setting
<a name="sectionSection0"> </a>

The  **BrowseTo** action has the following argument.



|**Action argument**|**Description**|
|:-----|:-----|
| _Object Type_|The object type to which to browse.|
| _Object Name_|The object that loads inside the subform control referenced by the  _Path to Subform Control_ argument.|
| _Path to Subform Control_|If specified, the path from the main form of the application to the target subform control that loads the object specified by the  _Object Name_ argument.|
| _Where Condition_|If specified, replaces the Where condition of the object record source.|
| _Page_|If specified, sets the page of the continuous form that will be made the current page. This argument is Web only.|
| _Data Mode_|If specified, the data entry mode of the form.|

## Remarks
<a name="sectionSection1"> </a>

The  _PathToSubFormControl_ argument must be specified using the syntax in the following code example:


```text
Main Form.SubForm Ctrl 1>Form 2.SubForm Ctrl 2>Form 3.SubFormCtrl3
```

In this example, the Main Form is the top level form in the Access client application. The  _Path to Sub Form Control_ argument must alternately specify form and subform control names leading from the main form to the subform control that is the container of the object specified by the _Object Name_ argument. Each subform control specified must be a control on the form that precedes it. The path must end with a subform control.


## Example
<a name="sectionSection2"> </a>

The following example shows how to use the  **BrowseTo** action to open a report in a subform control or within a navigation control.

 **Sample code provided by:** The[Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.mdl)




```text
OnError
    Go to Next
    Macro Name

/* Try to load the report in the host form (frmAuthorsParameter)    */
BrowseTo
    Object Type Report
    Object Name rptChapters
    Path to Subform Control frmAuthorsParameter.sfrmChild
    Where Condition
    Page
    Data Mode Edit

Parameters
    SelectedAuthor =[cboAuthor]

/* if this fails, try to load it in the navigation subform     */
BrowseTo
    Object Type Report
    Object Name rptChapters
    Path to Subform Control frmMain.NavigationSubform>frmAuthorsParameter.sfrmChild
    Where Condition
    Page
    Data Mode Edit

Parameters
    SelectedAuthor =[cboAuthor]
```


## About the Contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 

