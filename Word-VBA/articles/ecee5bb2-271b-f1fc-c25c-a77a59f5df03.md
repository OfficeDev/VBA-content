
# Application.ShowStartupDialog Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** to display the **Task Pane** when starting Microsoft Word. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ShowStartupDialog**

 _expression_A variable that represents an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object.


## Remarks
<a name="sectionSection1"> </a>

The  **ShowStartupDialog** property is a global option, and the new setting will take effect only after you restart Word. Use the **Visible**property of the  **CommandBars** collection show or hide the Task Pane without restarting Word.


## Example
<a name="sectionSection2"> </a>

This example turns off the  **Task Pane**, so it won't display upon starting Word. This will not take effect until the next time the user starts Word.


```
Sub HideStartUpDlg() 
 Application.ShowStartupDialog = False 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Application Object](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)
#### Other resources


 [Application Object Members](71669f1e-65f1-b0f1-b67d-355dfdbebe50.md)
