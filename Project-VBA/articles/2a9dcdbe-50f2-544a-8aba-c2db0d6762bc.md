
# Project.ProjectNotes Property (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Gets or sets the notes for the project. Read/write  **String**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ProjectNotes**

 _expression_A variable that represents a  **Project** object.


## Remarks
<a name="sectionSection1"> </a>

To see the project  **Properties** dialog box in Project, choose the **FILE** tab on the ribbon to show the **Backstage** view, choose the **Info** tab, and then choose **Advanced Properties** in the **Project Information** drop-down menu.


## Example
<a name="sectionSection2"> </a>

The following example adds the date and time to the  **Comments** field in the project **Properties** dialog box, and then saves the project.


```
Sub SaveAndNoteTime() 
    Projects(1).ProjectNotes = Projects(1).ProjectNotes &amp; vbCrLf _ 
        &amp; "This project was last saved on " _ 
        &amp; Date$ &amp; " at " &amp; Time$ &amp; "." 
    FileSave 
End Sub
```

