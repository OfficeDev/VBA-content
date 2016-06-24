
# Project.ProjectNotes Property (Project)

Gets or sets the notes for the project. Read/write  **String**.


## Syntax

 _expression_. **ProjectNotes**

 _expression_ A variable that represents a **Project** object.


## Remarks

To see the project  **Properties** dialog box in Project, choose the **FILE** tab on the ribbon to show the **Backstage** view, choose the **Info** tab, and then choose **Advanced Properties** in the **Project Information** drop-down menu.


## Example

The following example adds the date and time to the  **Comments** field in the project **Properties** dialog box, and then saves the project.


```vb
Sub SaveAndNoteTime() 
    Projects(1).ProjectNotes = Projects(1).ProjectNotes &; vbCrLf _ 
        &; "This project was last saved on " _ 
        &; Date$ &; " at " &; Time$ &; "." 
    FileSave 
End Sub
```

