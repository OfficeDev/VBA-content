
# Table.TableFields Property (Project)

 **Last modified:** July 28, 2015

Gets a  ** [TableFields](7f749404-0723-7a17-b83f-f43725c45fc5.md)** collection representing the fields in the table. Read-only **TableFields**.

## Syntax

 _expression_. **TableFields**

 _expression_A variable that represents a  **Table** object.


## Example

The following example changes the alignment of a column in an entry table. The macro asks for input from the user to indicate which column the user wants to center, then changes the display and refreshes the view.


```
Sub AutoWrap() 
 Dim fieldNumber As Integer 
 
 fieldNumber = InputBox$(Prompt:="Enter the number of the " _ 
 &amp; "column you want to center in the Entry table." _ 
 &amp; Chr(13) &amp; "For example, Column 1 is the Indicators " _ 
 &amp; "column.") 
 
 ActiveProject.TaskTables("Entry").TableFields(fieldNumber _ 
 + 1).AlignData = pjCenter 
 
 TableApply Name:="&amp;Entry" 
End Sub
```

