
# Tasks.ExitWindows Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Closes all open applications, quits Microsoft Windows, and logs the current user off.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ExitWindows**

 _expression_Required. A variable that represents a  ** [Tasks](ff521e20-8a25-f9f6-dccf-effea9debeb7.md)** collection.


## Remarks
<a name="sectionSection1"> </a>

This method does not save changes to open Microsoft Word documents; however, it does prompt you to save changes to open documents in other Windows-based applications.


## Example
<a name="sectionSection2"> </a>

This example saves all open Word documents, closes Word, and then quits Microsoft Windows.


```
Documents.Save NoPrompt:=True, _ 
 OriginalFormat:=wdOriginalDocumentFormat 
Tasks.ExitWindows
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Tasks Collection Object](ff521e20-8a25-f9f6-dccf-effea9debeb7.md)
#### Other resources


 [Tasks Object Members](e6ca78c6-132d-6e7b-9f83-ea044a395040.md)
