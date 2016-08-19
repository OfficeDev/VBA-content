
# Tasks.ExitWindows Method (Word)

Closes all open applications, quits Microsoft Windows, and logs the current user off.


## Syntax

 _expression_ . **ExitWindows**

 _expression_ Required. A variable that represents a **[Tasks](ff521e20-8a25-f9f6-dccf-effea9debeb7.md)** collection.


## Remarks

This method does not save changes to open Microsoft Word documents; however, it does prompt you to save changes to open documents in other Windows-based applications.


## Example

This example saves all open Word documents, closes Word, and then quits Microsoft Windows.


```
Documents.Save NoPrompt:=True, _ 
 OriginalFormat:=wdOriginalDocumentFormat 
Tasks.ExitWindows
```


## See also


#### Concepts


[Tasks Collection Object](ff521e20-8a25-f9f6-dccf-effea9debeb7.md)
