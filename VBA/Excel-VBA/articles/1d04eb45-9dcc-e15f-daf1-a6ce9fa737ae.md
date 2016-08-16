
# Application.AskToUpdateLinks Property (Excel)

 **True** if Microsoft Excel asks the user to update links when opening files with links. **False** if links are automatically updated with no dialog box. Read/write **Boolean** .


## Syntax

 _expression_ . **AskToUpdateLinks**

 _expression_ A variable that represents an **Application** object.


## Example

This example sets Microsoft Excel to ask the user to update links whenever a file that contains links is opened.


```vb
Application.AskToUpdateLinks = True
```


## See also


#### Concepts


[Application Object](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
