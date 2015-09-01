
# EmailAuthor.Style Property (Word)

 **Last modified:** July 28, 2015

Returns a  **Style** object that represents the style associated with the current e-mail author for unsent replies, forwards, or new e-mail messages.

## Syntax

 _expression_. **Style**

 _expression_Required. A variable that represents an  ** [EmailAuthor](2749e018-42e9-7a1a-f18b-8605b38ff0ae.md)** object.


## Example

This example returns the style associated with the current author for unsent replies, forwards, or new e-mail messages and displays the name of the font associated with this style.


```
Set MyEmailStyle = _ 
 ActiveDocument.Email.CurrentEmailAuthor.Style 
Msgbox MyEmailStyle.Font.Name
```


## See also


#### Concepts


 [EmailAuthor Object](2749e018-42e9-7a1a-f18b-8605b38ff0ae.md)
#### Other resources


 [EmailAuthor Object Members](76ddf916-7e7f-4a5a-3330-cdb47e2b4d1c.md)
