
# MailMessage.Forward Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Opens a new e-mail message with an empty  **To** line for forwarding the active message.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Forward**

 _expression_Required. A variable that represents a  ** [MailMessage](d0109969-27f7-0180-c56d-5b49a3f0171b.md)** object.


## Remarks
<a name="sectionSection1"> </a>

This method is available only if you are using Word as your e-mail editor.


## Example
<a name="sectionSection2"> </a>

This example opens a new e-mail message for forwarding the active message.


```
Application.MailMessage.Forward
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [MailMessage Object](d0109969-27f7-0180-c56d-5b49a3f0171b.md)
#### Other resources


 [MailMessage Object Members](7e52ff10-90a9-5752-5adb-c70de2837165.md)
