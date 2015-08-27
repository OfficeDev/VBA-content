
# Font.Italic Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets an  **MsoTriState** constant indicating whether the specified text is formatted as italic. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Italic**

 _expression_A variable that represents an  **Font** object.


### Return Value

MsoTriState


## Remarks
<a name="sectionSection1"> </a>

The  **Italic** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|None of the characters in the range are formatted as italic.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|All of the characters in the range are formatted as italic.|

## Example
<a name="sectionSection2"> </a>

This example tests all the text in the second story of the active publication and, if it has some text formatted as italic, it sets all the text to italic. If the text is all italic or all not italic, a message is displayed informing the user that there is no mixed italic formatting.


```
Sub ItalicStory() 
 
 Dim stf As Font 
 
 Set stf = Application.ActiveDocument.Stories(2).TextRange.Font 
 With stf 
 If .Italic = msoTriStateMixed Then 
 .Italic = msoTrue 
 Else 
 MsgBox "There is no mixed italic formatting in this story." 
 End If 
 End With 
 
End Sub
```

