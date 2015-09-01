
# TextBox.BorderColor Property (Outlook Forms Script)

 **Last modified:** July 28, 2015

Returns or sets a  **Long** that specifies the border color of an object. Read/write.

## Syntax

 _expression_. **BorderColor**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```
RGB(15,200,75)
```

To use the  **BorderColor** property, the ** [BorderStyle](c71b8117-a731-d0ab-89a7-84dd9aa089c4.md)** property must be set to a value other than 0.

 **BorderStyle** uses **BorderColor** to define the border colors. The ** [SpecialEffect](b7365d4e-c25d-9fa6-c088-0cc5bb6bb200.md)** property uses system colors exclusively to define its border colors. For Windows operating systems, system color settings are set using the **Display** icon in Control Panel.

