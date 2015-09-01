
# Image.BorderStyle Property (Outlook Forms Script)

 **Last modified:** July 28, 2015

Returns or sets an  **Integer** that specifies the type of border of the control. Read/write.

## Syntax

 _expression_. **BorderStyle**

 _expression_A variable that represents an  **Image** object.


## Remarks

The possible values of  **BorderStyle** are 0 and 1. 0 represents no visible border line, 1 represents a single-line border (default).

 The default value for an ** [Image](d2bcc281-6af0-5bbf-fa7f-ac581dbcf5dc.md)** is 1 (Single).

You can use either  **BorderStyle** or ** [SpecialEffect](174b4b27-a50f-da85-5ffe-91e268fce837.md)** to specify the border for a control, but not both. If you specify a nonzero value for one of these properties, the system sets the value of the other property to zero. For example, if you set **BorderStyle** to 1, the system sets **SpecialEffect** to zero (Flat). If you specify a nonzero value for **SpecialEffect**, the system sets  **BorderStyle** to zero.

 **BorderStyle** uses ** [BorderColor](5c0a373c-1ca7-1907-83b7-c24e9066e020.md)** to define the colors of its borders. To use the **BorderColor** property, the **BorderStyle** property must be set to a value other than 0.

