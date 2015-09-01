
# TextBox.AutoSize Property (Outlook Forms Script)

 **Last modified:** July 28, 2015

Returns or sets a  **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.

## Syntax

 _expression_. **AutoSize**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

 **True** to automatically resize the control to display its entire contents. **False** to keep the size of the control constant; contents are clipped when they exceed the area of the control (default).

For controls without captions, this property specifies whether the control automatically adjusts to display the information stored in the control. In a  ** [ComboBox](31e7c1de-ee4e-b3d9-4579-7fc6b215bad3.md)**, for example, setting  **AutoSize** to **True** automatically sets the width of the display area to match the length of the current text.

If you manually change the size of a control while  **AutoSize** is **True**, the manual change overrides the size previously set by  **AutoSize**.

For a single-line  ** [TextBox](4a0e4a3d-beca-9f94-7e27-469c4bafe250.md)**, setting  **AutoSize** to **True** automatically sets the width of the display area to the length of the text in the text box.

For a multiline  **TextBox** that contains no text, setting **AutoSize** to **True** automatically displays the text as a column. The width of the text column is set to accommodate the widest letter of that font size. The height of the text column is set to display the entire text of the **TextBox**.

For a multiline  **TextBox** that contains text, setting **AutoSize** to **True** automatically enlarges the **TextBox** vertically to display the entire text. The width of the **TextBox** does not change.

