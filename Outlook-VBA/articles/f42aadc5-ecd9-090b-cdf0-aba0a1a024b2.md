
# TextBox.MultiLine Property (Outlook Forms Script)

 **Last modified:** July 28, 2015

Returns or sets a  **Boolean** that specifies whether a control can accept and display multiple lines of text. Read/write.

## Syntax

 _expression_. **MultiLine**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

True if the text is displayed across multiple lines (default). Falase if the text is not displayed across multiple lines.

A multiline  ** [TextBox](4a0e4a3d-beca-9f94-7e27-469c4bafe250.md)** allows absolute line breaks and adjusts its quantity of lines to accommodate the amount of text it holds. If needed, a multiline control can have vertical scroll bars.

A single-line  **TextBox** doesn't allow absolute line breaks and doesn't use vertical scroll bars.

For controls that support the  **MultiLine** property as well as the ** [WordWrap](fb50b340-9fe7-17b5-4f5f-d2fdd266f37d.md)** property, **WordWrap** is ignored when **MultiLine** is **False**.

Single-line controls ignore the value of the  **WordWrap** property.

If you change  **MultiLine** to **False** in a multiline **TextBox**, all the characters in the  **TextBox** will be combined into one line, including non-printing characters (such as carriage returns and new-lines).

The  ** [EnterKeyBehavior](2af4a64e-4939-ae46-0d25-67fe986d413a.md)** and **MultiLine** properties are closely related. The **EnterKeyBehavior** values of **True** and **False** only apply if **MultiLine** is **True**. If  **MultiLine** is **False**, pressing  **ENTER** always moves the focus to the next control in the tab order regardless of the value of **EnterKeyBehavior**.

The effect of pressing  **CTRL+ENTER** also depends on the value of **MultiLine**. If  **MultiLine** is **True**, pressing  **CTRL+ENTER** creates a new line regardless of the value of **EnterKeyBehavior**. If  **MultiLine** is **False**, pressing  **CTRL+ENTER** has no effect.

The  ** [TabKeyBehavior](5b8bdc3c-9000-a7fd-af39-743cc117e02d.md)** and **MultiLine** properties are closely related. The values described above only apply if **MultiLine** is **True**. If  **MultiLine** is **False**, pressing  **TAB** always moves the focus to the next control in the tab order regardless of the value of **TabKeyBehavior**.

The effect of pressing  **CTRL+TAB** also depends on the value of **MultiLine**. If  **MultiLine** is **True**, pressing  **CTRL+TAB** creates a new line regardless of the value of **TabKeyBehavior**. If  **MultiLine** is **False**, pressing  **CTRL+TAB** has no effect.

