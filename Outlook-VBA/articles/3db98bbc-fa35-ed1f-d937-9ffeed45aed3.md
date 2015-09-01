
# ComboBox.Text Property (Outlook Forms Script)

 **Last modified:** July 28, 2015

Returns or sets a  **String** that specifies text in a ** [ComboBox](31e7c1de-ee4e-b3d9-4579-7fc6b215bad3.md)**, changing the selected row in the control. Read/write.

## Syntax

 _expression_. **Text**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The default value is a zero-length string ("").

You can use  **Text** to update the value of the control. If the value of **Text** matches an existing list entry, the value of the ** [ListIndex](2c4e473b-15e1-dce2-8748-30953b00a60f.md)** property (the index of the current row) is set to the row that matches **Text**. If the value of  **Text** does not match a row, **ListIndex** is set to -1.

When the  **Text** property of a **ComboBox** changes (such as when a user types an entry into the control), the new text is compared to the column of data specified by ** [TextColumn](5ebf37ef-4cec-ec42-d42f-ab886b86e913.md)**.

You cannot use  **Text** to change the value of an entry in a **ComboBox**; use the  ** [Column](f00c388f-fe1f-5458-281f-4bfa549291d5.md)** or ** [List](687f44e8-7b4b-eab5-93b8-022cd4d1c302.md)** property for this purpose.

The  ** [ForeColor](256d695a-df00-d22c-b2aa-e21036beea35.md)** property determines the color of the text.

