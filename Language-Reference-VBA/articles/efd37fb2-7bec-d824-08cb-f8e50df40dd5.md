
# Ways to protect sensitive information

 **Last modified:** July 28, 2015

Many applications use data that should be available only to certain users. Here are some suggestions you can use to protect sensitive information in Microsoft Forms:




- Write code that makes a control (and its data) invisible to unauthorized users. The  **Visible** property makes a control visible or invisible. For more information about **Visible**, see  [Visible Property](a81f2ebc-2d35-ca33-dce9-05256a1491c5.md).
    
- Write code that sets the control's foreground and background to the same color when unauthorized users run the application. This hides the information from unauthorized users. The  **ForeColor** and **BackColor** properties determine the [foreground color](7ce2c60f-29fb-96e2-2516-73c99a6e7cff.md) and the [background color](7ce2c60f-29fb-96e2-2516-73c99a6e7cff.md). For information about  **ForeColor**, see  [ForeColor Property](00b455d1-adce-ebb2-bb15-34cafebc5b75.md). For information about  **BackColor**, see  [BackColor Property](70549eaf-d785-67e7-3f04-76151864d850.md).
    
- Disable the control when unauthorized users run the application. The  **Enabled** property determines when a control is disabled. For information about **Enabled**, see  [Enabled Property](7e0320e4-91fa-2d2d-c484-70e54831e33b.md).
    
- Require a password for access to the application or a specific control. You can use  [placeholders](7ce2c60f-29fb-96e2-2516-73c99a6e7cff.md) as the user types each character. The **PasswordChar** property defines placeholder characters. For information about **PasswordChar**, see  [PasswordChar Property](2dd645b2-fe8d-a644-b796-e0595627cbb8.md).
    


 **Note**  Using passwords or any other techniques listed can improve the security of your application, but does not guarantee the prevention of unauthorized access to your data.

