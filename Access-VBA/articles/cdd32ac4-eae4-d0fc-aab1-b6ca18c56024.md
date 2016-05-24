
# The language-specific code page was not specified or could not be found. (Error 3649)

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You have attempted to open a database that was created with a language that is not installed on your computer. You should determine what language was specified for this database when it was created and then make sure that language is installed on your system. If the database was created with DAO, the language was specified with the locale argument of the  **CreateDatabase** method. If the database was created with Microsoft Access, the language was specified with the option "New Database Sort Order" on the **General** tab of the **Options** dialog box, which is available by clicking **Options** on the **Tools** menu.

Languages can be added to your system through the Regional settings of the Control Panel.
