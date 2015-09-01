
# Determining Whether the Application Property Is Necessary

 **Last modified:** July 28, 2015

 _**Applies to:** Word 2013_

Many of the properties and methods of the  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object can be used without the **Application** object qualifier. For example the ** [ActiveDocument](c20a7c9f-f8a4-7913-f53f-10baa6807def.md)** property can be used without the **Application** object qualifier. Instead of writing `Application.ActiveDocument.PrintOut`, you can write  `ActiveDocument.PrintOut`.

Properties and methods that can be used without the  **Application** object qualifier are considered "global." To view the global properties and methods in the Object Browser, click **&lt;globals&gt;** at the top of the list in the **Classes** box.
