---
title: The Microsoft Access database engine does not recognize <name> as a valid field name or expression. (Error 3070)
keywords: jeterr40.chm5003070
f1_keywords:
- jeterr40.chm5003070
ms.prod: access
ms.assetid: 8866f9ea-4c2b-45f6-9ec7-8e23596efbf9
ms.date: 06/08/2017
---


# The Microsoft Access database engine does not recognize <name> as a valid field name or expression. (Error 3070)

  

**Applies to:** Access 2013 | Access 2016

The specified name is not a recognized field name or a valid expression. In a query, this error can occur if you enter a name that improperly refers to a database, table, or field.

Possible causes with Microsoft Access:


- You have a parameter in a crosstab query or in a query that a crosstab query or chart is based on, and the parameter data type is not explicitly specified in the  **Query Parameters** dialog box. To solve the problem:
    
    
    
      - In the query that contains the parameter, specify the parameter and its data type in the  **Query Parameters** dialog box. And;
    
  - Set the  **ColumnHeadings** property for the query that contains the parameter.
    

    
    
- In any type of query, you have improperly referred to a database, table, or field. For example, this error can occur if you refer to a field named Salary in an expression, but you misspell the field name, such as  `[Sallary]*1.1`.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

