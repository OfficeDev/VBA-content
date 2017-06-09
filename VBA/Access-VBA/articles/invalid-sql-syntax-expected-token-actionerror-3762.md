---
title: "Invalid SQL Syntax: expected token: ACTION. (Error 3762)"
ms.prod: access
ms.assetid: 73122947-9db6-f417-7e34-96bc4108bab3
ms.date: 06/08/2017
---


# Invalid SQL Syntax: expected token: ACTION. (Error 3762)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when defining referential integrity constraints through the CREATE TABLE syntax or the ALTER TABLE ALTER COLUMN syntax. It occurs when the keyword NO is not followed by the keyword ACTION. For example, by omitting the BOLD ON keyword, the following would generate the error:

CREATE TABLE OrderDetail (OrderId LONG CONSTRAINT fkOrdersOrderId REFERENCES Orders (OrderId) ON UPDATE CASCADE ON DELETE  **NO** ACTION, LineItem LONG, ProductID LONG CONSTRAINT fkProductsProductId REFERENCES Products (ProductId), Quantity LONG);
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

