---
title: Date Statement
keywords: vblr6.chm1008887
f1_keywords:
- vblr6.chm1008887
ms.prod: office
ms.assetid: 61cbe51b-f9a6-8d51-eba3-6d27a155b7c3
ms.date: 06/08/2017
---


# Date Statement

Sets the current system date.

 **Syntax**

 **Date** **=**_date_

For systems running Microsoft Windows 95, the required  _date_ specification must be a date from January 1, 1980 through December 31, 2099. For systems running Microsoft Windows NT, _date_ must be a date from January 1, 1980 through December 31, 2079. For the Macintosh, _date_ must be a date from January 1, 1904 through February 5, 2040.

## Example

This example uses the  **Date** statement to set the computer system date. In the development environment, the date literal is displayed in short date format using the locale settings of your code.


```vb
Dim MyDate 
MyDate = #February 12, 1985# ' Assign a date. 
Date= MyDate ' Change system date. 

```


