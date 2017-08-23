---
title: "Свойство MailMergeDataSource.TableName (издатель)"
keywords: vbapb10.chm6291491
f1_keywords: vbapb10.chm6291491
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.TableName
ms.assetid: 0418bf66-550e-7dfc-671f-db2570a768d9
ms.date: 06/08/2017
ms.openlocfilehash: 49eb6bb5dbb71c687dbbd718cdcdfcb2515ffcde
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcetablename-property-publisher"></a>Свойство MailMergeDataSource.TableName (издатель)

Возвращает **строку** , представляющую имя таблицы в данных исходный файл, содержащий записи слияния почты. Возвращаемое значение может быть пустым, если имя таблицы неизвестна или не применим к текущему источнику данных. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TableName**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере выводится сообщение с именем имя таблицы источника данных слияния почты.


```vb
Sub EmployeeTable() 
 With ActiveDocument.MailMerge.DataSource 
 Select Case .TableName 
 Case "Employees" 
 MsgBox "This is an Employee mail merge publication." 
 Case "Customers" 
 MsgBox "This is a Customers mail merge publication." 
 Case "Suppliers" 
 MsgBox "This is a Suppliers mail merge publication." 
 Case "Shippers" 
 MsgBox "This is a Shippers mail merge publication." 
 Case Else 
 MsgBox "This is a " &; .TableName &; " mail merge publication." 
 End Select 
 End With 
End Sub
```


