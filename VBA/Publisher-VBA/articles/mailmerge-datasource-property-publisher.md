---
title: "Свойство MailMerge.DataSource (издатель)"
keywords: vbapb10.chm6225923
f1_keywords: vbapb10.chm6225923
ms.prod: publisher
api_name: Publisher.MailMerge.DataSource
ms.assetid: 19b32513-fd57-617a-38e2-6230e3e036b9
ms.date: 06/08/2017
ms.openlocfilehash: 4acee980021143fe25bd03981ff2fa7b903ecdc8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasource-property-publisher"></a>Свойство MailMerge.DataSource (издатель)

Возвращает объект **[вывода](mailmergedatasource-object-publisher.md)** , который относится к источнику данных, подключенного к публикации главного слиянием слияния почты и каталогов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Источник данных**

 переменная _expression_A, представляет собой объект- **слияния** .


### <a name="return-value"></a>Возвращаемое значение

Вывода


## <a name="example"></a>Пример

В этом примере отображается путь и имя источника данных, подключенного к active публикации.


```vb
Sub DataSourceName() 
 With ActiveDocument.MailMerge.DataSource 
 If .Name <> "" Then _ 
 MsgBox "The path and file name of the " &; _ 
 "attached data source is : " &; vbCr &; .Name 
 End With 
End Sub
```


