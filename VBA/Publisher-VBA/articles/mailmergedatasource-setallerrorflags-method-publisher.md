---
title: "Метод MailMergeDataSource.SetAllErrorFlags (издатель)"
keywords: vbapb10.chm6291488
f1_keywords: vbapb10.chm6291488
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.SetAllErrorFlags
ms.assetid: 17c41fbb-3b21-c31a-63cd-ed26065bfa79
ms.date: 06/08/2017
ms.openlocfilehash: 41f32f083926e7d36fbd02a456dfbeafdcb44660
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcesetallerrorflags-method-publisher"></a>Метод MailMergeDataSource.SetAllErrorFlags (издатель)

Помечает все записи в источнике данных, содержащее недопустимые данные в поле адрес.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetAllErrorFlags** ( **_Недопустимые_**, **_InvalidComment_**)

 переменная _expression_A, представляющий объект **вывода** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Недопустимый|Обязательное свойство.| **Boolean**| **Значение true,** помечает все записи в источнике данных слияния почты как недопустимый.|
|InvalidComment|Необязательный| **String**|Текст с описанием недопустимый параметр.|

## <a name="remarks"></a>Заметки

Отдельно можно пометить записей в источнике данных, содержащих недопустимые данные в поле адреса с помощью свойства **[InvalidAddress](mailmergedatasource-invalidaddress-property-publisher.md)** и **[InvalidComments](mailmergedatasource-invalidcomments-property-publisher.md)** .


## <a name="example"></a>Пример

В этом примере все записи в источнике данных, содержащее поле недопустимый адреса, задает комментарий, почему он является недопустимым и исключает все записи из слияния почты.


```vb
Sub FlagAllRecords() 
 With ActiveDocument.MailMerge.DataSource 
 .SetAllErrorFlags Invalid:=True, InvalidComment:= _ 
 "All records in the data source have only 5-" _ 
 &; "digit ZIP Codes. Need 5+4 digit ZIP Codes." 
 .SetAllIncludedFlags Included:=False 
 End With 
End Sub
```


