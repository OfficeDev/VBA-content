---
title: "Метод MailMerge.OpenDataSource (издатель)"
keywords: vbapb10.chm6225937
f1_keywords: vbapb10.chm6225937
ms.prod: publisher
api_name: Publisher.MailMerge.OpenDataSource
ms.assetid: 4473e566-687f-595e-9fd6-a5483021cb48
ms.date: 06/08/2017
ms.openlocfilehash: 9e5bd11342f93b80e851f7528da2382be1661663
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergeopendatasource-method-publisher"></a>Метод MailMerge.OpenDataSource (издатель)

Подключает указанный публикации, который становится основной публикации, если она еще не источника данных.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OpenDataSource** ( **_bstrDataSource_**, **_bstrConnect_**, **_bstrTable_**, **_fOpenExclusive_**, **_fNeverPrompt_**)

 переменная _expression_A, представляет собой объект- **слияния** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|bstrDataSource|Необязательный| **String**|Данные источника путь и имя файла. Можно указать файл Microsoft Query (.qry) вместо определения источника данных, строка подключения и строка имя таблицы; значения в файле Microsoft Query переопределить значения для bstrConnect и bstrTable.|
|bstrConnect|Необязательный| **String**|Строка подключения.|
|bstrTable|Необязательный| **String**|Имя таблицы в источнике данных.|
|fOpenExclusive|Необязательный| **Длинный**| **Значение true,** чтобы запретить другим пользователям доступ к базе данных. **Значение false** позволяет другим пользователям чтения/записи разрешений для базы данных. Значение по умолчанию — **False**.|
|fNeverPrompt|Необязательный| **Длинный**| **Длинные**.  **Значение true,** никогда не выдает запрос при открытии источника данных. **Значение false,** отображается диалоговое окно Свойства связи с данными. Значение по умолчанию — **False**.|

## <a name="remarks"></a>Заметки




 **Примечание**  При использовании источника данных для слияния почты необходимо добавить область объединения в каталог на странице публикации перед присоединением к источнику данных.


## <a name="example"></a>Пример

В этом примере связывает таблицу из базы данных и не дает все еще доступ на запись в базу данных при открытии. 

В этом примере для ведения необходимо заменить _PathToFile_ допустимый путь к файлу и _TableName_ с именем таблицы источника допустимых данных.




```vb
Sub AttachDataSource() 
 
    ActiveDocument.MailMerge.OpenDataSource _ 
        bstrDataSource:="PathToFile",  _ 
        bstrTable:="TableName", _ 
        fNeverPrompt:=True, fOpenExclusive:=True 
 
End Sub
```


