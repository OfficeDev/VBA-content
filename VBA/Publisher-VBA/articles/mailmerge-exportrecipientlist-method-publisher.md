---
title: "Метод MailMerge.ExportRecipientList (издатель)"
keywords: vbapb10.chm6225941
f1_keywords: vbapb10.chm6225941
ms.prod: publisher
api_name: Publisher.MailMerge.ExportRecipientList
ms.assetid: 230d0f66-7368-51b7-8233-3fd54cfd0fe4
ms.date: 06/08/2017
ms.openlocfilehash: ff0e685b2b16aa3926c65fd36640ee1db0a94f36
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergeexportrecipientlist-method-publisher"></a>Метод MailMerge.ExportRecipientList (издатель)

Экспорт в список получателей слияния почты в файл Microsoft Office Access (MDB) или в файл с разделителями-запятыми (.csv).


## <a name="syntax"></a>Синтаксис

 _выражение_. **ExportRecipientList** ( **_Имя файла_**, **_Тип файла_**, **_IncludedOnly_**)

 переменная _expression_A, представляет собой объект- **слияния** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя файла|Обязательное свойство.| **String**|Имя файла, содержащего список получателей.|
|FileType|Необязательный| **PbRecipientListFileType**|Тип файла для сохранения. Возможные значения см.|
|IncludedOnly|Необязательный| **Boolean**|Следует ли ограничения записей в список получателей к определенным получателям.|

## <a name="remarks"></a>Заметки

Возможные значения для параметра FileType включают следующие константы из перечисления **PbRecipientListFileType** :



|**Константы**|**Значение**|**Описание**|
|:-----|:-----|:-----|
| **pbAsCsvFile**|1|Сохраните как CSV-файла с разделителями запятыми.|
| **pbAsMdbFile**|0|Сохраните как Microsoft Office Access MDB-файлу.|
Метод **ExportRecipientList** соответствует команду **Экспорт списка получателей в новый файл** в **Слияния почты** и **слияния почты** областями задач в интерфейсе пользователя Microsoft Publisher.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **ExportRecipientList** для экспорта в список получателей слияния почты в файле базы данных Access. Прежде чем запустить этот макрос, убедитесь, что активный документ подключен к источнику данных. Если активный документ не подключен к источнику данных, можно использовать ** [MailMerge.OpenDataSource](mailmerge-opendatasource-method-publisher.md)** метод для подключения.

Кроме того перед выполнением кода замените _имя пользователя_ в путь к папке на сохраненный файл с именем допустимого пользователя на вашем компьютере или замените путь к папке и имя файла, в которое включен путь и имя файла.

Обратите внимание, что путь к папке, в этом примере типичные для путей к папкам в Microsoft Windows Vista. Необходимо иметь разрешение на сохранение файлов в папке.




```vb
Public Sub ExportRecipientList_Example() 
 
 Dim pubMailMerge As Publisher.MailMerge 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 pubMailMerge.ExportRecipientList "C:\Users\username\Documents\My Data Sources\MyAddressList", pbAsMdbFile, True 
 
End Sub
```


