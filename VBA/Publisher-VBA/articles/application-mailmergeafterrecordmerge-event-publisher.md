---
title: "Событие Application.MailMergeAfterRecordMerge (издатель)"
keywords: vbapb10.chm268435472
f1_keywords: vbapb10.chm268435472
ms.prod: publisher
api_name: Publisher.Application.MailMergeAfterRecordMerge
ms.assetid: 550c3310-01ba-718f-4c1d-cbf3ce077d27
ms.date: 06/08/2017
ms.openlocfilehash: bcf892f6b4e8751565493728401b98ddbb32d1aa
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergeafterrecordmerge-event-publisher"></a>Событие Application.MailMergeAfterRecordMerge (издатель)

Происходит после успешного слияния каждой записи в источнике данных.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeAfterRecordMerge** ( **_Doc_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Основной документ слияния почты.|

## <a name="remarks"></a>Заметки

При обслуживании базы данных управления клиента событие **MailMergeAfterRecordMerge** используется для обновления базы данных для каждой из объединенных записей.

Для доступа к событий объекта **приложения** , объявите объектную переменную **приложения** в разделе Общие описаний модуля кода. Задайте переменную равно объект **приложения** , для которого требуется получить доступ к событиям. Сведения об использовании событий с помощью объекта Microsoft Publisher **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

Этот пример отображает сообщение со значением поля первый и второй записи, которая только что завершения объединения.


```vb
Private Sub MailMergeApp_MailMergeAfterRecordMerge(ByVal Doc As Document) 
 
 With ActiveDocument.MailMerge.DataSource 
 MsgBox .DataFields.Item(3).Value &; " " &; _ 
 .DataFields.Item(2).Value &; " is finished merging." 
 End With 
 
End Sub
```

Чтобы произошло это событие необходимо поместить следующую строку кода в разделе Общие описаний модуля и выполнить следующую процедуру инициализации.




```vb
Private WithEvents MailMergeApp As Application 
 
Sub InitializeMailMergeApp() 
 Set MailMergeApp = Publisher.Application 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

