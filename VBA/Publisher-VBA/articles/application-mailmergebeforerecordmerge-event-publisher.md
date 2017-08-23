---
title: "Событие Application.MailMergeBeforeRecordMerge (издатель)"
keywords: vbapb10.chm268435474
f1_keywords: vbapb10.chm268435474
ms.prod: publisher
api_name: Publisher.Application.MailMergeBeforeRecordMerge
ms.assetid: 67ae8255-336d-0ff8-7927-fbd31262c115
ms.date: 06/08/2017
ms.openlocfilehash: a556d4fb62cb4a43a9c0f03117e968a177c50257
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergebeforerecordmerge-event-publisher"></a>Событие Application.MailMergeBeforeRecordMerge (издатель)

Происходит, как для отдельных записей выполняется слияние.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeBeforeRecordMerge** ( **_Doc_**, **_Отменить_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Основной документ слияния почты.|
|Cancel|Обязательное свойство.| **Boolean**|Останавливает процесс слияния почты для текущей записи только, до его запуска.|

## <a name="remarks"></a>Заметки

Для доступа к событий объекта **приложения** , объявите объектную переменную **приложения** в разделе Общие описаний модуля кода. Задайте переменную равно объект **приложения** , для которого требуется получить доступ к событиям. Сведения об использовании событий с помощью объекта Microsoft Publisher **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

В этом примере выполняется проверка, что длина ПОЧТОВЫЙ индекс (который в этом примере — поля номер шести) меньше 5 и если он установлен, отменяет merge для этой записи только.


```vb
Private Sub MailMergeApp_MailMergeBeforeRecordMerge(ByVal _ 
 Doc As Document, Cancel As Boolean) 
 
 Dim intZipLength As Integer 
 
 intZipLength = Len(ActiveDocument.MailMerge _ 
 .DataSource.DataFields(6).Value) 
 
 'Cancel merge of this record only if 
 'the ZIP Code has fewer than five digits 
 If intZipLength < 5 Then 
 Cancel = True 
 End If 
 
End Sub
```

Чтобы произошло это событие необходимо поместить следующую строку кода в разделе глобальные описаний модуля и выполнить следующую процедуру инициализации.




```vb
Private WithEvents MailMergeApp As Application 
 
Sub InitializeMailMergeApp() 
 Set MailMergeApp = Publisher.Application 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

