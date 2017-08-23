---
title: "Событие Application.MailMergeDataSourceValidate (издатель)"
keywords: vbapb10.chm268435480
f1_keywords: vbapb10.chm268435480
ms.prod: publisher
api_name: Publisher.Application.MailMergeDataSourceValidate
ms.assetid: 8e18b0a0-8fe8-f72e-8a75-1585367cc796
ms.date: 06/08/2017
ms.openlocfilehash: a31b68277639cfbe63fbb36056864fdc2c12e8ac
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergedatasourcevalidate-event-publisher"></a>Событие Application.MailMergeDataSourceValidate (издатель)

Происходит, когда пользователь выполняет проверку адреса, нажав кнопку **Проверить** в диалоговом окне **Получатели слияния** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeDataSourceValidate** ( **_Doc_**, **_обрабатывается_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Основной документ слияния почты.|
|Обработано|Обязательное свойство.| **Boolean**| **Значение true,** выполняется сопутствующий код проверки источника данных для слияния. **False** отменяет проверку источника данных.|

## <a name="remarks"></a>Заметки

Если у вас на компьютере программное обеспечение для проверки адреса, использование **событием MailMergeDataSourceValidate, чтобы создать простое** фильтрации процедур, таких как циклический перебор записей для проверки почтового индекса и удалить все, что пользователи не - США не в США можно отфильтровать все почтовые индексы США путем изменения в образце кода и с помощью Microsoft Visual Basic команды для поиска текста или специальные символы.

Для доступа к событий объекта **приложения** , объявите объектную переменную **приложения** в разделе Общие описаний модуля кода. Задайте переменную равно объект **приложения** , для которого требуется получить доступ к событиям. Сведения об использовании событий с помощью объекта Microsoft Publisher **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

В этом примере выполняется проверка индексы в источник данных для пяти цифр. Если длина ПОЧТОВЫЙ индекс составляет менее пяти цифр, записи исключается из процесс слияния почты. В этом примере предполагается, что почтовые индексы являются ПОЧТОВЫЕ индексы США. Можно изменить в этом примере для поиска индексы, содержащие код локатор из четырех цифр, добавляется в конец ПОЧТОВЫЙ индекс и затем исключить все записи, не содержащих кода локатор.


```vb
Private Sub MailMergeApp_MailMergeDataSourceValidate( _ 
 ByVal Doc As Document, _ 
 Handled As Boolean) 
 
 Dim intCount As Integer 
 
 Handled = True 
 
 On Error Resume Next 
 
 With ActiveDocument.MailMerge.DataSource 
 
 'Set the active record equal to the first included record in the 
 'data source 
 .ActiveRecord = 1 
 Do 
 intCount = intCount + 1 
 
 'Set the condition that field six must be greater than or 
 'equal to five 
 If Len(.DataFields.Item(6).Value) < 5 Then 
 
 'Exclude the record if field six is shorter than five digits 
 .Included = False 
 
 'Mark the record as containing an invalid address field 
 .InvalidAddress = True 
 
 'Specify the comment attached to the record explaining 
 'why the record was excluded from the mail merge 
 .InvalidComments = "The ZIP Code for this record has " _ 
 &; "fewer than five digits. It will be removed " _ 
 &; "from the mail merge process." 
 
 End If 
 
 'Move the record to the next record in the data source 
 .ActiveRecord = .ActiveRecord + 1 
 
 'End the loop when the counter variable 
 'equals the number of records in the data source 
 Loop Until intCount = .RecordCount 
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

