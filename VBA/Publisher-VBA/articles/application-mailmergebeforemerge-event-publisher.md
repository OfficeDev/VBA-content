---
title: "Событие Application.MailMergeBeforeMerge (издатель)"
keywords: vbapb10.chm268435473
f1_keywords: vbapb10.chm268435473
ms.prod: publisher
api_name: Publisher.Application.MailMergeBeforeMerge
ms.assetid: 735ef282-e99f-b3f2-c509-b180bea30d36
ms.date: 06/08/2017
ms.openlocfilehash: 93bc7ab8244105361bc316c46da69286f842f292
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergebeforemerge-event-publisher"></a>Событие Application.MailMergeBeforeMerge (издатель)

Происходит, когда выполняется слияние, перед объединенные записи при слиянии.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeBeforeMerge** ( **_Doc_**, **_StartRecord_** **_EndRecord_**, **_Отменить_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Основной документ слияния почты.|
|Параметр StartRecord|Обязательное свойство.| **Длинный**|Первой записи в источник данных, чтобы включить в слияния почты.|
|EndRecord|Обязательное свойство.| **Длинный**|Последней записи в источник данных, чтобы включить в слияния почты.|
|Cancel|Обязательное свойство.| **Boolean**|Останавливает процесс слияния почты до его запуска.|

## <a name="remarks"></a>Заметки

Для доступа к событий объекта **приложения** , объявите объектную переменную **приложения** в разделе Общие описаний модуля кода. Задайте переменную равно объект **приложения** , для которого требуется получить доступ к событиям. Сведения об использовании событий с помощью объекта Microsoft Publisher **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

В этом примере выводится сообщение, перед началом процесса слияния почты вопросом, следует ли продолжить. При нажатии кнопки Нет процесс слияния отменяется.


```vb
Private Sub MailMergeApp_MailMergeBeforeMerge(ByVal Doc As Document, _ 
 ByVal StartRecord As Long, ByVal EndRecord As Long, _ 
 Cancel As Boolean) 
 
 Dim intVBAnswer As Integer 
 
 Set Doc = ActiveDocument 
 
 'Request whether the user wants to continue with the merge 
 intVBAnswer = MsgBox("Mail Merge for " &; Doc.Name &; _ 
 " is now starting. Do you want to continue?", _ 
 vbYesNo, "Event!") 
 
 'If user's response to question is No, then cancel merge process 
 'and deliver a message to the user stating the merge is canceled 
 If intVBAnswer = vbNo Then 
 Cancel = True 
 MsgBox "You have canceled mail merge for " &; _ 
 Doc.Name &; "." 
 End If 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

