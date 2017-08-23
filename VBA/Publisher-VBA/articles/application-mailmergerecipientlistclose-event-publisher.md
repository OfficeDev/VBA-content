---
title: "Событие Application.MailMergeRecipientListClose (издатель)"
keywords: vbapb10.chm268435488
f1_keywords: vbapb10.chm268435488
ms.prod: publisher
api_name: Publisher.Application.MailMergeRecipientListClose
ms.assetid: 4fb77771-9897-8623-f4e7-61f631f04922
ms.date: 06/08/2017
ms.openlocfilehash: f9fce9fa5fbe63c25f617754a034117fbd9a0103
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergerecipientlistclose-event-publisher"></a>Событие Application.MailMergeRecipientListClose (издатель)

Активируется, когда пользователь закрывает диалоговое окно **Получатели слияния** . (Из области задач **слияния почты** и **Объединение электронной почты** , нажмите кнопку **изменить получателя**). Также применяется, когда пользователь закрывает диалоговое окно **Списка продуктов объединение каталога** , которая открывается при нажатии кнопки на панели задач **Объединение в каталог** **Изменить список продуктов** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeRecipientListClose** ( **_Doc_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Текущей публикации.|

## <a name="remarks"></a>Заметки

Дополнительные сведения об использовании событий с помощью объекта **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как обрабатывать события **MailMergeRecipientListClose** . Будет выведено сообщение о том, что отображенные строки, описанных выше.


```vb
Private Sub pubApplication_MailMergeRecipientListClose(ByVal Doc As Document) 
 MsgBox "The Mail Merge Recipients dialog box has closed." 
End Sub
```

Для чтобы произошло это событие необходимо включить следующую строку кода в разделе **Общие описаний** модуля.




```vb
Private WithEvents pubApplication As Application
```

Затем выполните следующую процедуру инициализации.




```vb
Public Sub Initialize_pubApplication() 
 Set pubApplication = Publisher.Application 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

