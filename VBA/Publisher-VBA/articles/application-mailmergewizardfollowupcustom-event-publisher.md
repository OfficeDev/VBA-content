---
title: "Событие Application.MailMergeWizardFollowUpCustom (издатель)"
keywords: vbapb10.chm268435490
f1_keywords: vbapb10.chm268435490
ms.prod: publisher
api_name: Publisher.Application.MailMergeWizardFollowUpCustom
ms.assetid: ac8cb695-69a4-83f7-8e13-66762f52f611
ms.date: 06/08/2017
ms.openlocfilehash: e83e913830c9451a576651dce668e5afdf5fdd74
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergewizardfollowupcustom-event-publisher"></a>Событие Application.MailMergeWizardFollowUpCustom (издатель)

Возникает при нажатии строку, которая отображается как четвертый элемента в разделе **Подготовка к отслеживанию результатов рассылки** на **третий области задач в интерфейсе пользователя Microsoft Publisher** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeWizardFollowUpCustom** ( **_Doc_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Текущей публикации.|

## <a name="remarks"></a>Заметки

Свойство **[ShowFollowUpCustom](application-showfollowupcustom-property-publisher.md)** используется для отображения этой строки.

Дополнительные сведения об использовании событий с помощью объекта **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как обрабатывать события **MailMergeWizardFollowUpCustom** . Будет выведено сообщение о том, что отображенные строки, описанных выше.


```vb
Private Sub pubApplication_MailMergeWizardFollowUpCustom(ByVal Doc As Document) 
 MsgBox "The FollowUpCustom string is clicked." 
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

