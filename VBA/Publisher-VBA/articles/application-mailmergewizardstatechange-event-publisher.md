---
title: "Событие Application.MailMergeWizardStateChange (издатель)"
keywords: vbapb10.chm268435479
f1_keywords: vbapb10.chm268435479
ms.prod: publisher
api_name: Publisher.Application.MailMergeWizardStateChange
ms.assetid: 3d3fcdaa-af51-0a28-ff25-f2b92deceaf6
ms.date: 06/08/2017
ms.openlocfilehash: 01e7efbb9eab048e35d289a646bcfea503c18905
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergewizardstatechange-event-publisher"></a>Событие Application.MailMergeWizardStateChange (издатель)

Происходит при изменении пользователем из указанного действия в указанном этап в мастере слияния почты.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeWizardStateChange** ( **_Doc_**, **_FromState_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Основной документ слияния почты.|
|FromState|Обязательное свойство.| **Integer**|Мастер слияния шаг, из которой перемещается пользователь.|

## <a name="remarks"></a>Заметки

Для доступа к событий объекта **приложения** , объявите объектную переменную **приложения** в разделе Общие описаний модуля кода. Задайте переменную равно объект **приложения** , для которого требуется получить доступ к событиям.


## <a name="example"></a>Пример

В этом примере выводится сообщение при перемещении пользователей в шаге 3 мастера слияния в шаге 4. На основании ответов пользователя к сообщению, пользователь будет перейти на шаге 4 или вернуться к шагу 3.


```vb
Private Sub MailMergeApp_MailMergeWizardStateChange(ByVal Doc As Document, _ 
 ByVal FromState As Long) 
 
 Select Case FromState 
 Case 1 
 MsgBox "Now you will build your publication merge " &; _ 
 "by adding fields to your publication." 
 Case 2 
 MsgBox "Now you will see your publication " &; _ 
 "merged with the records in the data source." 
 Case 3 
 MsgBox "Now you will complete the mail merge process." 
 End Select 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

