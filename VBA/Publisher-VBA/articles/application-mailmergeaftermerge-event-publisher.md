---
title: "Событие Application.MailMergeAfterMerge (издатель)"
keywords: vbapb10.chm268435465
f1_keywords: vbapb10.chm268435465
ms.prod: publisher
api_name: Publisher.Application.MailMergeAfterMerge
ms.assetid: dd01d8f5-f95e-e833-bb8b-708ced54240c
ms.date: 06/08/2017
ms.openlocfilehash: 8f312b58be17c50a0ee727ea3d246e630f849f1f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergeaftermerge-event-publisher"></a>Событие Application.MailMergeAfterMerge (издатель)

Происходит после успешного слияния всех записей в слияния почты.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeAfterMerge** ( **_Doc_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Основной документ слияния почты.|

## <a name="remarks"></a>Заметки

Для доступа к событий объекта **приложения** , объявите объектную переменную **приложения** в разделе Общие описаний модуля кода. Задайте переменную равно объект **приложения** , для которого требуется получить доступ к событиям. Сведения об использовании событий с помощью объекта Microsoft Publisher **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

В этом примере отображается сообщение о том, что будут завершены все записи в указанный документ объединения.


```vb
Private Sub MailMergeApp_MailMergeAfterMerge(ByVal Doc As Document) 
 
 MsgBox "Your mail merge on " &; _ 
 ActiveDocument.Name &; " is now finished." 
 
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

