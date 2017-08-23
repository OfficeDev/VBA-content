---
title: "Событие Application.BeforePrint (издатель)"
keywords: vbapb10.chm268435491
f1_keywords: vbapb10.chm268435491
ms.prod: publisher
api_name: Publisher.Application.BeforePrint
ms.assetid: 4d819aab-726e-ab00-89e0-aedcb62d834e
ms.date: 06/08/2017
ms.openlocfilehash: 98632fd1cf93ba3f698d3d9f4e777752fcb502f6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationbeforeprint-event-publisher"></a>Событие Application.BeforePrint (издатель)

Возникает перед публикацией печати или просмотра. .


## <a name="syntax"></a>Синтаксис

 _выражение_. **BeforePrint** ( **_Doc_**, **_Отменить_**)

 _expression_An выражение, возвращающее объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Текущей публикации.|
|Cancel|Обязательное свойство.| **Boolean**| **Значение false,** при возникновении события. Если этот параметр задает процедуру события значение **True**, публикации не печатается после завершения работы процедуры.|

## <a name="remarks"></a>Заметки

Событие **BeforePrint** вызывается только после того, как документ загружен полностью и вернули события onload. Печать не выполняется, пока выполняется обработчик событий.

Дополнительные сведения об использовании событий с помощью объекта **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как обрабатывать события **BeforePrint** . Будет выведено сообщение о том, что документ будет готов к печати.


```vb
Private Sub pubApplication_BeforePrint(ByVal Doc As Document, Cancel As Boolean ) 
 MsgBox "Printing of " &; Doc.Name &; " is about to occur ." 
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

