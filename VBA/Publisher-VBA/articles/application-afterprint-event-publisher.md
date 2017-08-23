---
title: "Событие Application.AfterPrint (издатель)"
keywords: vbapb10.chm268435492
f1_keywords: vbapb10.chm268435492
ms.prod: publisher
api_name: Publisher.Application.AfterPrint
ms.assetid: ddd5a1a4-8130-9e75-039c-e069a37390e8
ms.date: 06/08/2017
ms.openlocfilehash: 7d3c3d721d4c6679730d7f4cdf9bf52f0f2956b3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationafterprint-event-publisher"></a>Событие Application.AfterPrint (издатель)

Печать активируется после всех переменных и полей.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AfterPrint** ( **_Doc_**)

 _expression_An выражение, возвращающее объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Текущей публикации.|

## <a name="remarks"></a>Заметки

Microsoft Publisher не возвращает элемент управления пользовательского интерфейса пользователя, пока не выполняется обработчик событий. Событие вызывается после завершения всех графических операций (иными словами, после завершения задания программного обеспечения и оборудования печати начинает работать).

Дополнительные сведения об использовании событий с помощью объекта **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как обрабатывать события **AfterPrint** . Будет выведено сообщение о том, что печати документа.


```vb
Private Sub pubApplication_AfterPrint(ByVal Doc As Document) 
 MsgBox "Printing of " &; Doc.Name &; "is complete." 
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

