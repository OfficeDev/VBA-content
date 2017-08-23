---
title: "Событие Application.DocumentOpen (издатель)"
keywords: vbapb10.chm268435463
f1_keywords: vbapb10.chm268435463
ms.prod: publisher
api_name: Publisher.Application.DocumentOpen
ms.assetid: 3bdd4b38-ec40-a08f-3742-f81a6ed333b3
ms.date: 06/08/2017
ms.openlocfilehash: f065ba80c37d2e88510fd00dcb9d16e8eefc7b73
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationdocumentopen-event-publisher"></a>Событие Application.DocumentOpen (издатель)

Происходит при открытии документа.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DocumentOpen** ( **_Doc_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Документ, открыта.|

## <a name="example"></a>Пример

В этом примере выводится сообщение с именем документа при открытии документа.


```vb
Private Sub appPub_DocumentOpen(ByVal Doc As Document) 
 MsgBox "Please wait. " &; Doc.Name &; " is opening." 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

