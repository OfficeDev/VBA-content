---
title: "Событие Document.BeforeClose (издатель)"
keywords: vbapb10.chm285212674
f1_keywords: vbapb10.chm285212674
ms.prod: publisher
api_name: Publisher.Document.BeforeClose
ms.assetid: d40e36b6-fea7-a9d5-0c88-55197983b888
ms.date: 06/08/2017
ms.openlocfilehash: 8edc9973a0effc742aa56a95761f4989d5fdb674
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentbeforeclose-event-publisher"></a>Событие Document.BeforeClose (издатель)

Происходит непосредственно перед закрытием любого открытого документа.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BeforeClose** ( **_Отмена_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Cancel|Обязательное свойство.| **Boolean**| **Значение false,** при возникновении события. Если этот аргумент задает процедуру события значение **True**, документ не закрыть после завершения процедуры.|

## <a name="remarks"></a>Заметки

Дополнительные сведения об использовании событий с помощью объекта **Document** содержатся в разделе [С помощью событий с помощью объекта Document](using-events-with-the-document-object-publisher.md).


## <a name="example"></a>Пример

В этом примере пользователю Да или нет ответа перед закрытием документа. Для работы этого примера необходимо поместить этот код в модуле **ThisDocument** .


```vb
Private Sub Document_BeforeClose(Cancel As Boolean) 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really want to close " _ 
 &; "the document?", vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub
```


