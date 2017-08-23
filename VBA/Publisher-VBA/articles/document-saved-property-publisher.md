---
title: "Свойство Document.Saved (издатель)"
keywords: vbapb10.chm196649
f1_keywords: vbapb10.chm196649
ms.prod: publisher
api_name: Publisher.Document.Saved
ms.assetid: d1f4357a-103c-2227-d1bd-50706e1f241c
ms.date: 06/08/2017
ms.openlocfilehash: 75d61abd6ef14efe4a3f2748cec1ca8b6dc628fd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentsaved-property-publisher"></a>Свойство Document.Saved (издатель)

Возвращает **значение True,** Если изменения не были внесены в публикацию с момента последнего сохранения. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Сохранить**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Если свойству **Saved** измененную публикацию возвращает **значение True**, пользователь не будет предложено сохранить изменения, при закрытии публикации, и будут потеряны все изменения, внесенные с момента последнего сохранения.


## <a name="example"></a>Пример

В этом примере сохраняет active публикации, если она была изменена со времени последнего сохранения.


```vb
Sub Saver() 
 
 With Application.ActiveDocument 
 If Not .Saved And .Path <> "" Then .Save 
 End With 
 
End Sub
```


