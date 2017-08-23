---
title: "Свойство Application.PathSeparator (издатель)"
keywords: vbapb10.chm131104
f1_keywords: vbapb10.chm131104
ms.prod: publisher
api_name: Publisher.Application.PathSeparator
ms.assetid: f8c07ce4-d171-9c5b-60ac-d544bf65e620
ms.date: 06/08/2017
ms.openlocfilehash: a004422bd6a15b0fdbc5ec9060a24db9eded2dfb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpathseparator-property-publisher"></a>Свойство Application.PathSeparator (издатель)

Возвращает **строку** , представляющую знак, используемый для разделения имена папок. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PathSeparator**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

**PathSeparator** можно использовать для создания веб-адреса, несмотря на то, что они содержат косую черту (/).

Свойство **[полное имя](document-fullname-property-publisher.md)** возвращает путь и имя файла в одну строку.

Для обеспечения совместимости по всему миру, мы рекомендуем использовать это свойство при построении пути, а не ссылается явно разделитель пути знаки в коде (например, «/»).


## <a name="example"></a>Пример

В этом примере отображается путь и имя активного документа.


```vb
Sub PathFileName() 
 
 With Application 
 MsgBox "The name of the active document: " &; vbLf &; _ 
 .Path &; .PathSeparator &; ActiveDocument.Name 
 End With 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

