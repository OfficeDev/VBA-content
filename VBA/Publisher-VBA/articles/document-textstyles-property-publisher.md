---
title: "Свойство Document.TextStyles (издатель)"
keywords: vbapb10.chm196662
f1_keywords: vbapb10.chm196662
ms.prod: publisher
api_name: Publisher.Document.TextStyles
ms.assetid: a628e5c1-aed7-dd70-81fa-d9fb54afb527
ms.date: 06/08/2017
ms.openlocfilehash: 5ec29c4b5d9ed87c8dc009412cae529a8bb70e95
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documenttextstyles-property-publisher"></a>Свойство Document.TextStyles (издатель)

Возвращает коллекцию **[TextStyles](textstyles-object-publisher.md)** , содержащий стили текста публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextStyles**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

TextStyles


## <a name="example"></a>Пример

Следующий пример отображает имя стиля и базового стиля первый стиль в коллекции **TextStyles** .


```vb
Sub BaseStyleName() 
 With ActiveDocument.TextStyles(1) 
 MsgBox "Style name= " &; .Name _ 
 &; vbCr &; "Base style= " &; .BaseStyle 
 End With 
End Sub
```


