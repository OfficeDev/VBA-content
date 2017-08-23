---
title: "Метод Document.ConvertPublicationType (издатель)"
keywords: vbapb10.chm196737
f1_keywords: vbapb10.chm196737
ms.prod: publisher
api_name: Publisher.Document.ConvertPublicationType
ms.assetid: e4bfe349-a22f-6017-ac9d-49f67e1f6dd2
ms.date: 06/08/2017
ms.openlocfilehash: b6185e05239970aae7d3c6974f15e59a3db871e9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentconvertpublicationtype-method-publisher"></a>Метод Document.ConvertPublicationType (издатель)

Преобразует тип указанного публикации указанной публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ConvertPublicationType** ( **_Значение_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **PbPublicationType**|Тип публикации, который будет преобразован публикации.|

## <a name="remarks"></a>Заметки

При преобразовании публикации все параметры, применимые к предыдущей типа остаются, но игнорируются. Например преобразование печатной публикации в веб-публикации результатов поиска в любые дополнительные настройки печати игнорируются. При публикации преобразуется обратно в печатной публикации, вступят в силу еще раз.

Используйте свойство **[PublicationType](document-publicationtype-property-publisher.md)** объекта **[Document](document-object-publisher.md)** для определения типа публикации публикации.

Значение параметра может иметь одно из следующих **PbPublicationType** константы, описанные в библиотеке типов, Microsoft Publisher.



| **pbTypePrint**|| **pbTypeWeb**|

## <a name="example"></a>Пример

Следующий пример определяет, является ли активная публикация печати публикации. Если он установлен, публикация преобразуется в веб-публикации.


```vb
Sub ChangePublicationType() 
 With ActiveDocument 
 If .PublicationType = pbTypePrint Then 
 .ConvertPublicationType (pbTypeWeb) 
 End If 
 End With 
End Sub
```


