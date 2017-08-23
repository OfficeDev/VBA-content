---
title: "Свойство Document.PublicationType (издатель)"
keywords: vbapb10.chm196736
f1_keywords: vbapb10.chm196736
ms.prod: publisher
api_name: Publisher.Document.PublicationType
ms.assetid: 264c2769-2452-0009-4853-84a6a426db38
ms.date: 06/08/2017
ms.openlocfilehash: f3a3f906633ad2587d3b66de5d0208b13174dcac
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentpublicationtype-property-publisher"></a>Свойство Document.PublicationType (издатель)

Возвращает константу **PbPublicationType** , представляющий тип указанной публикации. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PublicationType**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

PbPublicationType


## <a name="remarks"></a>Заметки

Значение свойства **PublicationType** может иметь одно из следующих констант **PbPublicationType** .



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


