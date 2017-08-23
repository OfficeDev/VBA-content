---
title: "Свойство ShapeRange.LinkFormat (издатель)"
keywords: vbapb10.chm2293862
f1_keywords: vbapb10.chm2293862
ms.prod: publisher
api_name: Publisher.ShapeRange.LinkFormat
ms.assetid: 1f0add8d-7baa-65f0-e82b-a047a7bc0507
ms.date: 06/08/2017
ms.openlocfilehash: c56e66a6bd68aa25da71d46e622a3cc6bb80950f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangelinkformat-property-publisher"></a>Свойство ShapeRange.LinkFormat (издатель)

Возвращает объект [LinkFormat](linkformat-object-publisher.md), который содержит свойства, которые являются уникальными для связанные объекты OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LinkFormat**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="example"></a>Пример

В этом примере обновляется связи между любые объекты OLE по одному в активной публикации и исходные файлы.


```vb
Dim sh As Shape 
 
For Each sh In ActiveDocument.Pages(1).Shapes 
 If sh.Type = pbLinkedOLEObject Then 
 With sh.LinkFormat 
 .Update 
 End With 
 End If 
Next
```


