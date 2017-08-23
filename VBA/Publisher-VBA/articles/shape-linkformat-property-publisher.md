---
title: "Свойство Shape.LinkFormat (издатель)"
keywords: vbapb10.chm2228326
f1_keywords: vbapb10.chm2228326
ms.prod: publisher
api_name: Publisher.Shape.LinkFormat
ms.assetid: 801c3a87-7cc6-8c7b-094a-55e8d8d7a004
ms.date: 06/08/2017
ms.openlocfilehash: 3db7e05d9f093086f23e2dbe1579f77219bf2961
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapelinkformat-property-publisher"></a>Свойство Shape.LinkFormat (издатель)

Возвращает объект [LinkFormat](linkformat-object-publisher.md), который содержит свойства, которые являются уникальными для связанные объекты OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LinkFormat**

 переменная _expression_A, представляющий объект **фигуры** .


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


