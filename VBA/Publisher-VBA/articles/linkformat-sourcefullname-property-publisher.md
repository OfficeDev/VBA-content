---
title: "Свойство LinkFormat.SourceFullName (издатель)"
keywords: vbapb10.chm4390915
f1_keywords: vbapb10.chm4390915
ms.prod: publisher
api_name: Publisher.LinkFormat.SourceFullName
ms.assetid: a83aad48-ce27-6fe7-d26b-f00bec42e614
ms.date: 06/08/2017
ms.openlocfilehash: a9361e0763d7b8322b6cd7dc0913e36af6c08ddd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="linkformatsourcefullname-property-publisher"></a>Свойство LinkFormat.SourceFullName (издатель)

Возвращает **строку** , представляющую путь и имя исходного файла для указанного связанного объекта, рисунков или поля. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SourceFullName**

 переменная _expression_A, представляет собой объект- **LinkFormat** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере отображается путь и имя исходного файла для всех внедренных OLE фигур на первой странице active публикации.


```vb
Sub DisplaySourceName() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 If shp.Type = pbEmbeddedOLEObject Then 
 With shp.LinkFormat 
 MsgBox .SourceFullName 
 End With 
 End If 
 Next 
End Sub
```


