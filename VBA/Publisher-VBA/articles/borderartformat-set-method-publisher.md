---
title: "Метод BorderArtFormat.Set (издатель)"
keywords: vbapb10.chm7602185
f1_keywords: vbapb10.chm7602185
ms.prod: publisher
api_name: Publisher.BorderArtFormat.Set
ms.assetid: e068037b-56b6-a114-6b22-568ea20d6b25
ms.date: 06/08/2017
ms.openlocfilehash: 648da309ad7dcdc9a0fe8322619381fa12890ff5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartformatset-method-publisher"></a>Метод BorderArtFormat.Set (издатель)

Задает тип Узорные применена к указанной фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Установка** ( **_BorderArtName_**)

 переменная _expression_A, представляет собой объект- **BorderArtFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|BorderArtName|Обязательное свойство.| **Variant**|Имя типа Узорные, применяемые к указанной фигуры.|

## <a name="remarks"></a>Заметки

Также можно задать тип Узорные, применяемые к фигуре с помощью свойства **[Name](borderartformat-name-property-publisher.md)** .


## <a name="example"></a>Пример

Следующий пример проверяет наличие Узорные на каждой фигуры для каждой страницы активных документов. Все найденные Узорные присвоено значение того же типа.


```vb
Sub SetBorderArt() 
Dim anyPage As Page 
Dim anyShape As Shape 
Dim strBorderArtName As String 
 
strBorderArtName = Document.BorderArts(1).Name 
 
For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .Set(strBorderArtName) 
 End If 
 End With 
 Next anyShape 
 Next anyPage 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект BorderArtFormat](borderartformat-object-publisher.md)

