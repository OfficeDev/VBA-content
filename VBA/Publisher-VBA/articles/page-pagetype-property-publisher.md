---
title: "Свойство Page.PageType (издатель)"
keywords: vbapb10.chm393221
f1_keywords: vbapb10.chm393221
ms.prod: publisher
api_name: Publisher.Page.PageType
ms.assetid: 0bb34de5-ac3e-386c-3b9f-814a476c9695
ms.date: 06/08/2017
ms.openlocfilehash: 758c339a6488daf46951e20e347116d2495892fa
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagepagetype-property-publisher"></a>Свойство Page.PageType (издатель)

Возвращает константу **PbPageType** , представляющий тип страницы. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageType**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

PbPageType


## <a name="remarks"></a>Заметки

Значение свойства **PageType** может иметь одно из следующих **PbPageType** константы, описанные в библиотеке типов, Microsoft Publisher.



| **pbPageLeftPage**|| **pbPageMasterPage**|| **pbPageRightPage**|| **pbPageScratchPage**|

## <a name="example"></a>Пример

В этом примере добавляется фигура на разный углов каждой страницы в активной публикации.


```vb
Sub GetPageType() 
 Dim pgCount As Page 
 For Each pgCount In ActiveDocument.Pages 
 If pgCount.PageType = pbPageLeftPage Then 
 pgCount.Shapes.AddShape Type:=msoShapeOval, _ 
 Left:=50, Top:=50, Width:=50, Height:=50 
 ElseIf pgCount.PageType = pbPageRightPage Then 
 pgCount.Shapes.AddShape Type:=msoShapeOval, _ 
 Left:=512, Top:=50, Width:=50, Height:=50 
 End If 
 Next 
End Sub
```


