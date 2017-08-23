---
title: "Свойство Page.Master (издатель)"
keywords: vbapb10.chm393222
f1_keywords: vbapb10.chm393222
ms.prod: publisher
api_name: Publisher.Page.Master
ms.assetid: f206b4f1-cde3-458d-f26c-a970ad3bd21b
ms.date: 06/08/2017
ms.openlocfilehash: a14e6aaf54c8fa129fb21404b90299aae247ea64
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagemaster-property-publisher"></a>Свойство Page.Master (издатель)

Задает или возвращает объект **[страницы](page-object-publisher.md)** , представляющий главную страницу свойств для указанного страницы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Образец**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

Page


## <a name="remarks"></a>Заметки

Главные страницы не имеют **главных** свойства. Любая попытка доступа к свойству **главных** главной страницы приведет к ошибке времени выполнения.


## <a name="example"></a>Пример

В этом примере добавляется фигура на главную страницу для первой страницы в активной публикации.


```vb
Sub AddNewMasterPageShape() 
 With ActiveDocument.Pages(1).Master.Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=512, _ 
 Top:=50, Width:=50, Height:=50) 
 .Fill.ForeColor.CMYK.SetCMYK Cyan:=255, _ 
 Magenta:=255, Yellow:=0, Black:=0 
 End With 
End Sub
```

Свойство **главных** также может использоваться для применения главной страницы на страницу в публикации. В следующем примере задается главную страницу первой страницы публикации на главную страницу вторая страница публикации. В этом примере предполагается, что существует по крайней мере два страницы и двумя главными страницами в документе.




```vb
ActiveDocument.Pages(1).Master = _ 
 ActiveDocument.Pages(2).Master
```


