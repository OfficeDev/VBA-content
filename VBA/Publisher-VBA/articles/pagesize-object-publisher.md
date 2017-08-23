---
title: "Объект PageSize (издатель)"
keywords: vbapb10.chm8912895
f1_keywords: vbapb10.chm8912895
ms.prod: publisher
api_name: Publisher.PageSize
ms.assetid: 80767524-6f0c-0d3f-388a-a38891b2d04a
ms.date: 06/08/2017
ms.openlocfilehash: 24b7e9db2e48c8d12bed0d449eeb30add10be5da
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesize-object-publisher"></a>Объект PageSize (издатель)

Представляет размер страницы публикации Microsoft Publisher.


## <a name="remarks"></a>Заметки

Размер страницы, представленного объектом **PageSize** соответствует одному значков, отображаемых в разделе **Пустая страница размеры** в диалоговом окне **Параметры страницы** в пользовательском интерфейсе Publisher.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать свойство **Name** объекта **PageSize** для получения списка имен всех странице размеры доступно в текущем документе и напечатайте список в окне **Интерпретация** .


```
Public Sub PageSizes_Example() 
 
 Dim pubPageSizes As Publisher.PageSizes 
 Dim pubPageSize As Publisher.PageSize 
 
 Set pubPageSizes = ThisDocument.PageSetup.AvailablePageSizes 
 For Each pubPageSize In pubPageSizes 
 Debug.Print pubPageSize.Name 
 Next 
 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](http://msdn.microsoft.com/library/397e9db8-e12d-55bb-0b34-406e0c3666e0%28Office.15%29.aspx)|
|[HasBackgroundImage](http://msdn.microsoft.com/library/544e8e73-e134-c297-42da-bc96c3d498e0%28Office.15%29.aspx)|
|[HorizontalGap](http://msdn.microsoft.com/library/14c14534-c1c7-db2d-c7bf-8b7fd66c245e%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/e1cb706e-6b0e-a7c2-494f-3e77717215cb%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/7ed8d2d1-7aab-ec6a-f24a-a93bb05dcdfd%28Office.15%29.aspx)|
|[PageHeight](http://msdn.microsoft.com/library/25cfa836-9109-f360-ee6c-a6824639c911%28Office.15%29.aspx)|
|[PageWidth](http://msdn.microsoft.com/library/5b8d9f75-06b6-51a8-8463-57eac69f0197%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/3a141bb0-9fd7-3522-7ea2-0a51fe2a6b10%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/1d1755c2-bb53-5bc2-002c-93714df13784%28Office.15%29.aspx)|
|[VerticalGap](http://msdn.microsoft.com/library/cc6e66ff-9a74-d88f-cfde-2f5bee66432f%28Office.15%29.aspx)|

