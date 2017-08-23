---
title: "Свойство PageSetup.AvailablePageSizes (издатель)"
keywords: vbapb10.chm6946849
f1_keywords: vbapb10.chm6946849
ms.prod: publisher
api_name: Publisher.PageSetup.AvailablePageSizes
ms.assetid: 5ad79ee6-6d32-6c46-c02e-a9ab252267cf
ms.date: 06/08/2017
ms.openlocfilehash: a918034390616eb804a08952c76e3bc9771a4225
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetupavailablepagesizes-property-publisher"></a>Свойство PageSetup.AvailablePageSizes (издатель)

Возвращает коллекцию **PageSizes** , которая содержит все объекты **[PageSize](pagesize-object-publisher.md)** , доступных в текущей публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AvailablePageSizes**

 переменная _expression_A, представляет собой объект- **PageSetup** .


### <a name="return-value"></a>Возвращаемое значение

PageSizes


## <a name="remarks"></a>Заметки

 Объекты **PageSize** соответствующие значки, отображаемые в разделе **Пустая страница размеры** в диалоговом окне **Параметры страницы** в интерфейсе пользователя Microsoft Publisher.

Размер страницы, возвращаемой свойством **AvailablePageSizes** включают не только размеры страниц, предоставляемых Microsoft Publisher, но размеров настраиваемые страницы, создайте или загрузить, при их наличии.

Как добавить или удалить пользовательскую страницу размеры, изменить номер индекса для всех существующих размеров страницы. 


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как для создания текстового файла, содержащего список всех размер бумаги в текущей публикации и их соответствующих значений индекса. Она сохраняет текстовый файл документов (в Windows Vista) или папка Мои документы (в Windows XP) текущего пользователя.


```vb
Public Sub AvailablePageSizes_Example() 
 
 Dim pubPageSize As Publisher.PageSize 
 Dim pubPageSizes As Publisher.PageSizes 
 Dim intCount As Integer 
 Dim lngPageSizeFile As Long 
 
 intCount = 1 
 
 Set pubPageSizes = ThisDocument.PageSetup.AvailablePageSizes 
 
 lngPageSizeFile = FreeFile 
 Open Environ("USERPROFILE") + "\My Documents\PageSizeList.txt" For Output Access Write As lngPageSizeFile 
 
 For Each pubPageSize In pubPageSizes 
 Write #lngPageSizeFile, pubPageSize.Name, intCount 
 intCount = intCount + 1 
 Next 
 
 Close lngPageSizeFile 
 
End Sub
```


