---
title: "Объект InstalledPrinters (издатель)"
keywords: vbapb10.chm8978431
f1_keywords: vbapb10.chm8978431
ms.prod: publisher
api_name: Publisher.InstalledPrinters
ms.assetid: 8cf9b194-70bc-7963-6a08-d08401d4b6f3
ms.date: 06/08/2017
ms.openlocfilehash: b51e9db4a666838a7f21f30027fdab150661ee1d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="installedprinters-object-publisher"></a>Объект InstalledPrinters (издатель)

Представляет коллекцию объектов все **принтера** , каждая из которых представляет один из принтеров, установленных на компьютере.
 


## <a name="remarks"></a>Заметки

Чтобы предоставить пользователю возможность выбора принтера для печати публикации, можно выполнять итерации по коллекции **InstalledPrinters** для получения списка имен всех принтеров, установленных на компьютере, как показано в следующем примере.
 

 
Свойство по умолчанию коллекции **InstalledPrinters** — **элемента**.
 

 

## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как можно использовать свойства **[Имя_принтера](printer-printername-property-publisher.md)** и **[IsActivePrinter](printer-isactiveprinter-property-publisher.md)** объекта **принтера** для получения списка всех установленных принтеров на компьютере и определить, какой из них в настоящее время активного принтера.
 

 

```
Public Sub InstalledPrinters_Example() 
 
 Dim pubInstalledPrinters As Publisher.InstalledPrinters 
 Dim pubApplication As Publisher.Application 
 Dim pubPrinter As Publisher.Printer 
 
 Set pubApplication = ThisDocument.Application 
 Set pubInstalledPrinters = pubApplication.InstalledPrinters 
 
 For Each pubPrinter In pubInstalledPrinters 
 Debug.Print pubPrinter.PrinterName 
 If pubPrinter.IsActivePrinter Then 
 Debug.Print "This is the active printer." 
 End If 
 Next 
 
End Sub 

```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](installedprinters-application-property-publisher.md)|
|[Count](installedprinters-count-property-publisher.md)|
|[Элемент](installedprinters-item-property-publisher.md)|
|[Родительский раздел](installedprinters-parent-property-publisher.md)|

