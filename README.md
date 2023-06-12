# ExcelBuilder
ExcelBuilder, C# ile Excel dosyalarını oluşturmayı, düzenlemeyi ve veri aktarımını kolaylaştıran bir yardımcı sınıftır. Bu sınıf, EPPlus kütüphanesini kullanarak Excel dosyalarıyla etkileşim sağlar.

## Özellikler
* Excel dosyası oluşturma veya mevcut bir şablona dayalı olarak düzenleme
* Satır ve sütun ekleme, kopyalama ve silme işlemleri
* Hücre verilerini ayarlama, biçimlendirme ve formül eklemeyi destekleme
* Veri koleksiyonlarını hızlı bir şekilde Excel tablosuna aktarma
* Excel dosyasını farklı formatlarda (xlsx, csv) kaydetme
* Excel dosyasından veri aktarımı ve dönüştürme

## Bu projede çeşitli programlama prensipleri ve kalıpları kullanılmıştır:
### 1. SOLID Prensipleri:
* **Single Responsibility Principle (Tek Sorumluluk Prensibi):** Her metot ve sınıf, bir işlevi yerine getirmek için tasarlanmıştır ve tek bir sorumluluğu vardır. Örneğin, SetRow, SetColumn, SetDataList gibi yöntemler belirli bir işlevi yerine getirmek için tasarlanmıştır.
* **Open/Closed Principle (Açık/Kapalı Prensibi):** Sınıf, yeni işlevselliği eklemek için genişletilebilir (inheritance) ve mevcut işlevselliği değiştirmek yerine yeni kod eklenerek genişletilebilir.
* **Dependency Inversion Principle (Bağımlılığı Tersine Çevirme Prensibi):** ExcelBuilder sınıfı, ExcelPackage nesnesine bağımlıdır, ancak bağımlılık tersine çevirme prensibiyle bağımlılığı en aza indirgemeye çalışır.

### 2. Builder Tasarım Deseni
* ExcelBuilder sınıfı, bir Excel dosyası oluşturmak ve düzenlemek için bir builder benzeri yapı sunar. Builder tasarım deseni, karmaşık bir nesneyi adım adım oluşturmak ve farklı varyasyonlarına izin vermek için kullanılır.

### 3. Null Object Tasarım Deseni
* _worksheet değişkeni, null değer alabilir ancak null kontrolü yerine null conditional operatörü (?.) kullanılarak null olup olmadığı kontrol edilir. Bu, Null Object tasarım desenine bir benzerlik gösterir ve null değerleriyle daha güvenli bir şekilde çalışmayı sağlar.

## Kullanım
```cs
// Yeni bir ExcelBuilder örneği oluşturma
var builder = new ExcelBuilder("Sheet1");

// Başlık ve veri ekleme
builder.SetData("A1", "My Title")
    .SetData("A2", "Data 1")
    .SetData("B2", "Data 2");

// Verileri koleksiyondan aktarma
List<MyDataModel> data = GetDataFromSomewhere();
builder.SetDataList("A4", data);

// Excel dosyasını kaydetme
builder.Build("C:\\Documents", "MyExcelFile");

// Excel dosyasından veri almak
var importedData = ExcelBuilder.ImportToEntity<MyDataModel>("C:\\Documents\\MyExcelFile.xlsx");
```

Bu örnekler, ExcelBuilder sınıfının temel kullanımını göstermektedir. Sınıf, daha birçok işlevselliği destekler ve farklı senaryolara uyarlanabilir.

## Notlar
* ExcelBuilder sınıfı, EPPlus kütüphanesine dayanmaktadır. Bu nedenle, projenize EPPlus kütüphanesini eklemeniz gerekmektedir.
* Sınıfın içinde yer alan metodlar ve parametreler, gerekli uyarlama ve genişletmeler yapılarak ihtiyaçlara göre özelleştirilebilir.


Bu ExcelBuilder sınıfı, C# ile Excel işlemleri gerçekleştirmek isteyen geliştiriciler için bir temel oluşturmayı hedeflemektedir. İhtiyaçlarınıza göre bu sınıfı genişletebilir veya uyarlayabilirsiniz.
