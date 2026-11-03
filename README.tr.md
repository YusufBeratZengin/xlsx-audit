# xlsx-audit

**Excel'i açmadan `.xlsx` denetleyen bir linter: kırık formül referansları, döngüsel
başvurular, formül içine gömülmüş sabit sayılar, gizli sayfalar ve dış bağlantılar —
hepsi hücre adresiyle birlikte.**

`xlsx-audit` dosyanın içindeki OOXML paketini (aslında bir zip arşivi olan `.xlsx`)
doğrudan okur ve formülleri hiçbir zaman hesaplamaz. Bu sayede Linux CI sunucusunda,
pre-commit kancasında veya konteynerde çalışır; Excel, COM otomasyonu ya da LibreOffice
kurulu olması gerekmez.

```console
$ xlsx-audit check examples/fake_financial_model.xlsx --lang tr
✖ [FAIL] fake_financial_model.xlsx denetimi başarısız (1 Hata, 4 Uyarı)

- [HATA] Gelirler!E15: Kırık formül referansı bulundu (#REF!). (BROKEN_REFERENCE)
- [UYARI] <workbook>: Dış çalışma kitabı bağlantısı: …/Budget_2025.xlsx (EXTERNAL_LINK)
- [UYARI] Giderler!B8: Formül içine gömülü sabit sayı: 1.2. (MAGIC_NUMBER_IN_FORMULA)
- [UYARI] Giderler!D11: Giderler!C8:C9 toplam aralığı dolu hücreyi kapsamıyor: C10 (25). (OMITTED_CELL_WARNING)
- [UYARI] Notlar: Sayfa veryHidden olarak gizlenmiş: Excel arayüzünde görünmez. (VERY_HIDDEN_SHEET)

5 bulgu · 3 sayfa · 15 hücre · 4 formül
```

Varsayılan dil İngilizcedir (`--lang en`); yukarıdaki `--lang tr` çıktısının aynısını
İngilizce almak için seçeneği kaldırmanız yeterli. *(Bu çıktı depodaki
`examples/fake_financial_model.xlsx` dosyasından alınmıştır; içindeki her değer
uydurmadır.)*

## Neden var

E-tablolar kimsenin gözden geçirmediği iş kurallarıdır. Silinen bir satır geride `#REF!`
bırakır, bir oran doğrudan formülün içine yazılır (`=A8*1.20`), `veryHidden` bir sayfa
dosya denetime gönderilirken içinde taşınır. Avrupa E-tablo Riskleri Grubu bu tür
olayları yıllardır topluyor ([eusprig.org](http://www.eusprig.org)); aynı soruna eğilen
araştırma prototipi [ExceLint](https://github.com/ExceLint/ExceLint) ise bir Excel
eklentisi olduğu için sunucuda çalışamaz. `xlsx-audit` bunun başsız (headless) karşılığı:
tek komut, deterministik çıktı, makine tarafından okunabilir biçimler.

## Kurulum

```bash
pip install xlsx-audit
```

Python 3.10 ve üzeri. Tek çalışma zamanı bağımlılığı `defusedxml`: dosya güvenilmeyen bir
kaynaktan geldiğinde XML ayrıştırmasını varlık genişletme ve dış varlık saldırılarına
karşı sertleştirir. Geri kalan her şey standart kütüphane.

Geliştirme için:

```bash
git clone https://github.com/YusufBeratZengin/xlsx-audit
cd xlsx-audit
pip install -e ".[dev]"
bash scripts/check.sh
```

## 30 saniyelik örnek

```bash
xlsx-audit check butce.xlsx                 # insan için rapor, hata varsa çıkış kodu 1
xlsx-audit check butce.xlsx --strict        # uyarılar da başarısız sayılsın (yayın kapısı)
xlsx-audit check butce.xlsx --format json   # makine okunur
xlsx-audit check butce.xlsx --format sarif -o denetim.sarif
xlsx-audit rules                            # kuralları listele
```

Elinde dosya olmadan denemek için:

```bash
python3 scripts/make_example.py
xlsx-audit check examples/fake_financial_model.xlsx
```

## Kurallar

| Kod | Önem | Ne demek |
| --- | --- | --- |
| `BROKEN_REFERENCE` | hata | `#REF!` — genellikle satır/sütun silinince ölen başvuru |
| `FORMULA_ERROR_VALUE` | hata | hücrede veya formülde `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NUM!`, `#N/A`… |
| `CIRCULAR_REFERENCE` | hata | formül kendine ya da bir zincir üzerinden kendine bağlı (tüm döngüler, tam yolla) |
| `EXTERNAL_URL` | hata | kayıtlı URL köprüsü/bağlantısı, formülde URL metni veya `WEBSERVICE`/`RTD`/`STOCKHISTORY`/`FILTERXML` |
| `EXTERNAL_LINK` | uyarı | başka çalışma kitabına bağ veya köprü (`[1]Sheet1!A1`, `[C:\…\Budget.xlsx]…`) |
| `MAGIC_NUMBER_IN_FORMULA` | uyarı | formülün içine gömülmüş iş sabiti (`=B2*0.18`) |
| `OMITTED_CELL_WARNING` | uyarı | `=SUM(A1:A9)` yazılmışken `A10` dolu — bir satırı atlayan toplam |
| `VERY_HIDDEN_SHEET` | uyarı | Excel arayüzünden görünür hâle getirilemeyen sayfa |
| `HIDDEN_SHEET` | bilgi | gizli sayfa |
| `HIDDEN_DEFINED_NAME` | bilgi | gizli adlandırılmış aralık |
| `METADATA_PRESENT` | bilgi | yazar/şirket/başlık üst verisi; değerler asla yazdırılmaz |
| `FORMULA_PARSE_ERROR` | uyarı | statik ayrıştırıcı bir yapıyı çözemedi; diğer kurallar o hücreyi atlar |
| `GRAPH_LIMIT_EXCEEDED` | bilgi | kitap döngü analizi için çok büyük; kalan kurallar yine çalışır |

Yapısal sabitler (`0`, `1`, `2`, `100`) ve indeks/hassasiyet argümanları
(`ROUND(x, 2)`, `VLOOKUP(x, r, 3)`…) sihirli sayı sayılmaz. Bir sabit hücre başına bir
kez raporlanır, her geçtiği yerde değil.

Bulgular önem ve hücreye göre sıralanır; aynı dosyada iki çalıştırma her zaman aynı sırayı
verir. Dış hedefler maskelenir: yerel dizin ve sorgu parametreleri terminale, JSON'a veya
SARIF'e yazılmadan önce atılır.

## Çıkış kodları

| Kod | Anlamı |
| --- | --- |
| `0` | bulgu yok (ya da `--strict` yokken yalnızca uyarı var) |
| `1` | çalıştırmayı düşüren bulgular var |
| `2` | kullanım hatası, okunamayan dosya veya rapor yazma hatası |

Yalnızca rapor üreten hatlarda `--exit-zero`, bir kuralı susturmak için
`--ignore CODE[,CODE]`, sadece belirli kuralları çalıştırmak için `--rules ad[,ad]`,
renkleri kapatmak için `--no-color` (veya `NO_COLOR`) kullanın. `-v` bilgi düzeyindeki
bulguları da gösterir.

## Python API

```python
from xlsx_audit import audit_workbook

rapor = audit_workbook("veri.xlsx", strict=True)
if not rapor.passed:
    for bulgu in rapor.diagnostics:
        print(bulgu.location, bulgu.code, bulgu.message)
    raise SystemExit(1)
```

`rapor.to_dict()` JSON biçimlendiricisinin ürettiği yapının aynısıdır. Belirli kurallar
`audit_workbook(yol, rules=["broken_refs"])` ile seçilebilir; `audit_workbook_object`
okunmuş bir `Workbook` nesnesi alır (testler ve araçlar için).

## CI ve pre-commit

GitHub Actions ve Security sekmesine SARIF yükleme:

```yaml
- uses: actions/checkout@v4
- uses: actions/setup-python@v5
  with:
    python-version: "3.12"
- run: pip install xlsx-audit
- run: xlsx-audit check modeller/*.xlsx --format sarif -o xlsx-audit.sarif
  continue-on-error: true
- uses: github/codeql-action/upload-sarif@v3
  with:
    sarif_file: xlsx-audit.sarif
```

pre-commit — bu depoyu kanca olarak ekleyin (`xlsx-audit` yalnızca sahnelenen
`.xlsx`/`.xlsm` dosyalarında çalışır):

```yaml
repos:
  - repo: https://github.com/YusufBeratZengin/xlsx-audit
    rev: v0.1.0
    hooks:
      - id: xlsx-audit
```

## Nasıl çalışır

```
xlsx_audit/
├── ooxml.py           # sınırlı zip okuyucu: workbook.xml, rels, sharedStrings, sayfalar
├── formula_parser.py  # tokenizer + AST (başvurular, aralıklar, çağrılar, diziler)
├── refs.py            # A1 gösterimi yardımcıları
├── graph_audit.py     # bağımlılık çizgesi, yinelemeli Tarjan SCC → döngüler
├── rules/             # her kural için bir modül, ortak kayıt
├── reporter.py        # metin / JSON / Markdown
├── sarif_reporter.py  # SARIF 2.1.0
└── cli.py             # argparse giriş noktası
```

Paylaşılan formüller (`<f t="shared" si="0" ref="A1:A9"/>`), Excel'in yaptığı göreli
başvuru kaydırmasıyla her izleyici hücreye uygulanır; böylece kural, o hücrede gerçekten
çalışan formülü görür. Sayfa sınırının dışına kayan başvurular aynen Excel'deki gibi
`#REF!` olur. Dizi (spill) formülleri yazıldığı gibi okunur. Formül metni hiçbir zaman
hesaplanmaz; bu yüzden sonuçları önbelleğe almayan bir kütüphaneyle kaydedilmiş dosya da
denetlenebilir.

## Sınırlar

- **Yalnızca statik.** `xlsx-audit` değer hesaplamaz; bir formülün *doğru* olduğunu
  kanıtlayamaz, yalnızca bilinen risk kalıplarını gösterir. Bir denetim firmasının
  incelemesinin ya da finansal modelin dört göz kontrolünün yerini tutmaz.
- Tüm sütun başvuruları (`A:A`) ve 20 000 hücreden büyük aralıklar döngü analizi için
  açılmaz; yalnızca böyle bir aralığın arkasına saklanmış döngü gözden kaçar.
- Tüm satır başvuruları (`1:1`), parantezli birleşimler (`SUM((A1,B2))`) ve taşma
  başvuruları (`A1#`) ayrıştırıcının dilbilgisine dahil değildir; bunları kullanan hücre
  `FORMULA_PARSE_ERROR` olarak raporlanır ve diğer kurallar o hücreyi atlar.
- 64 seviyeden derin iç içe formüller (Excel'in kendi sınırı) ayrıştırılmaz; yığını
  tüketmesine izin verilmez, `FORMULA_PARSE_ERROR` olarak raporlanır.
- `IFERROR`/`IF` kısa devreleri ile `INDIRECT`/`OFFSET` dolaylılığı izlenmez.
- VBA makroları incelenmez ve çalıştırılmaz (`.xlsm` de statik XML olarak okunur).
- `veryHidden` sayfanın içeriği hücre hücre raporlanmaz; kişisel veri olabileceği için
  yalnızca sayfa düzeyinde bulgu üretilir.
- Bu projede resmî vergi, KDV veya yasal oran yoktur. Testlerde bir oran gerekiyorsa
  uydurma bir değerdir; aksi durum `NEEDS_SOURCE` olarak işaretlenir.

## Lisans

MIT — bkz. [LICENSE](LICENSE). Copyright (c) 2026 YusufBeratZengin.
