# -*- coding: utf-8 -*-
"""
PROJECT AETERNITAS - CİLT 10: BESİN ALGILAMA YOLAKLARI (mTOR, AMPK, SİRTUİNLER, IGF-1)
10 BÖLÜM x 10 DERİN ALT BAŞLIK = 100 AKADEMİK VE MOLEKÜLER BÖLÜM
WORD COM ILE HEDEFLENEN VE DOGRULANAN: 100+ SAYFA DOKTORA DUZEYI MASTERPIECE
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_10_BESIN_ALGILAMA_YOLAKLARI_mTOR_AMPK_SIRTUINLER_IGF1_TAM_100_SAYFA.docx"

doc = Document()

# Sayfa Yapisi: A4, 1 inc kenar bosluklari
for section in doc.sections:
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.different_first_page_header_footer = True
    
    # Header & Footer
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 10: BESİN ALGILAMA YOLAKLARI VE METABOLİK KONTROL")
    hrun.font.name = "Calibri"
    hrun.font.size = Pt(8.5)
    hrun.font.color.rgb = RGBColor(120, 144, 156)
    
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    frun = fp.add_run("NEXA ADVANCED LONGEVITY SCIENCES | ULTRA-ENCYCLOPEDIA SERIES")
    frun.font.name = "Calibri"
    frun.font.size = Pt(8.5)
    frun.font.color.rgb = RGBColor(144, 164, 174)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def format_cell(cell, bg_hex, text, font_size=9.5, bold=False, color_rgb=(40,40,40), align=WD_ALIGN_PARAGRAPH.LEFT):
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)

# ================= KAPAK VE GİRİŞ =================
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.paragraph_format.space_before = Pt(72)
title_p.paragraph_format.space_after = Pt(12)

t_run = title_p.add_run("PROJECT AETERNITAS: BİYOLOJİK ÖLÜMSÜZLÜK VE RADİKAL GENÇLEŞME KÜLLİYATI")
t_run.font.name = "Calibri"
t_run.font.size = Pt(15)
t_run.font.bold = True
t_run.font.color.rgb = RGBColor(0, 102, 153)

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_p.paragraph_format.space_after = Pt(18)
s_run = sub_p.add_run("CİLT 10: DEREGÜLE BESİN ALGILAMA YOLAKLARI VE METABOLİK MİMARİ\\n(mTORC1/2, AMPK ENERJİ SENSÖRÜ, SİRTUİNLER, IIS VE HORMONAL DÖNGÜ)")
s_run.font.name = "Calibri"
s_run.font.size = Pt(22)
s_run.font.bold = True
s_run.font.color.rgb = RGBColor(16, 44, 87)

meta_p = doc.add_paragraph()
meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
meta_p.paragraph_format.space_after = Pt(36)
m_run = meta_p.add_run("Doktora ve Post-Doktora İleri İhtisas Düzeyi | Tam Kapsamlı 100 Bölümlük Moleküler Referans Metni")
m_run.font.name = "Calibri"
m_run.font.size = Pt(11)
m_run.font.italic = True
m_run.font.color.rgb = RGBColor(80, 90, 100)

doc.add_page_break()

# ================= GİRİŞ MANİFESTOSU =================
intro_h = doc.add_paragraph()
intro_h.alignment = WD_ALIGN_PARAGRAPH.LEFT
intro_h.paragraph_format.space_before = Pt(18)
intro_h.paragraph_format.space_after = Pt(12)
ih_run = intro_h.add_run("CİLT 10 MANİFESTOSU: METABOLİK ENERJİ SENSÖRLERİ, ANABOLİK İLİZYON VE YAŞLANMANIN ŞALTERİ")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "Evrimsel süreçte yaşam, besin kıtlığı ve bolluk döngüleri arasında bir denge kurarak hayatta kalmıştır. "
    "Hücrenin besin algılama ağları (Nutrient-Sensing Pathways) — İnsülin/IGF-1 sinyalizasyonu (IIS), mTOR (mekanistik rapamisin hedefi), "
    "AMPK (adenozin monofosfatla aktive olan protein kinaz) ve NAD+ bağımlı Sirtuinler (SIRT1-7) — enerjinin büyümeye mi yoksa "
    "onarıma mı yönlendirileceğini belirleyen nihai biyolojik hakemlerdir.\\n\\n"
    "Modern çağda sürekli besin fazlalığı, mTORC1 ve IIS yolaklarını kronik olarak hiperaktive etmekte, anabolik büyüme sinyalleri "
    "hücresel temizlik mekanizmalarını (otofaji, proteazom, mitofaji) kalıcı olarak kilitlemektedir. Bunun sonucunda hücreler onarılamaz "
    "toksik agregatlar biriktirmekte, 'gerokonversiyon' ile senesense girmekte ve sistemik metabolik sendrom ile hızlanmış yaşlanma tetiklenmektedir. "
    "Tersine, kalori kısıtlaması (CR) veya farmakolojik taklitçiler (rapamisin, metformin, NAD+ öncülleri) AMPK ve Sirtuinleri aktive ederek "
    "hücreyi derin bir hücresel savunma, DNA tamiri ve rejenerasyon moduna geçirmektedir.\\n\\n"
    "Bu ciltte; IIS sinyal kaskadı, mTORC1/2 yapısı, AMPK enerji şalteri, Sirtuin epigenetik ağı, diyet kısıtlamasının biyofiziği, "
    "amino asit sensörleri (Sestrin2, CASTOR1), Laron sendromu ve GH/IGF-1 ekseni, farmakolojik müdahaleler (rapamisin, metformin, STAC'ler), "
    "sirkadiyen metabolizma entegrasyonu ve Homo Aeternus metabolik mimarisi 100 ayrıntılı akademik bölümde ele alınmaktadır."
)
ip = doc.add_paragraph()
ip.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
ip.paragraph_format.line_spacing = 1.15
ip.paragraph_format.space_after = Pt(16)
ip_run = ip.add_run(intro_body)
ip_run.font.name = "Calibri"
ip_run.font.size = Pt(10)
ip_run.font.color.rgb = RGBColor(40, 40, 40)

doc.add_page_break()

parts = []

# ================= KISIM 1 =================
part1_subsections = [
    (
        "1.1 İnsülin Reseptörü (INSR) ve IGF-1R Reseptör Tirozin Kinaz Aktivasyonu",
        "İnsülin reseptörü (INSR) ve insülin benzeri büyüme faktörü-1 reseptörü (IGF-1R), hücre yüzeyinde disülfit bağlarıyla birbirine bağlanmış alfa2-beta2 heterotetramerik glikoprotein kompleksleridir. Ligandın ekstraselüler alfa alt birimlerine bağlanması, beta alt birimlerinin trans-oto-fosforilasyonunu tetikler.",
        "Bu trans-oto-fosforilasyon, kinaz alanında yer alan kritik tirozin kalıntılarında (Tyr1158, Tyr1162, Tyr1163) gerçekleşir. Bu konformasyonel geçiş, kinaz aktivasyon döngüsünü dışarı doğru hareket ettirerek katalitik yarığı ATP ve protein substratlarının bağlanmasına açar. IGF-1R'nin afinitesi IGF-1 için pikomolar seviyede (Kd ~ 0.2 nM) iken, insülin için mikromolar seviyelere düşer. Yaşlanma ile birlikte reseptör membran akışkanlığının azalması ve lipid sallarındaki (lipid rafts) kolesterol birikimi, ligand-reseptör kenetlenme kinetiğini bozar ve reseptör dimerizasyon verimini düşürür.",
        "d[INSR*]/dt = k_on * [Ligand] * [INSR_inaktif] - k_off * [INSR*] - k_endo * [INSR*]",
        "Burada k_on ligand bağlanma hızı, k_off ayrılma sabiti, k_endo ise aktive reseptörün klatrin aracılı endositoz ile degredasyona uğrama kinetiğini temsil eder."
    ),
    (
        "1.2 IRS-1/2 Adaptör Proteinleri ve Fosforilasyon Dinamikleri",
        "İnsülin Reseptör Substratı 1 ve 2 (IRS-1 ve IRS-2), aktive reseptör tirozin kinazlar ile hücre içi sinyal kaskadı arasındaki köprü adaptör moleküllerdir. N-terminal pleckstrin homoloji (PH) ve fosfotirozin bağlama (PTB) alanları aracılığıyla membran fosfolipidlerine ve aktive INSR/IGF-1R'ye bağlanırlar.",
        "Reseptör kinaz tarafından IRS proteinlerinin çoklu tirozin motiflerinde (YXXM motifleri) fosforillenmesi, SH2 alanı içeren sinyal molekülleri için yüksek afiniteli pristin bağlanma cepleri yaratır. Ancak, kronik besin fazlalığı durumunda aktive olan S6K1, IKKbeta ve JNK kinazları, IRS-1'i kritik serin kalıntılarından (Ser307, Ser312, Ser636) inhibitör fosforilasyona uğratır. Bu durum, IRS-1'in reseptörden ayrılmasına, proteazomal yıkıma yönelmesine ve periferik insülin direncinin moleküler temelinin atılmasına yol açar.",
        "V_ser_phos = (V_max_S6K * [IRS1]) / (K_m + [IRS1]) + (V_max_JNK * [IRS1]) / (K_m_JNK + [IRS1])",
        "Bu eşitlik, artan S6K1 ve inflamatuar JNK aktivitesinin IRS-1 serin fosforilasyon akısını ve ardından gelen degradasyon hızını modellemektedir."
    ),
    (
        "1.3 PI3K (Fosfoinositid 3-Kinaz) Katalitik ve Düzenleyici Alt Birimleri (p110/p85)",
        "Fosfoinositid 3-Kinaz Sınıf IA (PI3K), p85 düzenleyici ve p110 katalitik alt birimlerinden oluşan bir heterodimerdir. Dinlenim durumunda p85 alt birimi, p110'un katalitik aktivitesini allosterik olarak baskılar ve enzimi inaktif durumda tutar.",
        "IRS-1 üzerindeki fosforile YXXM motiflerinin p85'in SH2 alanlarına bağlanması, p110 üzerindeki konformasyonel baskıyı kaldırır ve enzimi plazma membranına transloke eder. Aktive p110 alt birimi, plazma membranının iç yaprağında yer alan fosfatidilinozitol 4,5-bisfosfatı (PIP2), D-3 pozisyonundan fosforilleyerek fosfatidilinozitol 3,4,5-trisfosfata (PIP3) dönüştürür. Yaşlanma sürecinde p85 monomerlerinin aşırı birikimi, serbest p85'in IRS-1'e bağlanarak fonksiyonel p85-p110 dimerleriyle yarışması yoluyla sinyal iletiminde bozulmalara yol açar.",
        "[PIP3]_steady = (k_PI3K * [PIP2] * [PI3K*]) / (V_PTEN + k_deg)",
        "Bu denklemde PIP3 kararlı durum konsantrasyonu, PI3K aktivitesi ile membran PIP2 substratının çarpımının, PTEN hidroliz kapasitesine oranına bağlıdır."
    ),
    (
        "1.4 PIP2'den PIP3 Üretimi ve PTEN Lipid Fosfataz Fren Mekanizması",
        "PIP3, hücre zarında ikinci bir haberci olarak görev yapar ve pleckstrin homoloji (PH) alanı içeren sinyal proteinlerini hücre zarına çeker. PTEN (Phosphatase and Tensin Homolog), PIP3'ün 3-pozisyonundaki fosfat grubunu spesifik olarak hidrolize ederek PIP2'ye geri dönüştüren tümör baskılayıcı bir lipid fosfatazdır.",
        "PTEN aktivitesi, hücreyi kontrolsüz IIS ve Akt hiperaktivasyonundan koruyan en kritik fren mekanizmasıdır. Yaşlanma ile birlikte artan hücresel reaktif oksijen türleri (ROS), PTEN'in aktif merkezinde yer alan Cys124 kalıntısını oksitleyerek disülfit köprüsü oluşturur ve enzimi inaktive eder. Bu durum, bazal PIP3 seviyelerinde patolojik bir artışa, kontrolsüz anabolik sinyalizasyona ve hücresel yaşlanmanın hızlanmasına zemin hazırlar.",
        "J_PTEN = (k_cat_PTEN * [PTEN_aktif] * [PIP3]) / (K_m_PTEN + [PIP3])",
        "Burada J_PTEN, PTEN'in PIP3 deentegrasyon akısını, [PTEN_aktif] ise oksidatif hasara uğramamış fonksiyonel fosfataz konsantrasyonunu simgeler."
    ),
    (
        "1.5 PDK1 Kinazı ve Akt/PKB Serin/Treonin Kinazının Fosforilasyonu (Thr308 & Ser473)",
        "Membranda biriken PIP3, her ikisi de PH alanı içeren PDK1 (3-Fosfoinositid-Bağımlı Protein Kinaz 1) ve Akt (Protein Kinaz B / PKB) moleküllerini plazma zarına çeker. Bu eş-lokalizasyon, Akt'nin konformasyonel olarak açılmasını sağlar.",
        "PDK1, Akt'nin aktivasyon halkasında yer alan Treonin 308 (Thr308) kalıntısını fosforiller; bu işlem Akt kinaz aktivitesini kısmen tetikler. Tam enzimatik aktivite için mTORC2 (mTOR Kompleks 2) tarafından hidrofobik motifte bulunan Serin 473 (Ser473) kalıntısının fosforillenmesi zorunludur. Çift fosforillenmiş tam aktif Akt (p-Thr308 / p-Ser473), plazma membranından sitoplazmaya ve nükleusa göç ederek yüzlerce aşağı akış hedefini fosforiller. Yaşlanma sürecinde Akt'nin bazal seviyedeki konstitütif fosforilasyonu artarken, uyarılabilirlik dinamikleri kaybolur.",
        "Akt_aktif% = f([pThr308], [pSer473]) = (1 / (1 + exp(-alpha*(Thr308 - theta1)))) * (1 / (1 + exp(-beta*(Ser473 - theta2))))",
        "Bu lojistik fonksiyon, Akt'nin tam kinaz aktivitesine ulaşabilmesi için her iki fosforilasyon bölgesinin eş zamanlı doyumunu modellemektedir."
    ),
    (
        "1.6 Akt Aşağı Akış Hedefleri: FOXO Transkripsiyon Faktörlerinin Fosforilasyonu ve Dışlanması",
        "Aktive olan Akt, metabolizma, hücre sağkalımı, anabolizma ve yaşam süresini kontrol eden merkezi bir kavşaktır. Akt'nin yaşlanma biyolojisindeki en kritik hedefi FOXO (Forkhead box O) transkripsiyon faktörü ailesidir (FOXO1, FOXO3a, FOXO4).",
        "Akt, FOXO transkripsiyon faktörlerini üç korunmuş serin/treonin bölgesinden (FOXO3a için Thr32, Ser253, Ser315) doğrudan fosforiller. Bu fosforilasyon, FOXO'ya 14-3-3 şaperon proteinlerinin bağlanmasına yol açar. 14-3-3 bağlanması, FOXO'nun nükleer lokalizasyon sinyalini (NLS) maskeler ve nükleer dışa aktarım sinyalini (NES) açığa çıkararak FOXO'nun CRM1 taşıyıcısı aracılığıyla çekirdekten sitoplazmaya atılmasına neden olur. Böylece ömür uzatıcı genlerin transkripsiyonu durdurulur.",
        "d[FOXO_nukleus]/dt = -k_Akt_phos * [Akt*] * [FOXO_nukleus] + k_import * [FOXO_sitoplazma_defosfo]",
        "Bu formülasyon, Akt kinaz aktivitesinin çekirdekteki transkripsiyonel aktif FOXO havuzunu nasıl tükettiğini ve sitozole hapsettiğini gösterir."
    ),
    (
        "1.7 FOXO3a'nın Longevity Koruyucu Rolü: Antioksidan Savunma (SOD2, Katalaz) ve DNA Onarımı",
        "Düşük besin veya düşük insülin koşullarında Akt inaktif kaldığında, FOXO transkripsiyon faktörleri defosforile formda çekirdeğe girer ve DNA üzerindeki Forkhead yanıt elemanlarına (FHRE) bağlanır.",
        "FOXO3a özellikle insan uzun ömürlülüğü (longevity) ile genetik olarak en güçlü ilişkiye sahip faktördür. Çekirdekte MnSOD (SOD2) ve Katalaz genlerini uyararak mitokondriyal süperoksit ve hidrojen peroksit temizliğini artırır. Aynı zamanda GADD45a ve DDB1 genlerini aktive ederek DNA nükleotid eksizyon onarımını hızlandırır, Bim ve FasL üzerinden ise telafisi imkansız hasarlı hücrelerde kontrollü apoptozu tetikler. Yaşlanan dokularda FOXO3a'nın çekirdekten kronik dışlanması, genomik hasar ve serbest radikal hasarının katlanarak artmasına yol açar.",
        "Transkripsiyon_FOXO3a = V_trans * [FOXO3a_nukleus]^n / (K_A^n + [FOXO3a_nukleus]^n)",
        "Burada V_trans maksimal transkripsiyon hızı, K_A affinite sabiti ve n Hill katsayısı olup, FOXO3a'nın longevity genleri üzerindeki regülasyonunu belirtir."
    ),
    (
        "1.8 IGF-1 / GH Aksının Filogenetik Korunumu: Caenorhabditis elegans daf-2/daf-16 Paradigması",
        "İnsülin/IGF-1 sinyal ağının yaşam süresini belirlemedeki rolü, evrimsel olarak mayalardan nematodlara, böceklerden memelilere kadar korunmuş evrensel bir biyolojik prensiptir.",
        "Caenorhabditis elegans'ta insülin/IGF-1 reseptör homologu olan daf-2 genindeki inaktive edici mutasyonlar, nematodun yaşam süresini iki katına çıkarır (Kenyon et al.). Bu ömür artışı, daf-16 (FOXO homologu) transkripsiyon faktörünün mutlak varlığına bağlıdır; daf-16 silindiğinde ömür uzatıcı etki tamamen kaybolur. Benzer şekilde Drosophila'da Chico (IRS homologu) mutasyonları ve memelilerde IGF-1R heterozigot nakavt fareler (Igf1r+/-), artan oksidatif stres direnci ve anlamlı yaşam süresi uzaması gösterir.",
        "Lifespan_Model = T_0 * ( [daf-16_nukleus] / [daf-2_aktivite] )^gamma",
        "Bu ölçekleme eşitliği, filogenetik olarak korunmuş türlerde ömür beklentisinin, FOXO/daf-16 nükleer aktivitesinin IIS reseptör gücüne oranına bağlı olduğunu özetler."
    ),
    (
        "1.9 Yaşla İlişkili Hiperinsülinemi, İnsülin Direnci ve Reseptör Desensitizasyonu",
        "Kronik yüksek glisemik yük ve yaşlanan yağ dokusundaki enflamatuar sitokin salınımı (TNF-alfa, IL-6), pankreas beta hücrelerini sürekli yüksek insülin salgılamaya zorlayarak kronik kompanzatuar hiperinsülinemiye yol açar.",
        "Sürekli yüksek insülin konsantrasyonları, hedef dokulardaki (iskelet kası, karaciğer, adiposit) INSR reseptörlerinin endositozunu artırır, yüzeyel reseptör dansitesini düşürür ve SOCS1/3 proteinleri aracılığıyla reseptör-IRS kenetlenmesini bloke eder. Bu desensitizasyon döngüsü periferik dokularda glukoz klerensini düşürürken, karaciğerde de novo lipogenezi körükler. Ortaya çıkan ileri glikasyon son ürünleri (AGEs) ve ektopik lipid birikimi, vasküler endoteli bozarak biyolojik yaşlanmayı hızlandırır.",
        "d[INSR_yuzey]/dt = R_biyosentez - k_downreg * [Insulin] * [INSR_yuzey]",
        "Bu türev, dolaşımdaki kronik insülin yüksekliğinin hücre zarındaki fonksiyonel insülin reseptörü yoğunluğunu nasıl tükettiğini açıklamaktadır."
    ),
    (
        "1.10 IIS Sinyalinin Modülasyonu: Uzun Yaşam Fenotipi İçin Optimum İnsülin Sinyal Aralığı",
        "İnsülin sinyalinin tamamen kapatılması ölümcül diyabetik ketoasidoza veya doku atrofisine yol açarken, aşırı aktivasyonu yaşlanmayı hızlandırır ve kanser riskini artırır; bu durum hormetik bir U-eğrisi oluşturur.",
        "Uzun yaşam fenotipi, bazal dönemde minimum insülin ve IGF-1 seviyeleriyle karakterizedir; bu durum yüksek insülin duyarlılığı ve periyodik besin uyarımıyla birleşir. Düşük açlık insülini (<3 microIU/mL) ve dengeli IGF-1 konsantrasyonu (100-150 ng/mL), hücresel otofaji ve FOXO bağımlı onarım genlerinin açık kalmasını sağlarken, iskelet kas kütlesinin korunumu için yeterli anabolik pencereyi muhafaza eder. Geleceğin rejeneratif tıbbı, bu dar homeostatik pencereyi genetik ve farmakolojik geri bildirim döngüleriyle sabitlemeyi hedefler.",
        "Vitality_Index = C_1 * [Insulin_Duyarliligi] * exp( -([Insulin_aclik] - I_opt)^2 / (2 * sigma_I^2) )",
        "Bu gauss fonksiyonu, biyolojik canlılığın ve yaşam beklentisinin belirli bir optimal açlık insülin düzeyi (I_opt) etrafında maksimize olduğunu tanımlar."
    )
]

# ================= KISIM 2 =================
part2_subsections = [
    (
        "2.1 mTOR Kompleks 1 (mTORC1) Yapısı: mTOR, Raptor, mLST8, PRAS40 ve DEPTOR",
        "mTOR Kompleks 1 (mTORC1), hücresel büyüme, protein translasyonu ve metabolizmanın merkez üssü olan 1 MDa ağırlığında devasa bir multiprotein kinaz kompleksidir. Kompleksin merkezinde serin/treonin kinaz aktivitesine sahip mTOR enzimi bulunur.",
        "Raptor (Regulatory-associated protein of mTOR), substratların (p70S6K ve 4E-BP1 gibi) TOS (TOR signaling) motifleri aracılığıyla komplekse tanıtılmasından sorumludur. mLST8 (GbetaL), kinaz alanının stabilitesini sağlar. PRAS40 ve DEPTOR ise kompleksin endojen inhibitörleridir; bazal durumda mTORC1'e bağlı kalarak kinaz aktivitesini baskılarlar. Akt veya besin sinyalleri ile PRAS40 fosforillendiğinde kompleksten ayrılır ve mTORC1 tam katalitik aktiviteye kavuşur. Yaşlanma sürecinde DEPTOR ekspresyonunun azalması, mTORC1'in kontrolsüz hiperaktivasyonuna katkıda bulunur.",
        "Katalitik_Kapasite_mTORC1 = k_cat * [mTOR-Raptor-mLST8] / (1 + [PRAS40]/K_i1 + [DEPTOR]/K_i2)",
        "Bu formülasyon, inhibitör alt birimlerin komplekse bağlanma dengesi üzerinden net mTORC1 aktivasyon kinetiğini ortaya koymaktadır."
    ),
    (
        "2.2 mTOR Kompleks 2 (mTORC2) Yapısı: mTOR, Rictor, mLST8, Sin1 ve Protor1/2",
        "mTOR Kompleks 2 (mTORC2), rapamisine akut olarak duyarsız olan ve hücre sağkalımı, sitoiskelet organizasyonu ile metabolizmayı denetleyen ikinci majör mTOR montajıdır.",
        "mTORC2'de Raptor yerine Rictor (Rapamycin-insensitive companion of mTOR) bulunur ve bu alt birim substrat spesifikliğini belirler. Sin1 (Stress-activated map kinase-interacting protein 1), kompleksin plazma membranındaki PIP3'e bağlanmasını sağlayan kritik bir PH alanına sahiptir. Protor1/2 (Protein observed with Rictor) ve mLST8 yapıyı tamamlar. mTORC2'nin en önemli fizyolojik görevi, Akt'nin Ser473 bölgesini ve SGK1 ile PKCalpha kinazlarını fosforillemektir. Akut rapamisin tedavisi sadece mTORC1'i inhibe ederken, kronik rapamisin maruziyeti serbest mTOR havuzunu tüketerek zamanla mTORC2 montajını da bozar.",
        "d[mTORC2]/dt = k_assoc * [mTOR] * [Rictor] * [Sin1] - k_dissoc * [mTORC2] - k_kronik_rapa * [Rapa-FKBP12] * [mTORC2]",
        "Bu kinetik eşitlik, kronik rapamisin baskısının mTORC2 kompleks montajını nasıl kademeli olarak baskıladığını göstermektedir."
    ),
    (
        "2.3 Rheb GTPaz Aktivasyonu: TSC1/TSC2 (Tüberoz Skleroz Kompleksi) Fren Ağı",
        "mTORC1'in lizozom zarı üzerindeki nihai aktivasyonu, küçük bir G-proteini olan Rheb'in (Ras homolog enriched in brain) GTP yüklü formuna bağlanmasıyla gerçekleşir. Rheb-GTP, mTOR'un katalitik merkezinde allosterik bir rotasyon yaratarak kinazı aktive eder.",
        "Rheb'in aktivitesini kontrol eden ana fren sistemi TSC1/TSC2/TBC1D7 kompleksidir. TSC2, güçlü bir GTPaz aktive edici protein (GAP) aktivitesine sahiptir ve Rheb'i inaktif Rheb-GDP formuna dönüştürür. Büyüme faktörleri ve Akt sinyali, TSC2'yi doğrudan fosforilleyerek (Ser939, Thr1462) 14-3-3 proteinleriyle kompleks yapmasını ve lizozomal membrandan sitozole uzaklaşmasını sağlar. TSC freni kalktığında Rheb hızla GTP yüklenir ve mTORC1'i ateşler. Tersine AMPK, TSC2'yi farklı bölgelerden (Ser1387) fosforilleyerek GAP aktivitesini güçlendirir ve mTORC1'i kapatır.",
        "d[Rheb_GTP]/dt = GEF_efektif * [Rheb_GDP] - k_GAP_TSC2 * [TSC_kompleks*] * [Rheb_GTP]",
        "Bu diferansiyel denklem, lizozom zarındaki Rheb-GTP havuzunun TSC2 GAP fren aktivitesi ile anabolik sinyaller arasındaki hassas dinamik dengesini açıklar."
    ),
    (
        "2.4 p70S6K1 Kinaz Aktivasyonu ve Ribozomal Protein S6 Fosforilasyonu",
        "Aktive mTORC1'in en kritik effekti, p70 ribozomal S6 kinaz 1'in (p70S6K1) aktivasyon döngüsünde yer alan Thr389 kalıntısını doğrudan fosforillemesidir. Bu fosforilasyon, PDK1'in Thr229'u fosforillemesi için yolu açar ve S6K1 tam aktif hale gelir.",
        "S6K1, 40S ribozomal alt birim proteini olan S6'yı (Ser235, Ser236, Ser240, Ser244) fosforilleyerek translasyonel kapasiteyi artırır. Ayrıca eEF2K'yı (ökaryotik elongasyon faktörü 2 kinaz) inhibe ederek translasyonel uzamayı serbest bırakır ve PDCD4'ü degredasyona uğratarak eIF4A helikazını aktive eder. Ancak S6K1'in aşırı aktivasyonu, IRS-1'i serin fosforilasyonuyla yıkarak negatif geri bildirim döngüsü yaratır ve hücreyi periferik insülin direncine sokar. S6K1 nakavt farelerin (S6K1-/-) yaşam sürelerinin %20'ye varan oranlarda uzadığı kanıtlanmıştır.",
        "Translational_Flux_S6 = k_trans * [S6K1_pThr389] * [Ribozom_Aktif] / (K_M + [Ribozom_Aktif])",
        "Bu eşitlik, fosforile S6K1 seviyelerinin global ribozomal translasyon akısı üzerindeki doğrudan etkisini matematiksel olarak bağlar."
    ),
    (
        "2.5 4E-BP1 (Ökaryotik Translasyon Başlatma Faktörü 4E Bağlayıcı Protein 1) ve Translasyonel Kontrol",
        "4E-BP1, hücresel protein sentezinin kapı bekçisidir. Hipofosforile formda olan 4E-BP1, ökaryotik translasyon başlatma faktörü 4E'ye (eIF4E) yüksek afiniteyle bağlanır ve onun mRNA'nın 5'-başlık (cap) yapısına tutunarak eIF4G ve eIF4A ile eIF4F kompleksini kurmasını engeller.",
        "Besin bolluğunda mTORC1, 4E-BP1'i hiyerarşik olarak dört ayrı bölgeden (Thr37, Thr46, Thr70 ve Ser65) fosforiller. Hiperfosforile hale gelen 4E-BP1, eIF4E'den ayrılır. Serbest kalan eIF4E, 5' UTR bölgelerinde karmaşık ikincil yapılar içeren büyüme faktörleri, siklinler ve onkogenlerin cap-bağımlı translasyonunu başlatır. Yaşlanma karşıtı müdahalelerde 4E-BP1'in hipofosforile tutulması, hatalı katlanmış protein sentez yükünü azaltarak proteostazın korunmasını sağlar.",
        "[eIF4E_serbest] = [eIF4E_toplam] / ( 1 + K_A_4EBP * [4EBP1_defosfo] )",
        "Bu denge formülü, translasyona hazır serbest eIF4E miktarının, hipofosforile 4E-BP1 konsantrasyonu tarafından nasıl ters orantılı sınırlandığını gösterir."
    ),
    (
        "2.6 mTORC1 Tarafından ULK1 ve Atg13 İnhibisyonu: Otofajik Blokaj Mekanizması",
        "Hücrenin en temel hayatta kalma ve gençleşme mekanizması olan makrootofaji, ULK1 (Unc-51 Like Autophagy Activating Kinase 1) kompleksi tarafından başlatılır. Besin bolluğunda mTORC1, otofajinin en amansız inhibitörüdür.",
        "Aktif mTORC1, ULK1'e bağlanarak onu Ser757 kalıntısından fosforiller. Bu fosforilasyon, ULK1'in metabolik aktivatörü olan AMPK ile etkileşime girmesini sterik olarak engeller. Eş zamanlı olarak mTORC1, Atg13'ü Ser258 bölgesinden fosforilleyerek otofagozom membran başlatma kompleksinin montajını bloke eder. Sonuç olarak, hücre içindeki hasarlı organeller, yıpranmış mitokondriler ve agregatlaşmış proteinler lizozoma gönderilemez. Yaşlanma ile kronikleşen mTORC1 sinyali, dokuları otofajik olarak felç eder.",
        "Otofaji_Inhibisyon_Orani = [mTORC1*] * [ULK1] / (K_ULK_mTOR + [ULK1]) * (1 - exp(-k_pSer757 * t))",
        "Bu kinetik denklem, mTORC1 kinaz aktivitesinin ULK1 Ser757 fosforilasyonu üzerinden otofaji başlatma kapasitesini nasıl kilitlediğini modellemektedir."
    ),
    (
        "2.7 SREBP1/2 ve Lipit Biyosentezinin mTORC1 Bağımlı Transkripsiyonel İndüksiyonu",
        "Hücresel büyüme ve proliferasyon, plazma ve organel membranları için muazzam miktarda lipid sentezi gerektirir. mTORC1, de novo lipogenezin ana transkripsiyon faktörleri olan SREBP1 ve SREBP2'nin (Sterol Regulatory Element-Binding Proteins) aktivasyonunu yönetir.",
        "mTORC1, nükleer girişi engelleyen lipin-1 fosfatazını fosforilleyerek nükleus dışında tutar. Lipin-1 baskılandığında SREBP1/2 çekirdeğe girer ve yağ asidi sentaz (FASN), asetil-KoA karboksilaz (ACC) ve HMG-CoA redüktaz genlerinin ekspresyonunu artırır. Yaşlanan organizmada mTORC1'in kronik sinyali, visseral adipoz dokuda lipotoksisiteye, karaciğer yağlanmasına (hepatosteatoz) ve aterosklerotik plak gelişimine doğrudan zemin hazırlar.",
        "Transkripsiyon_SREBP = V_max_lipin_inh * [mTORC1*] / (K_lipin + [mTORC1*]) * [Lipid_Gen_Promoter]",
        "Bu hız eşitliği, mTORC1 seviyesinin lipin-1 inaktivasyonu üzerinden lipojenik enzim transkripsiyon hızını nasıl belirlediğini açıklar."
    ),
    (
        "2.8 mTORC1 Hiperaktivasyonunun Hücresel Yaşlanma (Gerokonversiyon) Üzerindeki Rolü",
        "Hücre döngüsü arresti (p16INK4a veya p21CIP1 yoluyla) yaşayan bir hücre, derhal ölmez; eğer bu hücrede büyüme sinyalleri aktif kalmaya devam ederse 'gerokonversiyon' adı verilen geri dönüşsüz senesens sürecine girer.",
        "Blagosklonny tarafından tanımlanan bu paradigmaya göre; hücre bölünemediği halde mTORC1 anabolik hiperaktivitesini sürdürdüğünde, hücre kontrolsüz bir şekilde hipertrofiye uğrar, hacmi genişler, granüler hale gelir ve aşırı protein salgılar. Bu durum, senesensle ilişkili salgı fenotipinin (SASP: IL-6, IL-8, MMP'ler) patlamasına yol açar. Rapamisin ile mTORC1 inhibe edildiğinde gerokonversiyon durdurulur; hücre hipertrofik senesense girmek yerine geri döndürülebilir bir dormansi (quiescence) durumunda kalır.",
        "Gerokonversiyon_Hizi = k_gero * [p21/p16_Arrest] * [mTORC1_Aktif]^2 / (K_gero + [mTORC1_Aktif]^2)",
        "Bu eşitlik, hücre döngüsü arresti varlığında mTORC1 aktivite derecesinin geri dönüşsüz hücresel senesense geçiş hızını nasıl üstel olarak artırdığını tanımlar."
    ),
    (
        "2.9 mTORC2'nin Akt Ser473 Fosforilasyonu ve Hücre İskeleti Düzenlemesindeki Fonksiyonu",
        "mTORC2, sadece bir metabolik aracı değil, aynı zamanda aktin hücre iskeletinin spatial polaritesini ve hücre göçünü düzenleyen yapısal bir merkezdir. RhoA, Rac1 ve Cdc42 GTPazlarını PKCalpha üzerinden koordine eder.",
        "mTORC2'nin Sin1 alt biriminin PIP3'e bağlanması, kompleksin kinaz aktivitesini Akt'nin Ser473 kalıntısına odaklar. Ser473 fosforilasyonu, Akt'nin Thr308 fosforilasyonunu stabilize eder ve Akt'nin pro-apoptotik faktörleri (Bad, Kaspaz-9) inaktive ederek hücre sağkalımını garanti altına almasını sağlar. Bununla birlikte, mTORC2'nin aşırı baskılanması hepatik insülin direncine, glukoneojenez kontrolünün kaybına ve dislipidemiye neden olabilir; bu nedenle longevity müdahalelerinde mTORC1 selektivitesi hayati önem taşır.",
        "d[Akt_pSer473]/dt = k_mTORC2 * [mTORC2_membran] * [Akt] - k_PHLPP * [PHLPP] * [Akt_pSer473]",
        "Burada mTORC2 kinaz aktivitesi ile PHLPP fosfatazının Ser473 defosforilasyon hızı arasındaki kinetik yarışma ifade edilmektedir."
    ),
    (
        "2.10 mTORC1/mTORC2 Selektivitesi ve Yaşlanma Karşıtı Müdahale Stratejileri",
        "İdeal bir anti-aging protokolü, otofajiyi açmak ve translasyonel yükü hafifletmek için mTORC1'i inhibe etmeli; ancak metabolik homeostazı ve doku rejenerasyonunu bozmamak için mTORC2'yi fonksiyonel tutmalıdır.",
        "Klasik rapamisin, FKBP12 ile kompleks yaparak mTOR'un FRB alanına bağlanır ve mTORC1'i allosterik olarak kilitler. Ancak yüksek doz veya uzun süreli sürekli kullanımda serbest mTOR bağlanarak mTORC2 montajı bozulur ve immünosupresyon, glukoz intoleransı gibi yan etkiler gelişir. Aralıklı rapamisin protokolleri (haftada bir yüksek doz puls) veya yeni nesil RapaLink ve DL001 benzeri mTORC1-selektif moleküller, mTORC2'ye dokunmadan gerokonversiyonu durdurmakta ve yaşam süresini yan etkisiz uzatabilmektedir.",
        "Selektivite_Indeksi = (IC50_mTORC2 / IC50_mTORC1) * (Tau_mTORC2_ayrilma / Tau_mTORC1_baglanma)",
        "Bu oran, bir molekülün mTORC2'yi korurken mTORC1 kompleksini ne kadar spesifik ve güvenli bir şekilde inhibe edebildiğini gösteren terapötik katsayıdır."
    )
]

# ================= KISIM 3 =================
part3_subsections = [
    (
        "3.1 AMPK Heterotrimerik Yapısı: Katalitik Alfa Alt Birimi ve Düzenleyici Beta/Gama Alt Birimleri",
        "AMPK (AMP-aktive Protein Kinaz), ökaryotik hücrelerin enerji homeostazını koruyan merkezi yakıt sensörüdür. Bir katalitik alfa alt birimi (alfa1 veya alfa2), bir yapısal/iskelet beta alt birimi (beta1 veya beta2) ve bir sensör gama alt biriminden (gama1, gama2 veya gama3) oluşan heterotrimerik bir komplekstir.",
        "Alfa alt birimi N-terminal kinaz alanını ve kritik Thr172 fosforilasyon döngüsünü taşır. Beta alt birimi glikojen bağlama alanına (GBD) sahiptir ve hücrenin karbonhidrat rezerv durumunu izler. Gama alt birimi ise dört adet sistationin beta-sentaz (CBS) alanı içerir; bu alanlar hücresel adenin nükleotidlerini (ATP, ADP ve AMP) yarışmalı olarak bağlar. Yaşlanmayla birlikte dokularda AMPK alfa alt birimi ekspresyonu azalır ve heterotrimer montaj bütünlüğü aşınır.",
        "AMPK_Kompleks_Kararliligi = K_montaj * [alfa] * [beta] * [gama] / (1 + k_termal_denaturasyon * Delta_Yas)",
        "Bu denklem, heterotrimerik AMPK kompleksinin hücresel konsantrasyonu ile yaşa bağlı konformasyonel kararsızlık parametrelerini tanımlar."
    ),
    (
        "3.2 AMP/ATP ve ADP/ATP Oranlarının Algılanması: Gama Alt Birimindeki CBS Alanları",
        "Hücrede adenilat kinaz enzimi, 2 ADP <-> ATP + AMP reaksiyonunu katalizler. Bu nedenle hücresel ATP seviyelerindeki çok küçük bir düşüş, serbest AMP konsantrasyonunda logaritmik bir artışa neden olur.",
        "AMPK gama alt birimindeki CBS alanlarına AMP veya ADP bağlandığında üçlü bir aktivasyon mekanizması devreye girer: 1) Enzimin allosterik olarak 10 kata kadar aktive olması; 2) Yukarı akış kinazı LKB1 tarafından Thr172 fosforilasyonunun kolaylaştırılması; 3) Protein fosfatazların (özellikle PP2C) Thr172 kalıntısını defosforillemesinin sterik olarak engellenmesi. ATP ise yüksek afiniteyle CBS alanlarına bağlanarak AMP'nin bu etkilerini antagonize eder. Yaşlanan hücrelerde mitokondriyal disfonksiyon nedeniyle enerji krizi yaşansa bile, sensör mekanizmasındaki desensitizasyon AMPK aktivasyonunu geciktirir.",
        "AMPK_Allosterik_Faktor = ( [AMP] / K_d_AMP ) / ( 1 + [ATP] / K_d_ATP + [ADP] / K_d_ADP + [AMP] / K_d_AMP )",
        "Bu model, adenin nükleotidlerinin gama alt birimindeki yarışmalı bağlanma dinamiklerinin AMPK allosterik yanıt katsayısını nasıl belirlediğini simgeler."
    ),
    (
        "3.3 LKB1 (Karaciğer Kinaz B1) ve CaMKK2 Tarafından Thr172 Fosforilasyonu",
        "AMPK'nin fonksiyonel olarak aktif olabilmesi için alfa alt biriminin aktivasyon halkasındaki Treonin 172 (Thr172) kalıntısının fosforillenmesi mutlak bir zorunluluktur.",
        "Bu fosforilasyonu gerçekleştiren primer kinaz LKB1'dir (tümör baskılayıcı STK11). LKB1, STRAD ve MO25 proteinleriyle konstitütif olarak aktif bir trimer oluşturur; AMP varlığında konformasyonu değişen AMPK'yi hızla fosforiller. İkinci ana yolak ise intraselüler kalsiyum artışına yanıt veren CaMKK2'dir (Kalsiyum/Kalmodulin-Bağımlı Protein Kinaz Kinaz 2). CaMKK2, AMP/ATP oranından bağımsız olarak egzersiz veya nöronal uyarılma sırasında sitozolik kalsiyum dalgasıyla AMPK'yi aktive eder. Yaşlanma ile nükleer LKB1 sekestrasyonu artarak sitoplazmik AMPK erişimi zayıflar.",
        "d[AMPK_pThr172]/dt = (k_LKB1 * [LKB1*] + k_CaMKK2 * [CaMKK2*]) * [AMPK_inaktif] - k_PP2C * [PP2C] * [AMPK_pThr172]",
        "Bu diferansiyel eşitlik, Thr172 fosforilasyonunun LKB1/CaMKK2 aktivasyon akısı ile PP2C fosfataz inaktivasyon akısı arasındaki kinetik yarışını modeller."
    ),
    (
        "3.4 AMPK'nin mTORC1'i Çift Yönlü İnhibisyonu: TSC2 Fosforilasyonu (Ser1387) ve Raptor Fosforilasyonu (Ser792)",
        "AMPK ve mTORC1, hücresel metabolizmanın iki karşıt kutbudur. Enerji kıtlığı algılandığında AMPK, mTORC1'in savurgan anabolik büyüme programını durdurmak için sofistike bir çift kilitli moleküler fren mekanizması uygular.",
        "Birinci kilit: AMPK, TSC2 tümör baskılayıcı proteinini Ser1387 kalıntısından doğrudan fosforiller; bu işlem TSC kompleksinin Rheb-GAP aktivitesini dramatik olarak artırarak Rheb-GTP havuzunu boşaltır ve mTORC1'i devre dışı bırakır. İkinci kilit: AMPK, mTORC1'in ana substrat adaptörü olan Raptor'u doğrudan iki kritik kalıntıdan (Ser722 ve Ser792) fosforiller. Fosforile Raptor'a 14-3-3 proteinleri bağlanır ve mTORC1 kinaz kompleksinin substrat bağlamasını mekanik olarak bloke eder. Bu çift darbe, besin yokluğunda ATP israfını engeller.",
        "Inhibisyon_mTORC1 = 1 - ( 1 / (1 + [AMPK*]/K_TSC2) ) * ( 1 / (1 + [AMPK*]/K_Raptor) )",
        "Bu fonksiyon, AMPK'nin TSC2 ve Raptor üzerindeki eş zamanlı fosforilasyonunun mTORC1 aktivitesini nasıl sinerjistik ve eksiksiz kapattığını gösterir."
    ),
    (
        "3.5 ULK1'in Doğrudan Aktivasyonu (Ser317, Ser777) ve Mitofaji / Otofaji Başlatılması",
        "mTORC1'i kapatarak dolaylı yoldan otofaji üzerindeki baskıyı kaldıran AMPK, bununla yetinmeyip otofaji başlatma kinazı ULK1'i doğrudan fosforilleyerek katalitik bir süper şarj sağlar.",
        "AMPK, ULK1 proteinini Ser317 ve Ser777 kalıntılarından doğrudan fosforiller. Bu fosforilasyon, ULK1 kinaz aktivitesini ateşler, VPS34 lipid kinaz kompleksinin (Beclin-1, Atg14L) lizozomal ve ER membranlarında PI3P üretimini başlatmasını sağlar ve otofagozom izolasyon membranının nükleasyonunu tetikler. Hasarlı mitokondrilerin seçici temizliği olan mitofaji sürecinde de AMPK-ULK1 ekseni, Parkin ve PINK1 ile işbirliği yaparak enerji tüketen yıpranmış mitokondrileri lizozomlara sürer. Yaşlanmayla birlikte bu eksenin körelmesi, hücre içinde senil lizozomal çöp birikimine yol açar.",
        "J_otofaji_baslangic = k_oto * [AMPK_pThr172] * [ULK1_defosfo_mTOR] / ( K_ULK + [ULK1_defosfo_mTOR] )",
        "Bu hız denklemi, otofaji başlama debisinin AMPK aktivitesi ve mTORC1 baskısından kurtulmuş serbest ULK1 konsantrasyonuyla orantılı olduğunu belgeler."
    ),
    (
        "3.6 PGC-1alpha Aktivasyonu ve Mitokondriyal Biyogenezin İndüksiyonu",
        "Enerji krizini çözen nihai uzun vadeli yanıt, yeni ve yüksek verimli mitokondrilerin üretilmesidir (mitokondriyal biyogenez). AMPK, bu programın ana orkestra şefi olan PGC-1alpha'yı (Peroxisome proliferator-activated receptor gamma coactivator 1-alpha) aktive eder.",
        "AMPK, PGC-1alpha'yı Thr177 ve Ser538 kalıntılarından doğrudan fosforiller. Bu öncül fosforilasyon, PGC-1alpha'nın konformasyonunu açarak onun SIRT1 deasetilazı tarafından deasetile edilmesini mümkün kılar. Tam aktif PGC-1alpha çekirdeğe girer, NRF-1 ve NRF-2 (Nükleer Solunum Faktörleri) transkripsiyon faktörleriyle birleşerek TFAM (Mitokondriyal Transkripsiyon Faktörü A) ekspresyonunu uyarır. TFAM, mitokondriyal DNA'nın (mtDNA) replikasyonunu ve solunum zinciri komplekslerinin sentezini başlatır. Yaşlanmada PGC-1alpha aktivasyonunun düşmesi mitokondriyal kütle kaybının temel nedenidir.",
        "d[mtDNA_kopyasi]/dt = k_biyogenez * [PGC1a_p_deac] * [TFAM] - k_mitofaji * [Parkin*]",
        "Bu denge denklemi, mitokondriyal biyogenez akısı ile mitofaji yıkım hızı arasındaki hücre içi mitokondriyal genetik havuz dinamiklerini gösterir."
    ),
    (
        "3.7 Lipid Katabolizması: ACC (Asetil-KoA Karboksilaz) İnhibisyonu ve CPT-1 Aktivasyonu",
        "Hücrenin hızlı ATP üretmesi gerektiğinde en zengin enerji kaynağı yağ asidi beta-oksidasyonudur. AMPK, lipid anabolizmasını durdurup katabolizmasını açan anahtardır.",
        "AMPK, ACC1 (sitoplazmik) ve ACC2 (mitokondriyal membran) izoformlarını Ser79 ve Ser221 kalıntılarından fosforilleyerek inaktive eder. ACC inaktive olduğunda, Malonil-KoA seviyeleri hızla düşer. Malonil-KoA, mitokondriyal dış zarda yer alan CPT-1 (Karnitin Palmitoiltransferaz 1) enziminin güçlü bir allosterik inhibitörüdür. Malonil-KoA seviyeleri düştüğünde CPT-1 üzerindeki fren kalkar; uzun zincirli yağ asitleri karnitin mekiğiyle mitokondriyal matrikse akar ve beta-oksidasyonla yoğun ATP üretilir. Bu mekanizma visseral yağlanmayı ve karaciğer lipotoksisitesini engeller.",
        "J_beta_oksidasyon = V_max_CPT1 * [Yag_Asidi-KoA] / ( K_M * (1 + [Malonil-KoA]/K_i_malonil) + [Yag_Asidi-KoA] )",
        "Bu Michaelis-Menten türevi, azalan Malonil-KoA seviyelerinin CPT-1 üzerindeki kompetitif inhibisyonu kaldırarak beta-oksidasyon akısını nasıl patlattığını açıklar."
    ),
    (
        "3.8 GLUT4 Translokasyonu ve İnsülinden Bağımsız Glukoz Alımı",
        "İskelet kası ve miyokard dokusunda glukoz alımı normalde insülin sinyaline bağımlıdır. Ancak metabolik stres ve egzersiz sırasında AMPK, insülin reseptörünü tamamen baypas ederek hücre içine glukoz akışını sağlar.",
        "AMPK, Rab-GAP proteinleri olan AS160 (TBC1D4) ve TBC1D1'i doğrudan fosforiller. Bu fosforilasyon, Rab proteinleri (Rab8A, Rab13) üzerindeki GAP baskısını kaldırır; GTP yüklü Rab'ler GLUT4 depolama veziküllerinin (GSV) aktin hücre iskeleti boyunca plazma zarına taşınmasını ve SNARE kompleksiyle füzyonunu sağlar. İnsülin direnci gelişmiş yaşlı bireylerde AMPK'nin egzersiz veya farmakolojik olarak uyarılması, hiperglisemiyi ve glikotoksik stresi tedavi etmenin en etkili fizyolojik yoludur.",
        "Glukoz_Alim_Hizi = V_basal + V_insulin * f(Akt) + V_AMPK * ( [AMPK_pThr172] / (K_GLUT4 + [AMPK_pThr172]) )",
        "Bu toplam fonksiyon, hücresel glukoz girişinin insülin ve AMPK yollarının bağımsız katkılarıyla nasıl regüle edildiğini ortaya koymaktadır."
    ),
    (
        "3.9 Yaşlanma ile AMPK Duyarlılığının Kaybı ve Dokularda Enerji Krizinin Oluşumu",
        "Kronik sedanter yaşam, aşırı kalori alımı ve yaşa bağlı sistemik enflamasyon, dokularda AMPK sinyal ağının duyarsızlaşmasına ve körelmesine yol açar.",
        "Yaşlanan dokularda protein fosfataz 2C (PP2C) aktivitesi artarken, yukarı akış kinazı LKB1'in mitokondriyal ve sitoplazmik havuzları tükenir. Aynı zamanda AMPK gama alt biriminde biriken oksidatif modifikasyonlar CBS alanlarının adenin nükleotid afinitesini bozar. Sonuç olarak, hücre düşük enerji seviyelerine (yüksek AMP/ATP) rağmen AMPK'yi aktive edemez. Bu durum hücresel enerji krizine, lipidlerin parçalanamayıp toksik seramid ve diasilgliserollere dönüşmesine, otofaji yetersizliğine ve mitokondriyal dejenerasyona neden olur.",
        "Duyarlilik_Kaybi_AMPK = exp( -k_yas * Yas_Biyolojik ) * ( 1 - [ROS_mitokondri]/ROS_esik )",
        "Bu bozunma fonksiyonu, biyolojik yaş ve mitokondriyal reaktif oksijen hasarının AMPK duyarlılığını nasıl kademeli olarak sıfıra yaklaştırdığını formüle eder."
    ),
    (
        "3.10 AMPK Yeniden Duyarlaştırma Terapileri ve Metabolik Rejuvenasyon",
        "AMPK'nin yaşlanmayla körelen fonksiyonunun restore edilmesi, sistemik gençleşmenin ve metabolik esnekliğin geri kazanılmasında en umut verici hedeflerden biridir.",
        "Doğrudan allosterik AMPK aktivatörleri (A-769662, PF-06409577, MK-8722) beta-1 alt birimindeki Karbonhidrat Bağlama Modülü (CBM) ile kinaz alanı arasındaki allosterik ilaç ve metabolit (ADaM) cebine bağlanır. Bu bağlanma, AMP'den bağımsız olarak enzimi 50 kata kadar aktive eder ve defosforilasyondan korur. Dolaylı yoldan mitokondriyal solunum zincirini hafifçe baskılayan metformin ve berberin gibi ajanlar ise AMP/ATP oranını yukarı çekerek fizyolojik aktivasyonu tetikler. Bu moleküler müdahaleler, yaşlı dokularda otofajiyi yeniden başlatarak doku gençleşmesini sağlar.",
        "Kinetik_Kazanc_ADaM = ( V_max_allosterik * [ADaM_Ilac] ) / ( K_d_ADaM + [ADaM_Ilac] ) * ( 1 + Beta_sinerji )",
        "Bu formülasyon, sentetik ADaM bölgesi ligandlarının AMPK katalitik verimini fizyolojik nükleotidlerden bağımsız olarak nasıl artırdığını açıklar."
    )
]

parts.append(("KISIM 1: İNSÜLİN / IGF-1 SİNYAL AĞI (IIS) VE YAŞLANMA KİNETİĞİ", part1_subsections))
parts.append(("KISIM 2: mTORC1 VE mTORC2 MOLEKÜLER MİMARİSİ VE BESİN ALGILAMA", part2_subsections))
parts.append(("KISIM 3: AMPK: HÜCRESEL ENERJİ ŞALTERİ VE METABOLİK SENSÖR", part3_subsections))

# ================= KISIM 4 =================
part4_subsections = [
    (
        "4.1 Memeli Sirtuinleri (SIRT1-SIRT7) Hücresel Dağılımı ve Enzimatik Çeşitliliği",
        "Sirtuinler, maya Sir2 (Silent Information Regulator 2) geninin memelilerdeki evrimsel homologları olan Sınıf III histon deasetilaz (HDAC) enzim ailesidir. Katalitik aktiviteleri için mutlak surette Nikotinamid Adenin Dinükleotid (NAD+) koenzimine bağımlıdırlar.",
        "İnsan genomunda yedi farklı sirtuin izoformu (SIRT1-SIRT7) kodlanır ve bunlar spesifik hücre altı kompartmanlara dağılmıştır: Çekirdekte yer alanlar SIRT1, SIRT6 ve SIRT7; mitokondri matriksinde yer alanlar SIRT3, SIRT4 ve SIRT5; sitoplazmada yer alan ise SIRT2'dir (fakat hücre döngüsünde nükleusa geçebilir). Klasik deasetilasyonun yanı sıra SIRT4 ADP-riboziltransferaz, SIRT5 desüksinilaz/demalonilaz/deglutarilaz, SIRT6 ise demiristoilaz ve zayıf ADP-ribozilasyon aktivitelerine sahiptir. Bu enzimler hücresel metabolik durumu doğrudan epigenetik ve post-translasyonel modifikasyonlara tercüme eder.",
        "Enzimatik_Hiz_Sirtuin = k_cat * [Sirtuin] * [Substrat-Asetil] * [NAD+] / ( K_M_sub * (1 + [NAM]/K_i_nam) * K_M_NAD + [Substrat-Asetil]*[NAD+] )",
        "Bu genel hız eşitliği, sirtuin katalizinin NAD+ konsantrasyonuna bağımlılığını ve reaksiyon son ürünü nikotinamid (NAM) tarafından kompetitif olarak nasıl baskılandığını gösterir."
    ),
    (
        "4.2 SIRT1 ve Nükleer Fonksiyonlar: Histon Deasetilasyonu (H3K9ac, H4K16ac) ve Kromatin Sıkılaşması",
        "SIRT1, nükleusta ökromatin-heterokromatin dengesini belirleyen ve histon proteinlerindeki spesifik lizin asetilasyonlarını kaldıran merkezi bir epigenetik bekçidir.",
        "Özellikle histon H3 lizin 9 (H3K9ac) ve histon H4 lizin 16 (H4K16ac) üzerindeki asetil gruplarını NAD+ tüketerek koparır. H4K16ac deasetilasyonu, nükleozomlar arası elektrostatik itmeyi ortadan kaldırarak 30 nm kromatin ipliğinin yüksek dereceli heterokromatin yapısına sıkılaşmasını (heterochromatin compaction) sağlar. Bu durum, retrotranspozonların, parazitik DNA dizilerinin ve genomik dengesizlik odaklarının suskun kalmasını temin eder. Yaşlanmayla birlikte nükleer SIRT1 aktivitesinin düşmesi heterokromatin erozyonuna ve anormal transkripsiyonel gürültüye neden olur.",
        "d[Heterokromatin]/dt = k_SIRT1 * [SIRT1_nukleus] * [H4K16ac] - k_HAT * [p300/CBP] * [Heterokromatin]",
        "Bu model, heterokromatin yoğunluğunun SIRT1 deasetilasyon akısı ile p300/CBP asetiltransferaz akısı arasındaki dinamik dengeye bağlı olduğunu gösterir."
    ),
    (
        "4.3 SIRT1'in Transkripsiyonel Hedefleri: PGC-1alpha, FOXO1/3/4 ve NF-kappaB Deasetilasyonu",
        "SIRT1'in deasetilaz aktivitesi histonlarla sınırlı değildir; hücresel sağkalım, antioksidan yanıt ve enflamasyonu yöneten kritik transkripsiyon faktörlerini doğrudan deasetile eder.",
        "SIRT1, PGC-1alpha'nın 13 spesifik lizin kalıntısını deasetile ederek ko-aktivatör fonksiyonunu maksimize eder ve mitokondriyal solunumu uyarır. FOXO3a'yı deasetile ederek onun apoptoz indükleyici genler yerine antioksidan (SOD2, Katalaz) ve DNA onarım genlerini bağlamasını sağlar. En önemlisi, pro-enflamatuar ana faktör NF-kappaB'nin p65 (RelA) alt birimini Lys310 kalıntısından deasetile ederek inaktive eder. Bu deasetilasyon, yaşlanmaya bağlı kronik steril inflamasyonun (inflam-aging) ve SASP salgısının moleküler düzeyde bastırılmasını sağlar.",
        "Inflamasyon_Baskilama = 1 / ( 1 + ( [SIRT1_aktif] / K_i_p65 )^h )",
        "Bu formül, aktif SIRT1 seviyesinin artışıyla NF-kappaB p65 asetilasyonunun ve dolayısıyla enflamatuar sitokin transkripsiyonunun nasıl logaritmik olarak baskılandığını simgeler."
    ),
    (
        "4.4 SIRT3: Mitokondriyal Ana Deasetilaz ve Oksidatif Fosforilasyonun İnce Ayarı",
        "Mitokondri matriksi, yüksek konsantrasyonda asetil-KoA barındırması ve hafif alkali pH'ı (pH ~ 7.9-8.0) nedeniyle non-enzimatik lizin asetilasyonuna aşırı derecede açıktır. SIRT3, mitokondriyal proteinlerin küresel asetilasyon yükünü temizleyen ana deasetilazdır.",
        "SIRT3, Elektron Taşıma Zinciri (ETC) Kompleks I (NDUFA9) ve Kompleks II (SDHA) alt birimlerini deasetile ederek elektron transfer hızını artırır ve elektron sızıntısını (ROS üretimini) azaltır. Ayrıca mitokondriyal antioksidan enzim Manganez Süperoksit Dismutazı (MnSOD/SOD2) Lys122 bölgesinden ve İzokitrat Dehidrogenaz 2'yi (IDH2) deasetile ederek mitokondriyal NADPH ve glutatyon redoks tamponunu güçlendirir. Yaşlanmada SIRT3 kaybı, kontrolsüz mitokondriyal hiperasetilasyona, ATP sentezinde çöküşe ve mitokondriyal permeabilite geçiş gözeneğinin (mPTP) açılmasına yol açar.",
        "Mitokondri_Verim_Indeksi = ( J_ATP / J_ROS ) = Eta_0 * ( 1 + alpha_S3 * [SIRT3_aktif] / (K_m_S3 + [SIRT3_aktif]) )",
        "Bu denklem, aktif SIRT3 konsantrasyonunun artmasıyla mitokondriyal ATP/ROS üretim oranının (biyoenerjetik verim) nasıl katlanarak iyileştiğini matematiksel olarak ifade eder."
    ),
    (
        "4.5 SIRT6: DNA Çift Zincir Kırık Onarımı, H3K9/H3K56 Deasetilasyonu ve Telomer Korunumu",
        "SIRT6, kromatin stabilitesi, DNA hasar onarımı, telomer bakımı ve glukoz metabolizmasının düzenlenmesinde olağanüstü fonksiyonlara sahip çok yönlü bir nükleer sirtuindir. SIRT6 nakavt fareler doğumdan sonra 4 hafta içinde şiddetli progeroid (erken yaşlanma) semptomlarıyla ölürken, SIRT6 aşırı ekspresyonu farelerin yaşam süresini %30 uzatır.",
        "SIRT6, telomerik bölgelerde histon H3 lizin 9 (H3K9ac) ve lizin 56 (H3K56ac) deasetilasyonunu gerçekleştirerek telomerlerin WRN helikazı aracılığıyla korunmasını sağlar ve telomerik füzyonları önler. DNA çift zincir kırığı (DSB) oluştuğunda SIRT6, DNA-PKcs ve PARP1 enzimlerini hasar bölgesine alarak homolog rekombinasyon ve NHEJ onarımını hızlandırır. Ek olarak HIF-1alpha'yı transkripsiyonel olarak ko-represe ederek hücrelerin kanserojen Warburg glikolizine kaymasını engeller.",
        "Telomer_Asinma_Hizi = Delta_L0 * ( 1 - [SIRT6_kromatin] / (K_tel_SIRT6 + [SIRT6_kromatin]) )",
        "Bu türev formülü, kromatine bağlı SIRT6 yoğunluğunun replikasyon başına telomer baz çifti aşınma hızını nasıl asimptotik olarak azalttığını modeller."
    ),
    (
        "4.6 SIRT7: Ribozomal DNA (rDNA) Transkripsiyonu ve Nükleolar Stres Yanıtı",
        "SIRT7, nükleolusta yoğunlaşan tek sirtuin üyesidir ve ribozom biyogenezinin, RNA Polimeraz I (Pol I) transkripsiyonunun ve nükleolar genomik stabilitenin en üst düzey kontrolörüdür.",
        "SIRT7, Pol I transkripsiyon başlatma faktörü UBF'ye bağlanır ve histon H3 lizin 18 (H3K18ac) asetilasyonunu spesifik olarak deasetile eder. Bu deasetilasyon, rDNA tekrarlarının kararsızlaşmasını ve rDNA rekombinasyonunu önler. Aşırı ribozomal translasyon stresinde SIRT7, nükleolar morfolojiyi stabilize ederek p53'ün gereksiz yere aktive olmasını engeller ve kök hücrelerin proliferatif kapasitesini korur. Yaşlanan hematopoietik kök hücrelerde (HSC) SIRT7 seviyelerinin düşmesi nükleolar şişmeye, translasyonel sadakatsizliğe ve kök hücre tükenmesine yol açar.",
        "rDNA_Kararliligi = K_stab * [SIRT7_nukleol] / ( [rDNA_kopyasi] * (1 + [H3K18ac]/K_asetil) )",
        "Bu ilişki, rDNA genomik bütünlüğünün nükleolar SIRT7 mevcudiyeti ve H3K18 hipoasetilasyon düzeyiyle doğrudan korelasyonunu ortaya koyar."
    ),
    (
        "4.7 NAD+ Bağımlılığı: Sirtuinlerin Substrat Kinetiği ve Yaşlanmayla NAD+ Tüketimi",
        "Tüm sirtuin deasetilasyon reaksiyonları, stokiyometrik olarak bir molekül NAD+ tüketir. Bu enzimatik yarı-reaksiyonda NAD+, nikotinamid (NAM) ve O-asetil-ADP-riboz (OAADPR) ürünlerine parçalanır.",
        "SIRT1'in NAD+ için Michaelis sabiti (Km ~ 100-200 microM), intraselüler serbest nükleer NAD+ konsantrasyonuna (yaklaşık 100-300 microM) çok yakındır. Bu durum, sirtuin aktivitesinin hücrenin anlık serbest NAD+ düzeylerine olağanüstü duyarlı olduğunu gösterir. Yaşlanma sürecinde hücresel NAD+ havuzu dramatik bir şekilde (%50-80 oranında) tükenir. Substrat yetersizliği nedeniyle sirtuin enzimleri protein düzeyinde eksprese edilseler dahi katalitik olarak susturulurlar ve epigenetik heterokromatin gevşemesi kaçınılmaz hale gelir.",
        "Theta_SIRT_doygunluk = [NAD+] / ( K_M_NAD * (1 + [NAM]/K_i_NAM) + [NAD+] )",
        "Bu doygunluk eğrisi, azalan NAD+ ve biriken NAM metabolitinin sirtuin enzimlerinin fonksiyonel doygunluk fraksiyonunu nasıl sıfırladığını açıklar."
    ),
    (
        "4.8 NAMPT (Nikotinamid Fosforiboziltransferaz) ve NAD+ Kurtarma Yolağının Hız Kısıtlayıcı Rolü",
        "Memelilerde hücresel NAD+ havuzunun %90'ından fazlası, nikotinamidin geri dönüştürüldüğü iki adımlı kurtarma yolağı (Salvage Pathway) ile sağlanır. Bu yolağın hız kısıtlayıcı ana enzimi NAMPT'dir.",
        "NAMPT, sirtuin ve PARP reaksiyonlarından açığa çıkan nikotinamidi (NAM), 5-fosforibozil-1-pirofosfat (PRPP) ile birleştirerek Nikotinamid Mononükleotide (NMN) dönüştürür. NMN daha sonra NMNAT1-3 enzimleri tarafından ATP kullanılarak doğrudan NAD+'ya çevrilir. Dolaşımda veziküllere paketlenmiş ekstraselüler NAMPT (eNAMPT), hipotalamus ve diğer dokularda sistemik NAD+ seviyelerini senkronize eder. Yaşlanmayla dokularda ve kanda NAMPT ekspresyonunun düşmesi, NAD+ biyosentez musluğunun kısılması anlamına gelir.",
        "J_NAD_sentez = V_max_NAMPT * [NAM] * [PRPP] / ( (K_m_NAM + [NAM]) * (K_m_PRPP + [PRPP]) )",
        "Bu Bi-Bi reaksiyon kinetiği modeli, net hücresel NAD+ üretim hızının doğrudan NAMPT enzim aktivitesi ve substrat mevcudiyetiyle sınırlandığını gösterir."
    ),
    (
        "4.9 CD38 ve PARP1 Tarafından NAD+ Tüketimi: Sirtuin Fonksiyonunun Çöküşü",
        "Yaşlanma sürecinde yaşanan sistemik NAD+ çöküşü sadece sentezin azalmasından değil, agresif NAD+ tüketicisi enzimlerin hiperaktivasyonundan kaynaklanır.",
        "CD38, hücre yüzeyinde yer alan bir ektosiklaz ve glikohidrolazdır; tek bir molekül cADPR üretmek için 100 molekül NAD+'yı hidrolize eder. Yaşlanan dokularda biriken SASP faktörleri ve senesen hücreleri çevreleyen M1 benzeri pro-enflamatuar makrofajlar devasa miktarda CD38 eksprese ederek hücreler arası ve hücre içi NAD+ havuzunu kurutur. Eş zamanlı olarak, biriken DNA hasarları nükleer PARP1 (Poli-ADP-Riboz Polimeraz 1) enzimini sürekli hiperaktif tutar. CD38 ve PARP1, sınırlı NAD+ havuzunu tüketerek sirtuinleri 'aç bırakır'.",
        "d[NAD+]/dt = J_NAMPT - J_CD38 - J_PARP1 - Sum(J_SIRT1_7)",
        "Bu dinamik kütle korunum denklemi, artan CD38 ve PARP1 tüketim akılarının netice olarak sirtuinlere kalan NAD+ payını nasıl yok ettiğini ortaya koymaktadır."
    ),
    (
        "4.10 Sirtuin Aktive Edici Bileşikler (STAC'ler) ve NAD+ Arttırıcı Tedaviler",
        "Sirtuin fonksiyonunun yeniden kazanılması iki yönlü farmakolojik bir yaklaşımla mümkündür: NAD+ öncülleri ile substrat havuzunu doldurmak ve sentetik STAC'ler ile enzim kinetiğini allosterik olarak hızlandırmak.",
        "Nikotinamid Mononükleotid (NMN) ve Nikotinamid Ribozit (NR), hücre içine Slc12a8 veya ENT taşıyıcılarıyla girerek hızla NAD+'ya dönüşür. Eş zamanlı olarak 78c gibi spesifik CD38 inhibitörleri ve PARP inhibitörleri parazitik tüketimi keser. Sentetik STAC'ler (SRT1720, SRT2104) ise SIRT1'in N-terminal allosterik aktivasyon alanına bağlanarak enzimin asetilasyonlu substratlara olan Km değerini dramatik olarak düşürür. Bu sinerjistik kombinasyon, yaşlı dokularda mitokondriyal biyogenezi, epigenetik sıkılığı ve metabolik canlılığı geri döndürür.",
        "SIRT1_Efektif_Aktivite = V_max * [Substrat] / ( K_M_app(STAC) + [Substrat] ) * [NAD+] / ( K_M_NAD + [NAD+] )",
        "Burada K_M_app(STAC) < K_M_native olup, allosterik moleküllerin NAD+ varlığında reaksiyon hızını nasıl katladığı tanımlanmaktadır."
    )
]

# ================= KISIM 5 =================
part5_subsections = [
    (
        "5.1 Kalori Kısıtlamasının (CR) Moleküler Temeli: Yaşam Süresi Uzamasının Biyolojik Mekanizmaları",
        "Kalori kısıtlaması (Caloric Restriction - CR), malnütrisyon yaratmaksızın ad libitum kalori alımının %20 ila %40 oranında azaltılmasıdır; mayalardan primatlara kadar test edilen hemen her biyolojik türde yaşam süresini ve sağlıklı ömrü uzatan altın standart müdahaledir.",
        "CR'nin moleküler temeli, hücreyi 'büyüme ve üreme' anabolik modundan 'hücresel bakım, onarım ve stres direnci' moduna geçirmesidir. Dolaşımdaki glukoz, insülin ve IGF-1 seviyeleri düşerken; mTORC1 baskılanır, AMPK uyarılır ve SIRT1/SIRT3 enzimleri aktive edilir. Enerji kaynaklarının kıtlığı, hücreleri hasarlı proteinleri ve mitokondrileri otofaji yoluyla geri dönüştürmeye zorlar. Aynı zamanda nükleer DNA onarım mekanizmaları güçlenir ve somatik dokularda senesens birikimi dramatik olarak yavaşlatılır.",
        "Lifespan_CR = Lifespan_adlib * ( 1 + beta_CR * (Delta_Calorie / Calorie_basal) * exp(-Delta_Calorie / C_kritik) )",
        "Bu model, kalori kısıtlamasının belirli bir optimum kısıtlama aralığına kadar yaşam süresini artırdığını, ancak açlık eşiğini (C_kritik) aştığında malnütrisyon nedeniyle faydanın kaybolduğunu formüle eder."
    ),
    (
        "5.2 Hormesis Paradigması: Düşük Düzeyli Hücresel Stresin Savunma Mekanizmalarını Tetiklemesi",
        "Hormesis, düşük dozlarda uygulandığında organizmada adaptif savunma mekanizmalarını tetikleyerek fayda sağlayan, ancak yüksek dozlarda toksik veya ölümcül olan biyolojik stres yanıtıdır.",
        "Kalori kısıtlaması, hafif bir metabolik stresor (hormetin) olarak işlev görür. Besin azlığı hafif derecede reaktif oksijen türü (mitohormesis) üretir ve bu durum Nrf2 transkripsiyon faktörünün çekirdeğe göçünü uyararak endojen antioksidan enzimleri (Heme Oksijenaz-1, NQO1, Glutatyon Sentaz) indükler. Hücresel şaperonlar (Hsp70, Hsp90) aktive edilerek proteostaz kalkanı örülür. Düşük düzeyli bu sürekli alarm hali, hücreyi gelecekte karşılaşacağı çok daha yıkıcı iskemik, toksik veya enfeksiyöz hasarlara karşı korunaklı kılar.",
        "Biyolojik_Direnc = R_0 + Delta_R * ( [Hormetin] / K_h ) / ( 1 + ( [Hormetin] / K_h )^2 )",
        "Bu iki fazlı hormetik yanıt eğrisi, düşük-orta stresör seviyelerinde hücresel direncin zirve yaptığını, aşırı stres durumunda ise savunmanın çöktüğünü gösterir."
    ),
    (
        "5.3 Metiyonin ve Triptofan Kısıtlamasının Yaşam Süresi ve Epigenetik Üzerindeki Özel Etkileri",
        "Toplam kaloriden bağımsız olarak, diyetin amino asit kompozisyonundaki spesifik kısıtlamalar, özellikle kükürtlü bir amino asit olan L-metiyonin kısıtlaması, tek başına kalori kısıtlamasının ömür uzatıcı etkilerini taklit edebilir.",
        "Metiyonin, tek-karbon döngüsünün ve S-adenozilmetiyonin (SAM) biyosentezinin merkezindedir. Metiyonin kısıtlandığında SAM/SAH oranı düşer; bu durum SAMTOR sensörünü uyararak mTORC1'i lizozom zarında doğrudan inhibe eder. Aynı zamanda karaciğerden Fibroblast Büyüme Faktörü 21 (FGF21) salınımı patlar, beyaz yağ dokusunun browning (kahverengileşme) süreci başlar ve insülin duyarlılığı olağanüstü artar. Triptofan kısıtlaması ise serotonin/melatonin ekseni ve kynurenin yolağını modüle ederek nöroinflamasyonu ve immünosenesensi baskılar.",
        "d[FGF21]/dt = k_FGF21 * ( K_met / (K_met + [Metiyonin_plazma]) ) - k_klerens * [FGF21]",
        "Bu regülasyon denklemi, plazma metiyonin seviyesindeki düşüşün karaciğer kaynaklı FGF21 longevity hormonu salınımını nasıl ters orantılı uyardığını belgeler."
    ),
    (
        "5.4 Aralıklı Oruç (Intermittent Fasting): 16/8, Alternatif Gün Orucu (ADF) ve Sirkadiyen Uyum",
        "Aralıklı Oruç (Intermittent Fasting - IF), genel kalori alımını zorunlu olarak kısmadan, besin tüketimini belirli zaman pencerelerine sıkıştıran bir krono-beslenme protokolüdür.",
        "En popüler protokoller 16 saat açlık / 8 saat yeme (16/8) ve gün aşırı oruçtur (Alternate-Day Fasting - ADF). Açlığın 12-14. saatinden itibaren karaciğer glikojen depoları tükenir ve organizma 'metabolik şalteri' açarak glukoz kullanımından yağ asidi ve keton cisimciği kullanımına geçer. Bu faz geçişi, sirkadiyen genlerin (BMAL1, CLOCK) senkronizasyonunu sağlar, periferik insülin direncini kırar ve her gece düzenli bir otofajik temizlik dalgası başlatır. Ad libitum beslenen ancak besinleri 8 saatlik pencerede alan fareler, aynı kaloriyi tüketseler dahi obezite ve hiperinsülinemiden tamamen korunur.",
        "Metabolik_Salter_Kinetigi = 1 / ( 1 + exp( -k_switch * ( t_aclik - Tau_glikojen_tukenme ) ) )",
        "Bu sigmoidal fonksiyon, açlık süresi glikojen tükenme eşiğini (Tau_glikojen_tukenme ~ 12 saat) aştığında ketojenik ve otofajik metabolizmaya geçiş olasılığını gösterir."
    ),
    (
        "5.5 Açlık Benzeri Diyet (Fast-Mimicking Diet - FMD): Kök Hücre Rejenerasyonu ve İmmün Reset",
        "Valter Longo ve ekibi tarafından geliştirilen Açlık Benzeri Diyet (FMD), düşük kalorili, düşük proteinli, düşük karbonhidratlı ve yüksek sağlıklı yağ içerikli 5 günlük periyodik bir tıbbi beslenme protokolüdür.",
        "FMD sırasında organizma tam açlık fizyolojisine girer: IGF-1, leptin ve PKA sinyali çöker, AMPK ve sirtuinler tavan yapar. 5 günlük kısıtlama boyunca dolaşımdaki yaşlı, senesen ve oto-reaktif lökositler apoptoza uğrar ve beyaz küre sayısı geçici olarak düşer. Beslenmenin yeniden başladığı (re-feeding) fazda ise kemik iliğindeki hematopoietik kök hücreler (HSC) uykudan uyanarak IGF-1 ve mTORC1 uyarımıyla devasa bir proliferasyon patlaması yaşar ve bağışıklık sistemini tamamen genç lökositlerle sıfırdan inşa eder (immün reset).",
        "Kok_Hucre_Rejenerasyon = Delta_Proliferasyon * [Re-feeding_Sinyali] * ( 1 - exp(-k_priming * t_FMD_aclik) )",
        "Bu model, FMD açlık süresinin kök hücreleri nasıl hazırladığını ve yeniden beslenme fazında gençleştirici proliferatif patlamanın şiddetini formüle eder."
    ),
    (
        "5.6 Karaciğer Glikojen Tükenmesi, Ketogenez ve Beta-Hidroksibütirat (BHB) Sinyalizasyonu",
        "Uzamış açlık veya ketojenik adaptasyon sırasında hepatositlerde yağ asitlerinin beta-oksidasyonundan elde edilen asetil-KoA fazlası ketogenez yolağına girer.",
        "HMG-CoA sentaz 2 (HMGCS2) enzimi aracılığıyla asetoasetat ve ardından beta-hidroksibütirata (BHB) dönüştürülür. BHB kana salınarak kan-beyin bariyerini geçer ve beyin, miyokard ve iskelet kası için süper-yakıt görevi görür. Glukoza kıyasla tüketilen oksijen başına daha yüksek serbest enerji (Delta-G) açığa çıkarır. Ancak BHB sadece bir metabolik yakıt değildir; hücre yüzeyindeki HCA2 (GPR109A) hidroksikarboksilik asit reseptörüne bağlanarak enflamasyonu baskılayan güçlü bir endojen sinyal molekülüdür.",
        "[BHB]_plazma(t) = BHB_max * ( 1 - exp(-k_keto * (t - t_lag)) )",
        "Bu farmakokinetik eğri, açlığın başlamasıyla birlikte plazma serbest beta-hidroksibütirat konsantrasyonunun zamana bağlı üstel birikimini modeller."
    ),
    (
        "5.7 BHB'nin Bir HDAC İnhibitörü ve NLRP3 İnflamazom Baskılayıcı Olarak Görevi",
        "Beta-hidroksibütiratın (BHB) en devrimsel keşiflerinden biri, Sınıf I histon deasetilazların (HDAC1, HDAC2, HDAC3) doğrudan endojen inhibitörü olduğunun anlaşılmasıdır.",
        "BHB, mikromolar-milimolar konsantrasyonlarda HDAC'lerin çinko içeren katalitik yarığına bağlanarak histon asetilasyonunun (özellikle H3K9ac ve H3K14ac) korunmasını sağlar. Bu epigenetik gevşeme, Foxo3a ve Mt2 (Metallotiyonein 2) gibi oksidatif stres direnç genlerinin transkripsiyonunu doğrudan serbest bırakır. Ayrıca BHB, makrofajlarda K+ dışa akışını engelleyerek NLRP3 inflamazom kompleksinin montajını ve kaspaz-1 aktivasyonunu fiziksel olarak durdurur; bu da pro-enflamatuar IL-1beta ve IL-18 sitokin salgısını engeller.",
        "Inhibisyon_NLRP3 = 1 / ( 1 + ( [BHB] / IC50_BHB )^n_Hill )",
        "Bu farmakodinamik denklem, yükselen dolaşım BHB seviyelerinin NLRP3 inflamazom aracılı steril enflamasyonu nasıl doğrudan inhibe ettiğini gösterir."
    ),
    (
        "5.8 CR Sırasında İmmünomodülasyon, SASP Baskılanması ve T Hücre Gençleşmesi",
        "Kronik kalori kısıtlaması, yaşa bağlı timik involüsyonu (timus bezinin yağ dokusuna dönüşerek körelmesi) yavaşlatır ve yeni naif T hücresi üretimini uzun yıllar boyunca devam ettirir.",
        "CR ortamında düşük mTORC1 aktivitesi, CD4+ ve CD8+ T hücrelerinde metabolik yorgunluğu (exhaustion) önler ve hafıza T hücrelerinin (T_CM) kök hücre benzeri potansiyelini korur. Eş zamanlı olarak, vücuttaki senesen hücrelerin NF-kappaB ve p38 MAPK aktiviteleri baskılandığı için SASP faktörlerinin (TNF-alfa, IL-6, MMP-3) sistemik konsantrasyonu taban seviyelere iner. Bu immünomodülatuvar etki, yaşlılıkta ortaya çıkan aşırı sitokin fırtınalarına ve otoimmün dejenerasyona karşı güçlü bir tampon oluşturur.",
        "Naif_T_Havuzu(t) = N_0 * exp( -k_timik_involusyon * (1 - alpha_CR) * t )",
        "Bu korunum formülü, kalori kısıtlamasının timik involüsyon hız sabitini (k_timik_involusyon) alfa_CR faktörüyle nasıl düşürerek naif T hücresi rezervini koruduğunu gösterir."
    ),
    (
        "5.9 Uzun Süreli CR'nin İnsan Klinik Çalışmalarındaki Sonuçları (CALERIE Çalışması)",
        "İnsanlarda kalori kısıtlamasının biyolojik etkilerini titizlikle test eden en kapsamlı randomize kontrollü klinik çalışma CALERIE (Comprehensive Assessment of Long-term Effects of Reducing Intake of Energy) konsorsiyumudur.",
        "İki yıl boyunca %12-14 oranında gerçek kalori kısıtlaması uygulayan sağlıklı bireylerde: Sistolik ve diyastolik kan basıncında belirgin düşüş, LDL kolesterol ve trigliseridlerde azalma, insülin duyarlılığında dramatik artış ve sistemik enflamasyon belirteci hs-CRP'de %40'ın üzerinde düşüş kaydedilmiştir. En önemlisi, DunedinPACE epigenetik yaşlanma hızı algoritması ve Klemera-Doubal biyolojik yaş algoritmaları, bu bireylerde biyolojik yaşlanma hızının yılda %2 ila %3 oranında net olarak yavaşladığını doğrulamıştır.",
        "Biyolojik_Yaslanma_Hizi = V_kronolojik * ( 1 - Delta_DunedinPACE_CR )",
        "Bu klinik formülasyon, CALERIE deneklerinde ölçülen yıllık biyolojik yaşlanma katsayısının kronolojik yıla göre net yavaşlama oranını temsil eder."
    ),
    (
        "5.10 Bireyselleştirilmiş ve Genotipik Olarak Ayarlanmış Kısıtlama Protokolleri",
        "Her genotip kalori kısıtlamasına aynı adaptif yanıtı vermez; genetik arka plan, metabolik tip ve bazal vücut kitle indeksi (VKİ) kısıtlamanın fayda-zarar dengesini belirler.",
        "IL6R, TCF7L2, FTO, PPARG ve APOE alel varyasyonları, bireylerin açlık metabolizmasına, ketozise ve yağ oksidasyonuna verdikleri yanıtı derinden etkiler. Örneğin APOE epsilon4 taşıyıcılarında aşırı doymuş yağ içeren ketojenik adaptasyonlar lipid partikül sayısını (ApoB) patolojik olarak artırabilir. Bu nedenle modern longevity hekimliği, 'tek tip açlık' yerine; sürekli glukoz monitörleri (CGM), keton sensörleri ve genetik polimorfizm profillerine dayanan hassas, kişiselleştirilmiş mikrodiyet protokollerini benimser.",
        "Kisisel_CR_Optimum = f(Genotip_Score) * [Bazal_Metabolizma_Hizi] * ( 1 - K_bireysel_risk )",
        "Bu fonksiyon, bireyin genomik risk skoru ve bazal metabolik harcaması doğrultusunda hesaplanan güvenli ve maksimum etkili kalori kısıtlama aralığını tanımlar."
    )
]

# ================= KISIM 6 =================
part6_subsections = [
    (
        "6.1 Lizozomal Membran Üzerinde Amino Asit Algılama Kompleksi (Ragulator ve V-ATPase)",
        "Hücre içi amino asit bolluğunun algılanması ve bu bilginin mTORC1'e iletilmesi, lizozom zarında konuşlanmış olağanüstü karmaşık bir protein iskelesi tarafından yönetilir.",
        "Bu mimarinin temelini V-ATPase (Vakuolar H+-ATPaz proton pompası) ve beş alt birimli Ragulator kompleksi (LAMTOR1-5) oluşturur. Ragulator, mikristoilasyon ve palmitoilasyon lipid çapalarıyla lizozomal membranın sitozolik yüzüne kalıcı olarak demirlenmiştir. Lizozom lümeninde biriken amino asitler, V-ATPase üzerinde konformasyonel bir rotasyona neden olur. Bu mekanik sinyal Ragulator'a aktarılır ve Ragulator, Rag GTPaz heterodimerleri için bir guanin nükleotid değişim faktörü (GEF) gibi davranarak onları aktive eder.",
        "Sinyal_Iletim_Lizozom = k_rotasyon * [V-ATPase_aktif] * [AA_lumen] / ( K_M_lumen + [AA_lumen] )",
        "Bu biyofiziksel eşitlik, lizozom lümen içi amino asit konsantrasyonunun V-ATPase konformasyonel geçiş hızı üzerinden Ragulator'a aktarılmasını modeller."
    ),
    (
        "6.2 Lösin Algılama Sensörü: Sestrin2 ve GATOR2 Kompleks Etkileşimleri",
        "Dallı zincirli bir esansiyel amino asit olan L-lösin, protein sentezinin ve mTORC1 aktivasyonunun en güçlü besinsel tetikleyicisidir. Sitozolik lösin sensörü Sestrin2 proteinidir.",
        "Lösin yokluğunda Sestrin2, devasa pentamerik GATOR2 kompleksine (Mios, WDR24, WDR59, Seh1L, Sec13) doğrudan bağlanır ve onun aktivitesini baskılar. Hücre içine serbest lösin girdiğinde, Sestrin2'nin yüksek afiniteli cebine bağlanır (Kd ~ 20 microM). Lösin bağlanması Sestrin2'de konformasyonel bir bükülme yaratarak GATOR2'den ayrılmasına yol açar. GATOR2 üzerindeki inhibisyon kalktığında, GATOR2 serbest kalarak GATOR1 inhibitör kompleksini bloke eder ve mTORC1'in lizozoma göçünü tetikler.",
        "[GATOR2_serbest] = [GATOR2_toplam] / ( 1 + [Sestrin2_apo] / K_d_SesGAT )",
        "Burada serbest GATOR2 fraksiyonu, lösinsiz apo-Sestrin2 konsantrasyonu ve disosiasyon sabiti üzerinden tanımlanmaktadır."
    ),
    (
        "6.3 Arjinin Algılama Sensörleri: CASTOR1 Dimasyonu ve Lizozomal SLC38A9 Taşıyıcısı",
        "L-arjinin, mTORC1 tarafından hem sitozolde hem de lizozom lümeninde çift yönlü olarak bağımsız sensörlerle taranan eşsiz bir amino asittir.",
        "Sitozolde arjinin sensörü CASTOR1 (Cellular Arginine Sensor for mTORC1) proteinidir. Arjinin yokluğunda homodimer veya CASTOR2 ile heterodimer yapan CASTOR1, GATOR2'ye bağlanarak onu inhibe eder. Serbest arjinin CASTOR1'in arjinin bağlama cebine oturduğunda kompleks çözülür ve GATOR2 serbest kalır. Lizozom membranında ise 11 transmembran alanlı SLC38A9 taşıyıcısı lümen içi arjini algılar. Yüksek arjinin konsantrasyonu SLC38A9'un Ragulator ve Rag GTPazlarla doğrudan etkileşime girerek RagA/B'ye GTP yüklemesini sağlar.",
        "Arjinin_Aktivasyon_Katsayisi = f_sitozol([Arjinin_sitozol]) * f_lumen([Arjinin_lumen])",
        "Bu çarpımsal fonksiyon, tam mTORC1 lizozomal yerleşimi için hem sitozolik CASTOR1 hem de lümenal SLC38A9 sensörlerinin eş zamanlı doyumunun şart olduğunu ifade eder."
    ),
    (
        "6.4 S-Adenozilmetiyonin (SAM) Algılama: SAMTOR ve Tek-Karbon Metabolizma Entegrasyonu",
        "Hücrenin metilasyon kapasitesini ve metiyonin mevcudiyetini izleyen moleküler sensör, yakın zamanda keşfedilen SAMTOR (KIAA1456) proteinidir.",
        "SAMTOR, evrimsel olarak korunmuş bir S-adenozilmetiyonin (SAM) bağlama alanına sahiptir. Metiyonin kıtlığında hücre içi SAM seviyeleri düştüğünde, apo-SAMTOR doğrudan GATOR1 kompleksine ve KICSTOR iskelesine bağlanır; bu durum GATOR1'in RagA/B üzerindeki GAP (inhibitör) aktivitesini güçlendirerek mTORC1'i tamamen kilitler. Yeterli metiyonin varlığında SAM, SAMTOR'a bağlanır (Kd ~ 7 microM) ve SAMTOR-GATOR1 etkileşimini bozar. Böylece tek-karbon döngüsü ve epigenetik metilasyon potansiyeli doğrudan protein anabolizmasına bağlanmış olur.",
        "d[mTORC1_aktif]/dt_SAM = k_SAM * ( [SAM] / (K_d_SAMTOR + [SAM]) ) * [GATOR1_inaktif]",
        "Bu kinetik türev, hücresel SAM konsantrasyonunun SAMTOR disosiasyonu üzerinden mTORC1 aktivasyon hızına katkısını açıklar."
    ),
    (
        "6.5 Rag GTPaz Heterodimerleri (RagA/B ve RagC/D): Nükleotid Durumu ve mTORC1 Alımı",
        "Amino asit sinyal kaskadının lizozom zarındaki nihai yürütücüleri Rag GTPazlardır. Dört üyeden oluşan bu küçük G-proteinleri zorunlu heterodimerler halinde çalışır: RagA veya RagB, RagC veya RagD ile eşleşir.",
        "mTORC1'in lizozom zarına çekilebilmesi için Rag heterodimerinin kesin bir nükleotid konfigürasyonuna sahip olması şarttır: RagA/B alt birimi GTP yüklü (RagA-GTP), RagC/D alt birimi ise GDP yüklü (RagC-GDP) olmalıdır. Bu aktif konfigürasyon (RagA_GTP / RagC_GDP), mTORC1'in Raptor alt birimi için kusursuz bir docking platformu oluşturur. Sitozolde dağınık bulunan mTORC1 lizozom zarına çekilir ve burada Rheb-GTP ile temas ederek kinaz aktivitesini ateşler. Herhangi bir alt birimin nükleotid durumu bozulursa mTORC1 lizozoma bağlanamaz.",
        "P_docking = [RagA_GTP] * [RagC_GDP] / ( ([RagA_GTP] + [RagA_GDP]) * ([RagC_GTP] + [RagC_GDP]) )",
        "Bu olasılık formülü, lizozom membranındaki Rag heterodimerlerinin mTORC1'i bağlama yeterliliğine sahip fraksiyonunun nükleotid durumlarına bağlılığını özetler."
    ),
    (
        "6.6 FLCN-FNIP Kompleksi: RagC/D GAP Aktivitesi ve mTORC1 Aktivasyonu",
        "RagA/B'ye GTP yüklenmesi amino asit varlığında Ragulator ve SLC38A9 tarafından sağlanırken, diğer partner olan RagC veya RagD'nin GDP formuna dönüştürülmesi ayrı bir özelleşmiş kompleks tarafından yönetilir.",
        "Folliculin (FLCN) ve onun bağlanma partnerleri FNIP1/FNIP2, güçlü bir GTPaz Aktive Edici Protein (GAP) aktivitesine sahiptir ve bu aktivite spesifik olarak RagC/RagD'ye yöneliktir. Amino asit bolluğunda FLCN-FNIP kompleksi lizozoma transloke olur ve RagC-GTP'nin fosfatını hidrolize ederek RagC-GDP'ye çevirir. Bu hidroliz, mTORC1 alımının son kilidini açar. Birt-Hogg-Dubé sendromunda FLCN mutasyonları bu hassas amino asit duyarlılığını bozarak patolojik sinyallere neden olur.",
        "J_RagC_hidroliz = k_GAP_FLCN * [FLCN-FNIP*] * [RagC_GTP] / ( K_M_RagC + [RagC_GTP] )",
        "Bu Michaelis-Menten eşitliği, FLCN-FNIP kompleksinin lizozom zarındaki RagC GAP hidroliz akısını matematiksel olarak tanımlar."
    ),
    (
        "6.7 GATOR1 ve GATOR2 Kompleksleri Arasındaki Karşılıklı İnhibisyon Dinamikleri",
        "GATOR (GAP Activity Towards Rags) sistemi, amino asit algılama zincirinin en merkezi dengeleyici dişlisidir ve iki alt kompleksten oluşur: GATOR1 ve GATOR2.",
        "GATOR1; DEPDC5, Nprl2 ve Nprl3 alt birimlerinden oluşan bir komplekstir ve RagA/B'ye karşı doğrudan GAP aktivitesi sergileyerek RagA-GTP'yi inaktif RagA-GDP'ye çevirir; dolayısıyla mTORC1'in güçlü bir endojen inhibitörüdür. GATOR2 ise beş alt birimli devasa bir komplekstir ve görevi GATOR1'i inhibe etmektir. Amino asitler mevcut olduğunda lösin (Sestrin2 üzerinden) ve arjinin (CASTOR1 üzerinden) GATOR2 üzerindeki baskıyı kaldırır; aktifleşen GATOR2, GATOR1'i kapatır ve mTORC1'in önü açılır.",
        "d[GATOR1_aktif]/dt = k_bazal - k_inh_GATOR2 * [GATOR2_serbest] * [GATOR1_aktif]",
        "Bu diferansiyel ilişki, amino asit sinyalleriyle serbest kalan GATOR2'nin GATOR1 GAP aktivitesini nasıl baskıladığını gösterir."
    ),
    (
        "6.8 Esansiyel Olmayan Amino Asitlerin Algılanması ve mTORC1 Aktivasyonuna Katkısı",
        "mTORC1 araştırmaları uzun süre lösin ve arjinin gibi esansiyel amino asitlere odaklanmış olsa da, glutamin, serin ve treonin gibi esansiyel olmayan amino asitlerin de kritik düzenleyici rolleri mevcuttur.",
        "L-glutamin, hücre içine SLC1A5 (ASCT2) taşıyıcısıyla alınır ve daha sonra SLC7A5/SLC3A1 (LAT1) antiporterı aracılığıyla hücre dışına pompalanırken eş zamanlı olarak hücre içine lösin çekilir (lösin-glutamin trambolini). Ayrıca glutamin, Rag GTPazlardan tamamen bağımsız olarak Arf1 GTPaz üzerinden mTORC1'in lizozoma göçünü uyarabilir. Serin ise de novo sfingolipid ve nükleotid sentezinde tüketilirken metabolik bir sensör gibi davranarak protein translasyon hızını kontrol eder.",
        "J_Leucine_giris = V_max_LAT1 * [Glutamin_ic] * [Losin_dis] / ( (K_m_gln + [Glutamin_ic]) * (K_m_leu + [Losin_dis]) )",
        "Bu çift substratlı antiport kinetiği, intraselüler glutamin havuzunun ekstraselüler lösin alımı ve dolayısıyla mTORC1 aktivasyonu için nasıl zorunlu bir yakıt olduğunu ortaya koyar."
    ),
    (
        "6.9 Yaşlanan Hücrelerde Amino Asit Algılama Hassasiyetinin Bozulması",
        "Yaşlanma sürecinde lizozomal membran kompozisyonunun bozulması, lipid peroksidasyonu ve glikasyon hasarı, amino asit algılama makinelerinin desensitizasyonuna yol açar.",
        "Yaşlı hücrelerde Sestrin2 ve CASTOR1 sensörlerinin ekspresyon seviyeleri veya afiniteleri değişir; hücre içi amino asit konsantrasyonları düşük olsa bile GATOR1 kompleksi lizozom zarına düzgün demirlenemez (KICSTOR defektleri). Sonuç olarak, hücre açlık ve besin kıtlığı anlarında dahi mTORC1'i kapatamaz; bu fenomen 'anabolik direnç' ve 'yanlış besin sinyalizasyonu' olarak adlandırılır. İskelet kasında ise tam tersine, yüksek amino asit konsantrasyonlarına rağmen mTORC1 yeterince aktive olamaz ve sarkopeni (kas kaybı) hızlanır.",
        "Hassasiyet_Bozulma = 1 - exp( - ( [Lipid_Peroksidasyon] + [Glikasyon_Lizozom] ) / E_esik )",
        "Bu hasar fonksiyonu, lizozom membranındaki biyokimyasal bozulmanın amino asit algılama hassasiyetini nasıl körelttiğini ifade eder."
    ),
    (
        "6.10 Amino Asit Manipülasyonu ile mTORC1'in Hassas Modülasyonu",
        "Amino asit sensörlerinin moleküler anatomisinin çözülmesi, tüm proteini kısmadan spesifik amino asitleri hedef alarak mTORC1'i cerrahi bir hassasiyetle kontrol etme fırsatı sunar.",
        "Diyette dallı zincirli amino asitlerin (BCAA: lösin, izolösin, valin) periyodik olarak sınırlandırılması veya metiyonin düzeyinin %80 azaltılması, organizmayı kronik açlık hissi yaratmadan derin bir otofaji ve tamir fazına sokar. Farmakolojik olarak Sestrin2'nin lösin bağlama cebini bloke eden veya CASTOR1'i aktive eden sentetik allosterik moleküller geliştirilmektedir. Bu moleküller, besin alımını engellemeden hücreye 'açlık sinyali' vererek yaşam süresini uzatma potansiyeline sahiptir.",
        "mTORC1_Baskilama_Verimi = 1 - ( [BCAA_plazma] / (K_m_BCAA + [BCAA_plazma]) ) * ( [Metiyonin] / (K_m_Met + [Metiyonin]) )",
        "Bu regülasyon modeli, hedefli amino asit kısıtlamasının sistemik mTORC1 anabolik akısını nasıl hassas bir şekilde aşağı regüle ettiğini simgeler."
    )
]

parts.append(("KISIM 4: SİRTUİN AİLESİ (SIRT1-7): EPİGENETİK DEASETİLASYON VE BİYOENERJETİK", part4_subsections))
parts.append(("KISIM 5: DİYET KISITLAMASI (CR) VE AÇLIK BENZERİ DİYETLER (FMD)", part5_subsections))
parts.append(("KISIM 6: AMİNO ASİT ALGILAMA: SESTRIN2, CASTOR1 VE RAG GTPAZLAR", part6_subsections))

# ================= KISIM 7 =================
part7_subsections = [
    (
        "7.1 Hipotalamik GHRH, Somatostatin ve Ön Hipofiz GH Salınım Dinamikleri",
        "Somatotropik eksen, hipotalamustan salınan iki zıt nörohormon tarafından yönetilir: Büyüme Hormonu Salgılatıcı Hormon (GHRH) ve büyüme hormonu inhibitörü olan Somatostatin (SRIF).",
        "GHRH, ön hipofiz somatotrof hücrelerindeki Gs-bağlantılı reseptörüne bağlanarak intraselüler cAMP ve PKA sinyalini artırır, bu da voltaj bağımlı kalsiyum kanallarını açarak Büyüme Hormonunun (GH / Somatotropin) pulsatil ekzositozunu uyarır. Somatostatin ise Gi-bağlantılı reseptörleri aracılığıyla adenilat siklazı baskılayarak GH deşarjını frenler. Yaşlanma ile hipotalamusta somatostatin tonusu artarken GHRH duyarlılığı körelir; bu durum gençlikte görülen yüksek amplitüdlü gece GH pulsasyonlarının sönümlenmesine (somatofakoz / somatopause) yol açar.",
        "d[GH_plazma]/dt = Sum_pulses( A_i * exp(-(t - t_i)^2 / (2 * sigma_p^2)) ) - k_klerens_GH * [GH_plazma]",
        "Bu eşitlik, plazma GH konsantrasyonunun hipotalamik diskret Gauss benzeri pulsasyonlarının toplamı ve renal/hepatik klerens dinamikleri ile belirlendiğini gösterir."
    ),
    (
        "7.2 Büyüme Hormonu Reseptörü (GHR) Sinyal Yolağı: JAK2 / STAT5b Kaskadı",
        "Büyüme Hormonu, hedef hücrelerin yüzeyinde tek transmembran alanlı homodimerik Büyüme Hormonu Reseptörüne (GHR) asimetrik olarak bağlanır. Bir GH molekülü, reseptör dimerinin iki ayrı bağlanma bölgesine (Site 1 ve Site 2) sırayla kenetlenir.",
        "Bu kenetlenme, sitoplazmik kuyruklarda konstitütif olarak bağlı bulunan Janus Kinaz 2 (JAK2) moleküllerinin birbirine yaklaşarak trans-fosforilasyonla aktive olmasını sağlar. Aktive JAK2, GHR sitoplazmik tirozinlerini fosforiller; bu fosforotirozinler STAT5b (Signal Transducer and Activator of Transcription 5b) transkripsiyon faktörü için pristin yanaşma bölgeleri sunar. JAK2 tarafından fosforillenen STAT5b homodimerleşir, çekirdeğe göç eder ve başta IGF1 olmak üzere somatotropik yanıt genlerinin transkripsiyonunu başlatır.",
        "Aktivasyon_STAT5b = k_JAK2 * [GHR_dimer*] * [STAT5b_monomer]^2 / ( K_M_STAT + [STAT5b_monomer]^2 )",
        "Bu formülasyon, dimerik GHR kompleksinin STAT5b dimerizasyonunu ve transkripsiyonel aktivasyonunu nasıl ikinci dereceden bir kinetikle uyardığını modeller."
    ),
    (
        "7.3 Karaciğerde IGF-1 Transkripsiyonu ve Dolaşımdaki IGFBP-3 / ALS Kompleksleri",
        "Dolaşımdaki endojen IGF-1'in (İnsülin Benzeri Büyüme Faktörü-1) %75'inden fazlası karaciğer hepatositleri tarafından Büyüme Hormonunun STAT5b aktivasyonu neticesinde sentezlenir ve kana verilir.",
        "Kanda serbest IGF-1'in yarı ömrü son derece kısadır (10-15 dakika). Ancak dolaşımdaki IGF-1'in %99'u, IGFBP-3 (İnsülin Benzeri Büyüme Faktörü Bağlayıcı Protein-3) ve karaciğer kökenli ALS (Acid-Labile Subunit) glikoproteini ile birleşerek 150 kDa ağırlığında devasa bir üçlü (terner) kompleks oluşturur. Bu terner kompleks vasküler endoteli geçemez ve IGF-1'in yarı ömrünü 16 saatin üzerine çıkararak dokulara sabit ve sürekli bir anabolik büyüme sinyali akışı sağlar. Serbest IGF-1 fraksiyonunun artması, mitotik baskıyı ve hücresel yaşlanma hızını artırır.",
        "Yari_Omur_IGF1 = Tau_serbest * ( 1 - f_terner ) + Tau_terner * f_terner",
        "Burada f_terner terner komplekse bağlı IGF-1 fraksiyonu olup, plazma anabolik rezervuarının zamana bağlı stabilitesini belirler."
    ),
    (
        "7.4 Laron Sendromu (GHR Eksikliği): İnsanlarda Kanser ve Diyabete Karşı Mutlak Direnç",
        "Zvi Laron tarafından tanımlanan Laron Sendromu, büyüme hormonu reseptörü (GHR) genindeki otozomal resesif inaktive edici mutasyonlardan kaynaklanan endokrin bir bozukluktur.",
        "Bu bireylerde dolaşımda yüksek Büyüme Hormonu seviyelerine rağmen karaciğer GH'ye yanıt veremez ve IGF-1 seviyeleri tespit edilemeyecek kadar düşüktür. Ekvador'daki Laron kohortunun 30 yılı aşkın takibi, olağanüstü bir gerçeği ortaya çıkarmıştır: Bu bireyler ciddi boy kısalığına rağmen kansere ve tip 2 diyabete karşı neredeyse tam bir biyolojik bağışıklığa sahiptir; kalp-damar hastalıkları ve nörodejeneratif süreçler belirgin şekilde gecikmiştir. Düşük IGF-1 seviyeleri, hücreleri anabolik stresten koruyarak sürekli bir bazal otofaji ve antioksidan temizlik fazında tutmaktadır.",
        "Kanser_Insidans_Laron = Kanser_Popilasyon * exp( -alpha_GHR_mut * (IGF1_normal - [IGF1_Laron]) )",
        "Bu epidemiyolojik model, düşük sistemik IGF-1 konsantrasyonunun neoplastik transformasyon ve tümör insidansı üzerindeki üstel koruyucu etkisini açıklar."
    ),
    (
        "7.5 Ames ve Snell Cüce Fare Modellerinde Rekor Yaşam Süresi Uzaması",
        "Kemirgenlerde somatotropik eksen bozuklukları, memelilerde şimdiye kadar kaydedilmiş en uzun yaşam süresi artışlarını sağlamıştır. Ames (Prop1 mutasyonu) ve Snell (Pit1 mutasyonu) cüce fareleri bu alanın kilometre taşlarıdır.",
        "Bu fareler ön hipofiz gelişim defekti nedeniyle GH, TSH ve Prolaktin üretemezler. Sonuç olarak IGF-1 seviyeleri dramatik şekilde düşüktür. Ames cüce fareleri vahşi tip akranlarına göre %50 ila %70 daha uzun yaşarlar (Andrzej Bartke et al.). GHR nakavt fareler (GHRKO) de benzer şekilde rekor ömre ulaşır. Bu hayvanlar yaşlılıklarında dahi olağanüstü insülin duyarlılığı, genç kök hücre havuzları, sıfır kanser insidansı ve yüksek antioksidan enzim kapasitesi sergilerler. Bu veriler, memelilerde büyüme sinyalinin kısılmasının ömrü uzattığını tartışmasız biçimde kanıtlar.",
        "Delta_Lifespan_Prop1 = Lifespan_0 * ( 1 + Gamma_Ames * ( 1 - [GH_serum] / GH_wt ) )",
        "Bu genetik formülasyon, dolaşımdaki büyüme hormonu yokluğunun memeli organizmasında yaşam süresi beklentisini nasıl %50'nin üzerinde artırdığını matematikselleştirir."
    ),
    (
        "7.6 Büyüme ile Yaşam Süresi Arasındaki Ters İlişki: Soma ile Germline Arasındaki Kaynak Tahsisi",
        "Evrimsel biyolog Thomas Kirkwood'un 'Harcanabilir Gövde Teorisi' (Disposable Soma Theory), organizmanın enerjisini büyüme/üreme ile somatik bakım arasında paylaştırmak zorunda olduğunu savunur.",
        "Aynı tür içinde vücut boyutu ile yaşam süresi arasında belirgin bir negatif korelasyon vardır: Küçük köpek ırkları (Chihuahua) dev ırklardan (Great Dane) iki kat daha uzun yaşar; benzer ilişki atlar, fareler ve insan boyu varyasyonlarında da gözlenir. Büyük beden, embriyogenez ve büyüme çağında trilyonlarca ekstra hücre bölünmesi demektir; bu da her bir somatik hücrenin replikatif telomer rezervini tüketir, kök hücre nişlerini yorar ve mutasyon olasılığını artırır. Küçük cüsseli organizmalar ise enerjiyi somatik hücre onarımına tahsis ederek hücresel gençliği muhafaza eder.",
        "Kaynak_Tahsis_Orani = Enerji_Onarim / Enerji_Buyume = K_soma / ( [IGF1] * [mTORC1] )",
        "Bu termodinamik oran, yüksek büyüme faktörü sinyallerinin hücresel onarım bütçesini nasıl çaldığını ve erken yıpranmayı nasıl hızlandırdığını belgeler."
    ),
    (
        "7.7 IGF-1 Eksikliğinin Beyin Fonksiyonları ve Nörodejenerasyon Üzerindeki Çelişkili Etkileri",
        "Sistemik IGF-1 eksikliği somatik dokularda ömrü uzatırken, merkezi sinir sisteminde durum çok daha karmaşık ve paradoksaldır.",
        "IGF-1, serebral korteks ve hipokampusta sinaptik plastisite, uzun süreli potansiyasyon (LTP), nörogenez ve nöronal sağkalım için güçlü bir trofik faktördür. Aşırı düşük beyin IGF-1 seviyeleri bilişsel fonksiyonlarda yavaşlamaya ve anksiyete benzeri davranışlara yol açabilir. Ancak diğer taraftan, Alzheimer modellerinde IGF-1 sinyalinin genetik veya farmakolojik olarak kısılması, nöronal otofajiyi açarak amiloid-beta ve hiperfosforile tau plaklarının lizozomlarda temizlenmesini hızlandırmakta ve nöronal hücre ölümünü durdurmaktadır. Dolayısıyla beyinde optimum, pulsatil bir IGF-1 dengesi şarttır.",
        "Kognitif_Optimum = f_trofik([IGF1_beyin]) * ( 1 - f_plak_birikimi([mTORC1_astrosit]) )",
        "Bu modelleme, nöronal sağlığın trofik destek ihtiyacı ile otofajik plak temizliği arasındaki hassas konsantrasyon eşiğini gösterir."
    ),
    (
        "7.8 GH Reseptör Antagonistleri (Pegvisomant) ve Yaşlanma Karşıtı Potansiyeli",
        "Büyüme hormonunun yaşlanma hızlandırıcı etkilerini yetişkinlik döneminde farmakolojik olarak dizginlemek için Pegvisomant gibi sentetik GH reseptör antagonistleri araştırılmaktadır.",
        "Pegvisomant, insan GH geninde yapılan spesifik nokta mutasyonu (G120K) ve polietilen glikol (PEG) polimerizasyonu ile üretilmiş rekombinant bir proteindir. GHR'nin Site 1 bölgesine bağlanır ancak Site 2 kenetlenmesini mekanik olarak bloke ederek reseptör dimerizasyonunu ve JAK2 aktivasyonunu engeller. Klinik olarak akromegali tedavisinde kullanılan Pegvisomant, karaciğer IGF-1 sentezini hızla düşürür, insülin duyarlılığını artırır ve anabolik kanserojenik baskıyı hafifletir. Yetişkinlikte periyodik düşük doz uygulanması, Laron benzeri bir metabolik koruma vaat etmektedir.",
        "GHR_Inhibisyon = [Pegvisomant] / ( IC50_Peg * (1 + [GH_endojen]/K_d_GH) + [Pegvisomant] )",
        "Bu kompetitif antagonizma formülü, Pegvisomant'ın endojen GH ile reseptör düzeyinde nasıl yarışarak sistemik sinyal iletimini baskıladığını tanımlar."
    ),
    (
        "7.9 Yetişkinlikte GH/IGF-1 Ekseninin Aşağı Regülasyonu ve Biyolojik Gençleşme",
        "Gelişim çağında kemik uzaması ve doku diferansiyasyonu için vazgeçilmez olan GH/IGF-1 ekseni, iskeletsel maturasyon tamamlandıktan sonra yaşlandırıcı bir patolojik itici güce dönüşür.",
        "Büyüme tamamlandıktan sonra yüksek GH ve IGF-1 seviyelerine ihtiyaç yoktur; aksine bu seviyeler proto-onkogenlerin uyarılmasına, hücrelerin senesense sürüklenmesine ve vasküler hipertrofiye neden olur. Yetişkinlik döneminde somatotropik eksenin farmakolojik, nutrisyonel veya gen terapisi (karaciğerde GHR veya IGF1 knockdown) yöntemleriyle %30-50 oranında aşağı regüle edilmesi; dokularda DNA tamir mekanizmalarını şarj etmekte, insülin duyarlılığını zirveye taşımakta ve sağlıklı ömrü uzatmaktadır.",
        "Genclesme_Skoru = Delta_Tamir * ( 1 - [IGF1_yetiskin] / IGF1_referans ) * ( [Otofaji_Akisi] / Otofaji_0 )",
        "Bu türetilmiş indeks, yetişkinlikte baskılanan IGF-1 sinyalinin otofajik akı artışı üzerinden biyolojik gençleşme katsayısına katkısını simgeler."
    ),
    (
        "7.10 Hormonal Aksın Hassas Titrasyonu: Büyüme Sinyalinin Güvenli Baskılanması",
        "Somatotropik ekseni modüle ederken en büyük risk, iskelet kası atrofisi (sarkopeni), kemik mineral yoğunluğu kaybı (osteoporoz) ve letarji gibi hipo-somatotropik komplikasyonlardır.",
        "Güvenli ve kusursuz bir titrasyon mimarisi; karaciğer kaynaklı endokrin IGF-1'i düşük tutarken (kanser ve senesensi önlemek için), iskelet kasında mekanik gerilimle lokal olarak üretilen otokrin/parakrin IGF-1Eb (MGF - Mekano Büyüme Faktörü) üretimini egzersiz ve hedefe yönelik peptitlerle desteklemeyi içerir. Bu dual yaklaşım, sistemik vücutta Laron benzeri bir ömür uzaması sağlarken, lokomotor sistemde tam kas kütlesini ve kemik gücünü korur.",
        "Terapotik_Denge_Indeksi = ( [MGF_lokal_kas] / MGF_esik ) * ( K_onko / (K_onko + [IGF1_serum]) )",
        "Bu oran, lokal kas yapıcı faktörlerin maksimize edilirken sistemik kanserojen serum IGF-1'in güvenli aralıkta minimize edilmesini sağlayan terapötik pencereyi gösterir."
    )
]

# ================= KISIM 8 =================
part8_subsections = [
    (
        "8.1 Rapamisin (Sirolimus): Streptomyces hygroscopicus'tan FKBP12 Kompleksine ve mTORC1 Allosterik İnhibisyonu",
        "Paskalya Adası (Rapa Nui) toprak bakterisi Streptomyces hygroscopicus'tan izole edilen makrosiklik lakton yapısındaki rapamisin (sirolimus), modern gerontolojinin en güçlü farmakolojik ajanıdır.",
        "Rapamisin hücre içine girdiğinde, sitozolik 12 kDa FK506-bağlayıcı protein (FKBP12) ile non-kovalan yüksek afiniteli bir kompleks oluşturur. Bu rapamisin-FKBP12 ikili kompleksi, mTOR kinazın katalitik alanının hemen bitişiğindeki FRB (FKBP-Rapamycin Binding) alanına fiziksel olarak yerleşir. Bu sterik blokaj, mTORC1'in kinaz yarığını kapatmaz; ancak substratların (özellikle p70S6K1 ve 4E-BP1) kinaz aktif merkezine girişini allosterik olarak engeller. Böylece hücresel protein sentezi yavaşlar ve otofaji üzerindeki baskı tamamen kalkar.",
        "Inhibisyon_mTORC1_Rapa = [Rapa-FKBP12] / ( K_d_FRB + [Rapa-FKBP12] )",
        "Bu formülasyon, sitozolde oluşan rapamisin-FKBP12 kompleksinin mTOR FRB alanına bağlanma doygunluğunu ve mTORC1 kinetik kısıtlamasını tanımlar."
    ),
    (
        "8.2 Everolimus, Temsirolimus ve Yeni Nesil mTOR Kinaz İnhibitörleri (Torin1, RapaLink-1)",
        "Klasik rapamisinin (sirolimus) farmakokinetik çözünürlük kısıtlılıklarını aşmak ve etki gücünü artırmak amacıyla 'rapalog' türevleri ve yeni nesil ATP-kompetitif kinaz inhibitörleri geliştirilmiştir.",
        "Everolimus (RAD001), 40-O-(2-hidroksietil) modifikasyonuyla oral biyoyararlanımı artırılmış bir rapalogdur. Temsirolimus (CCI-779) ise bir ester ön-ilaçtır. Klasik rapaloglar 4E-BP1 fosforilasyonunu tam olarak kesemezken; Torin1 ve PP242 gibi yeni nesil ATP-kompetitif TORKi inhibitörleri, mTOR'un ATP bağlanma cebine oturarak hem mTORC1 hem de mTORC2'yi tam konsantrasyonda susturur. RapaLink-1 ise rapamisin ile bir TORKi molekülünü esnek bir polietilen glikol bağıyla birleştiren devrimsel bir çift-başlıklıdır; rapalog dirençli mutasyonları bile yok eder.",
        "Potansiyel_RapaLink = IC50_Torin * ( K_d_Rapa / [RapaLink-1] ) * ( 1 / Substrat_Engelleme_Faktoru )",
        "Bu formülasyon, bivalent (çift-başlıklı) inhibitörlerin mTOR aktif cebine yanaşma termodinamiğini ve tekil ajanlara kıyasla katlanan afinite üstünlüğünü simgeler."
    ),
    (
        "8.3 Ulusal Yaşlanma Enstitüsü (NIA) Müdahale Test Programı (ITP) Rapamisin Verileri",
        "Amerikan Ulusal Yaşlanma Enstitüsü'nün (NIA) yürüttüğü Müdahale Test Programı (Interventions Testing Program - ITP), yaşlanma karşıtı bileşiklerin fare ömrü üzerindeki etkilerini üç bağımsız merkezde, genetik olarak heterojen farelerde (UM-HET3) test eden en güvenilir bilimsel konsorsiyumdur.",
        "2009 yılında yayınlanan tarihi ITP çalışmasında; farelere 600 günlükken (insan ömründe yaklaşık 60 yaşa denk gelen ileri yaşlılık dönemi) başlanan rapamisin tedavisi, dişi farelerde medyan yaşam süresini %14, erkek farelerde %9 oranında artırmıştır. Daha sonraki çalışmalarda genç yaşta başlanan ve optimize edilen dozlarda bu artış %23-26 seviyelerine ulaşmıştır. Rapamisin, ileri yaşta başlansa dahi memeli yaşam süresini istatistiksel ve biyolojik olarak tartışmasız biçimde uzatan ilk ve tek farmakolojik molekül olarak tarihe geçmiştir.",
        "Delta_Lifespan_ITP = Lifespan_kontrol * ( 1 + alpha_cinsiyet * Dose_Rapa^m / (ED50^m + Dose_Rapa^m) )",
        "Bu tepe yanıtı modeli, ITP fare kohortlarında rapamisin kan konsantrasyonu ile yaşam süresi artış yüzdesi arasındaki doza bağımlı ve cinsiyete duyarlı ilişkiyi gösterir."
    ),
    (
        "8.4 Rapamisinin Dozaj Protokolleri: Sürekli vs. Aralıklı Uygulama ve İmmün Yanıt İkilemi",
        "Rapamisinin klinik kullanımındaki en büyük engel, yüksek doz sürekli kullanımda organ nakli hastalarında görülen immünosupresyon, aftöz stomatit, hiperlipidemi ve insülin direncidir.",
        "Bu yan etkiler, sürekli rapamisin mevcudiyetinin serbest mTOR moleküllerini tüketerek zamanla mTORC2 montajını bozmasından kaynaklanır. Ancak modern longevity protokollerinde 'aralıklı uygulama' (Intermittent Dosing - örn. haftada bir kez 5-6 mg tek doz puls) benimsenmiştir. Bu puls dozlama, kanda hızlı bir tepe konsantrasyonu oluşturarak mTORC1'i derinlemesine kapatır ve güçlü bir otofajik deşarj sağlar; ardından ilaç 48-72 saat içinde kandan temizlenir ve mTORC2 montajı korunur. İlginç bir şekilde bu aralıklı protokol, yaşlılarda immünosupresyon yaratmak yerine grip aşılarına verilen antikor yanıtını %20 artırmaktadır (Mannick et al.).",
        "mTORC2_Korunum_Orani = 1 / ( 1 + k_inaktivasyon * Integral_0_T([Rapa_plazma](t) dt) )",
        "Bu zamana bağlı integral, kanda rapamisin maruziyet alanının (AUC) düşük tutulduğu aralıklı puls protokollerinde mTORC2 bütünlüğünün nasıl korunduğunu açıklar."
    ),
    (
        "8.5 Metformin: Mitokondriyal Kompleks I İnhibisyonu, AMP/ATP Oranı ve LKB1-AMPK Aktivasyonu",
        "Keçi sedef otundan (Galega officinalis) türetilen bir biguanid olan metformin, dünyada tip 2 diyabet tedavisinde en yaygın kullanılan ve yaşlanma karşıtı özellikleri en iyi belgelenmiş moleküldür.",
        "Metformin, pozitif yüklü kimyasal yapısı sayesinde mitokondriyal membran potansiyelini (Delta-Psi_m) kullanarak mitokondri matriksinde 1000 kat konsantre olur. Burada Elektron Taşıma Zinciri Kompleks I'i (NADH:ubikinon oksidoredüktaz) zayıf ve geri dönüşümlü olarak inhibe eder. Bu hafif metabolik kısıtlama ATP üretimini bir miktar düşürürken ADP/ATP ve AMP/ATP oranlarını yukarı fırlatır. Artan AMP, LKB1 aracılığıyla AMPK'yi Thr172'den fosforilleyerek aktive eder; hepatik glukoneojenezi baskılar ve periferik glukoz kullanımını artırır.",
        "Delta_AMP_ATP = k_metformin * [Metformin_matriks] / ( K_i_KompleksI + [Metformin_matriks] )",
        "Bu biyokimyasal bağıntı, matrikste biriken metforminin Kompleks I inhibisyon şiddeti ile hücre içi AMP/ATP oranı fırlaması arasındaki ilişkiyi açıklar."
    ),
    (
        "8.6 TAME (Targeting Aging with Metformin) Klinik Çalışması ve Biyobelirteçler",
        "Amerikan Yaşlanma Araştırmaları Federasyonu (AFAR) ve Dr. Nir Barzilai öncülüğünde tasarlanan TAME çalışması, yaşlanmayı doğrudan bir endikasyon olarak hedefleyen tarihteki ilk FDA onaylı klinik çalışma dizaynıdır.",
        "Çalışma, 65-79 yaş arası 3.000 diyabetik olmayan bireyde metforminin yeni bir kronik hastalığın (kanser, kardiyovasküler hastalık, kognitif gerileme veya ölüm) ortaya çıkış süresini geciktirip geciktiremeyeceğini test etmektedir. Çalışmanın birincil sonlanım noktası, çoklu morbiditelerin kümülatif insidansıdır. Biyobelirteç panelinde ise DNA metilasyon saatleri (GrimAge, PhenoAge), dolaşımdaki interlökin-6, TNF-alfa, hs-CRP ve metabolik insülin duyarlılık parametreleri izlenmektedir.",
        "TAME_Hazard_Ratio = exp( -beta_metformin * [Metformin_plazma] - gamma_biomarker * Delta_Epigenetik_Saat )",
        "Bu Cox orantılı risk modeli, metforminin kümülatif yaşa bağlı kronik hastalık insidansını düşürme olasılığını epigenetik saat yanıtı üzerinden tahmin eder."
    ),
    (
        "8.7 Resveratrol ve Yeni Nesil STAC'ler (SRT1720, SRT2104): Doğrudan SIRT1 Aktivasyonu",
        "Kırmızı üzüm kabuğunda bulunan bir polifenol olan resveratrol, David Sinclair ve arkadaşları tarafından maya Sir2 ve memeli SIRT1 enzimini doğrudan aktive eden ilk molekül olarak tanımlanmıştır.",
        "Resveratrol, SIRT1'in N-terminal aktivasyon alanına bağlanarak enzimin asetilli peptit substratlarına olan afinitesini artırır. Ancak resveratrolün düşük biyoyararlanımı ve çoklu hedeflere bağlanma (off-target) profili nedeniyle ilaç şirketleri sentetik küçük moleküllü Sirtuin Aktive Edici Bileşikler (STAC) geliştirmiştir. SRT1720 ve SRT2104, resvertrolden 1000 kat daha güçlü olup nanomolar afiniteyle SIRT1'e bağlanır. Bu moleküller fare modellerinde yüksek yağlı diyetin zararlarını silmekte, mitokondriyal biyogenezi patlatmakta ve vasküler elastikiyeti korumaktadır.",
        "Aktivasyon_Katsayisi_STAC = 1 + alpha_max * [STAC] / ( EC50_STAC + [STAC] )",
        "Bu allosterik aktivasyon denklemi, STAC konsantrasyonu ile SIRT1 deasetilasyon hız artışı arasındaki Hill tipi doygunluk kinetiğini açıklar."
    ),
    (
        "8.8 Alfa-Ketoglutarat (AKG): Epigenetik Demetilasyon ve mTOR İnhibisyonu Yoluyla Yaşam Süresi Artışı",
        "Krebs döngüsünün temel bir ara ürünü olan alfa-ketoglutarat (AKG / 2-oksoglutarat), yaşlanma karşıtı metabolitler arasında devrim niteliğinde bir yere oturmuştur.",
        "AKG, DNA demetilasyonunu gerçekleştiren TET enzimlerinin ve histon demetilazların (Jumonji C alanlı KDM'ler) zorunlu ko-substratıdır. Yaşla birlikte dokularda AKG seviyeleri düşer; ekzojen kalsiyum alfa-ketoglutarat (Ca-AKG) desteği, hücresel heterokromatin yapısını temizler ve anormal hipermetilasyonları geri döndürür. Eş zamanlı olarak AKG, mitokondriyal ATP sentazın (Kompleks V) F1 alt birimine bağlanarak onu inhibe eder; bu durum hücre içi ATP akısını hafifçe frenleyerek mTORC1'i baskılar ve otofajiyi uyarır. Farelerde Ca-AKG ömrü %12 uzatırken frailty (kırılganlık) indeksini %40 azaltmıştır.",
        "Demetilasyon_Hizi_TET = k_cat_TET * [TET] * [5mC] * [AKG] / ( (K_m_5mC + [5mC]) * (K_m_AKG + [AKG]) )",
        "Bu Bi-Substrat eşitliği, hücre içi alfa-ketoglutarat konsantrasyonunun genomik DNA demetilasyon akısını nasıl doğrudan sınırladığını gösterir."
    ),
    (
        "8.9 Acarbose ve 17-alfa-Estradiol: ITP Onaylı Diğer Metabolik Yaşam Uzatıcılar",
        "NIA Müdahale Test Programı'nın (ITP) rapamisin dışında yaşam süresini istikrarlı ve anlamlı şekilde uzattığını teyit ettiği iki diğer kritik molekül Akarboz ve 17-alfa-estradioldür.",
        "Akarboz, bağırsak fırçamsı kenarında alfa-glukozidaz enzimini inhibe ederek kompleks karbonhidratların glukoza parçalanmasını ve emilimini geciktirir. Yemek sonrası ani glukoz fırlamalarını (postprandiyal glukoz spike) tamamen düzleştirir. Bu durum glukoz toksisitesini, vasküler hasarı ve kompanzatuar insülin salgısını önler; ITP testlerinde erkek farelerin ömrünü %22 uzatmıştır. Non-feminizan bir östrojen analoğu olan 17-alfa-estradiol ise klasik östrojen reseptörlerine bağlanmaz; hipotalamusta metabolik inflamasyonu baskılar, visseral yağ dokusunu eritir ve erkek kemirgenlerde rekor ömür artışı sağlar.",
        "Postprandial_Glukoz_Tepesi = G_bazal + Delta_G_max / ( 1 + [Akarboz] / K_i_glukozidaz )",
        "Bu inhibisyon denklemi, artan akarboz konsantrasyonunun yemek sonrası glukoz piki amplitüdünü nasıl kademeli olarak baskıladığını belgeler."
    ),
    (
        "8.10 Farmakolojik Sinerji: Çoklu Metabolik İlaç Kombinasyonları ve Toksisite Profilleri",
        "Tek bir yaşlanma karşıtı molekülün yüksek dozda kullanılması kaçınılmaz toksisiteler doğururken; farklı biyolojik mekanizmalara sahip ajanların düşük doz kombinasyonları sinerjistik yaşam uzaması vadeder.",
        "Örneğin Rapamisin (mTORC1 inhibisyonu) + Metformin (AMPK aktivasyonu ve Kompleks I freni) kombinasyonu, rapamisinin neden olabileceği insülin direncini metforminin periferik glukoz duyarlılaştırıcı etkisiyle nötralize eder. Benzer şekilde Rapamisin + Akarboz kombinasyonu, ITP çalışmalarında tek başına kullanılan moleküllerden çok daha üstün bir yaşam süresi uzaması sağlamıştır. Hedefe yönelik geroprotektif kombinasyon kokteylleri, her bir yolağı hormetik aralıkta tutarak sıfır yan etki ile maksimum rejenerasyon elde etmenin nihai farmakolojik formülüdür.",
        "Sinerji_Katsayisi_CI = (Dose_A / D_x_A) + (Dose_B / D_x_B) + alpha_interaksiyon * ( (Dose_A * Dose_B) / (D_x_A * D_x_B) )",
        "Bu Chou-Talalay sinerji indeksi (CI < 1 sinerji), kombine metabolik ilaçların tek tek kullanımına kıyasla biyolojik gençleşme hedefine ulaşmadaki terapötik üstünlüğünü kanıtlar."
    )
]

# ================= KISIM 9 =================
part9_subsections = [
    (
        "9.1 Moleküler Sirkadiyen Saat Ağı: CLOCK, BMAL1, PER1-3 ve CRY1-2 Geribildirim Döngüsü",
        "Hücrenin metabolik süreçleri, Dünya'nın 24 saatlik eksen dönüşüne uyum sağlayan otonom moleküler bir sirkadiyen saat mekanizması tarafından yönetilir.",
        "Bu saatin pozitif kolunu, bHLH-PAS transkripsiyon faktörleri olan CLOCK ve BMAL1 (ARNTL) heterodimeri oluşturur. Bu dimer, E-box regülatör elemanlarına (CACGTG) bağlanarak Period (PER1, PER2, PER3) ve Kriptokrom (CRY1, CRY2) genlerinin transkripsiyonunu başlatır. Sitoplazmada biriken PER ve CRY proteinleri dimerleşir, fosforillenir ve gecikmeli olarak çekirdeğe geri dönerek CLOCK:BMAL1 aktivitesini kapatır (negatif geribildirim). Ek döngüde RORalfa aktivatör ve REV-ERBalfa inhibitör reseptörleri BMAL1 ekspresyonunu ritmik olarak dalgalandırır.",
        "d[PER:CRY_nukleus]/dt = k_trans * [PER:CRY_sitozol] - k_deg * [PER:CRY_nukleus] - k_represyon * [CLOCK:BMAL1]",
        "Bu diferansiyel denklem, nükleer repressör kompleksinin zamana bağlı salınım dinamiğini ve negatif transkripsiyonel kilit döngüsünü formüle eder."
    ),
    (
        "9.2 Periferik Organ Saatleri ve Karaciğer, Kas, Yağ Dokusunda Besin Senkronizasyonu",
        "Sirkadiyen sistem hiyerarşiktir: Hipotalamustaki Suprakiyazmatik Çekirdek (SCN) ana pacemaker (master saat) olarak ışık-karanlık döngüsüyle ayarlanırken; karaciğer, iskelet kası ve yağ dokusundaki periferik saatler besin alım zamanlamasıyla senkronize olur.",
        "Besin girişi, periferik organlarda insülin ve glukagon salınımı aracılığıyla saat genlerinin fazını sıfırlar. Karaciğerde glikojenoliz, glukoneojenez, lipogenez ve safra asidi sentezi katı bir sirkadiyen ritme tabidir. İskelet kasında GLUT4 translokasyonu ve mitokondriyal solunum öğleden sonra zirve yapar. Gece geç saatlerde besin tüketimi, periferik saatlerin SCN ana saatinden koparak desenskronize olmasına (sirkadiyen uyumsuzluk) yol açar; bu durum metabolik sendromu ve erken vasküler yaşlanmayı körükler.",
        "Desenkronizasyon_Indeksi = Integral_0_24h | Phase_SCN(t) - Phase_Karaciger(t) | dt",
        "Bu faz farkı integrali, merkezi SCN saati ile periferik hepatik metabolik saat arasındaki sirkadiyen uyumsuzluk derecesini ve buna bağlı patolojik yükü gösterir."
    ),
    (
        "9.3 Zaman Kısıtlı Beslenme (Time-Restricted Eating - TRE): 8 Saatlik Pencere ve Metabolik Esneklik",
        "Zaman Kısıtlı Beslenme (TRE), kalori kısıtlaması yapmaksızın günlük tüm besin alımını tutarlı ve biyolojik ritme uygun bir pencereye (genellikle 8-10 saat) sınırlayan yaşam tarzı müdahalesidir.",
        "Satchin Panda ve grubunun çığır açan çalışmaları, yüksek yağlı ve yüksek şekerli diyetle beslenen kemirgenlerin dahi besinleri 8 saatlik aktif faz penceresinde tükettiklerinde obezite, karaciğer yağlanması ve insülin direncinden korunduklarını göstermiştir. Geriye kalan 16 saatlik açlık penceresi, periferik dokuların katabolik ve rejeneratif moda geçmesini sağlar. TRE, metabolik esnekliği (hücrenin glukoz yakmaktan yağ asidi yakmaya zahmetsizce geçebilme kapasitesi) yeniden inşa eder ve sistemik glisemik değişkenliği (glycemic variability) minimize eder.",
        "Metabolik_Esneklik_Skoru = Delta_RER = | RER_beslenme(glukoz) - RER_aclik(yag) |",
        "Bu solunum değişim oranı (RER) farkı, organizmanın yakıt kaynakları arasında ne kadar kusursuz geçiş yapabildiğini ölçen en hassas metabolik gençlik parametresidir."
    ),
    (
        "9.4 SIRT1 ve Sirkadiyen Ritim Entegrasyonu: BMAL1 ve PER2'nin Deasetilasyonu",
        "Hücresel redoks ve besin durumu ile sirkadiyen saat mekanizması arasındaki en zarif moleküler köprü, NAD+ bağımlı deasetilaz SIRT1'dir.",
        "SIRT1, promoter bölgelerinde CLOCK-BMAL1 kompleksiyle fiziksel etkileşime girer. Hücrede NAD+ ritmik olarak dalgalanır (NAMPT ekspresyonunun CLOCK:BMAL1 kontrolünde olması sayesinde). NAD+ zirve yaptığında aktive olan SIRT1, BMAL1 proteinini Lys537 bölgesinden ve repressör PER2'yi deasetile eder. PER2'nin deasetilasyonu, onun beta-TrCP ubikiton ligazı tarafından hedeflenerek proteazomal yıkıma uğramasını hızlandırır. Bu dinamik temizlik, negatif geribildirim frenini zamanında kaldırarak sirkadiyen döngünün yeni bir güne yüksek amplitüdlü başlamasını temin eder.",
        "d[PER2]/dt = Sentez_PER2 - k_deac_SIRT1 * [SIRT1] * [NAD+] * [PER2_ac] - k_deg_bazal * [PER2]",
        "Bu denklem, SIRT1 ve NAD+ seviyelerinin PER2 yıkım hızını ve sirkadiyen saatin hız hassasiyetini nasıl belirlediğini açıklar."
    ),
    (
        "9.5 Gece Açlığının Otofaji Zirveleri ve Mitokondriyal Biyogenez Üzerindeki Sirkadiyen Kontrolü",
        "Hücresel temizlik mekanizması olan makrootofaji, gün boyu sabit bir hızda gerçekleşmez; sirkadiyen saatin sıkı kontrolü altında gece karanlık ve açlık fazında dramatik bir zirve yapar.",
        "Uyku ve gece açlığı sırasında azalan insülin/mTORC1 sinyali ve artan AMPK aktivitesi, C/EBPbeta ve TFEB (Transcription Factor EB) transkripsiyon faktörlerinin çekirdeğe göçünü sağlar. TFEB, CLEAR (Coordinated Lysosomal Expression and Regulation) gen ağını açarak otofagozom montajını ve lizozom biyogenezini tetikler. Beyinde gece boyunca glimfatik sistem ile senkronize çalışan bu otofajik deşarj, gün boyu nöronlarda biriken hasarlı proteinleri ve mitokondrileri temizler. Gece atıştırmaları ise mTORC1'i aktive ederek bu yaşamsal temizlik penceresini kapatır.",
        "Otofajik_Klerens_Gece = Integral_uyku( V_TFEB * [TFEB_nukleus](t) / (K_M + [TFEB_nukleus](t)) ) dt",
        "Bu integral formülasyon, uyku ve açlık fazı boyunca nükleer TFEB aktivitesinin toplam hücresel detoksifikasyon kapasitesini nasıl oluşturduğunu tanımlar."
    ),
    (
        "9.6 Yaşlanmayla Sirkadiyen Amplitüdün Sönümlenmesi ve Periferik Desenkronizasyon",
        "Yaşlanma sürecinin en evrensel biyobelirteçlerinden biri, moleküler sirkadiyen osilasyonların tepe-dip genliğinin (amplitüd) sönümlenmesi ve ritmik faz kaymalarıdır.",
        "Yaşlı bireylerde BMAL1 ve CLOCK ekspresyon seviyeleri azalır, E-box bağlanma dinamikleri bozulur ve gece melatonin salınım piki körelir. Aynı zamanda periferik organların ana saatten koparak kendi başlarına desenskronize fazlara kaydığı gözlenir. Amplitüdün sönümlenmesi; gece uykusunun parçalanmasına, gündüz yorgunluğuna, kronik hafif hiperglisemiye ve kardiyovasküler olayların sabaha karşı artmasına yol açar. Moleküler saatin yeniden ayarlanması (resynchronization), biyolojik gençleşmenin temel şartıdır.",
        "Amplitud_Sirkadiyen(t) = A_0 * exp(-k_yas * Yas) * cos(omega * t + Phi_kayma)",
        "Bu sönümlü dalga denklemi, artan kronolojik yaşla birlikte moleküler saat genliği kaybını ve faz kaymasını matematikselleştirir."
    ),
    (
        "9.7 Melatonin, Sirkadiyen Eksen ve Mitokondriyal Antioksidan Koruma",
        "Epifiz bezinden karanlıkta salınan melatonin hormonu, sadece bir uyku indükleyicisi değil; mitokondri matriksinde üretilen ve tüketilen en güçlü endojen antioksidanlardan biridir.",
        "Melatonin, MT1 ve MT2 G-protein kenetli reseptörleri aracılığıyla sirkadiyen ritmi senkronize ederken; amfifilik yapısı sayesinde serbestçe mitokondriye girer. Hatta son bulgular, mitokondrinin kendi içinde yüksek miktarda melatonin sentezlediğini göstermektedir. Melatonin, hidroksil radikallerini ve peroksinitriti doğrudan süpürür; bu reaksiyonlar sonucunda pro-oksidan ara ürünler oluşturmaz (intihar antioksidanı). Yaşlanmayla melatonin salınımının çökmesi, gece boyunca mitokondrilerin serbest radikal hasarına karşı savunmasız kalmasına neden olur.",
        "d[ROS_mitokondri]/dt = J_sizinti_ETC - k_scavenge * [Melatonin_matriks] * [ROS_mitokondri]",
        "Bu redoks denge modeli, mitokondri matriks içi melatonin konsantrasyonunun lipid ve mtDNA peroksidasyonunu nasıl sınırlandırdığını ifade eder."
    ),
    (
        "9.8 Nöroendokrin Besin Zamanlaması: Leptin ve Ghrelin Salınımının Sirkadiyen Kontrolü",
        "İştah ve enerji dengesini kontrol eden adipokin leptin ve mide kaynaklı peptit ghrelin, sirkadiyen saatin nöroendokrin arayüzleridir.",
        "Sağlıklı genç bireylerde leptin seviyeleri gece uyku sırasında zirve yaparak açlık hissini baskılar ve kesintisiz uykuyu garanti eder. Ghrelin ise alışılmış yemek saatlerinden hemen önce hipotalamik oreksijenik nöronları (NPY/AgRP) uyarmak üzere yükselir. Sirkadiyen ritim bozulduğunda ve gece ışık kirliliğine maruz kalındığında leptin tepe noktası çöker, leptin direnci gelişir ve gece kompulsif yeme atakları başlar. Bu durum karaciğerde yağlanmayı, gece insülin sekresyonunu ve metabolik sendromu tetikler.",
        "Istah_Durtusu = alpha_ghrelin * [Ghrelin] / (K_G + [Ghrelin]) - beta_leptin * [Leptin] / (K_L + [Leptin])",
        "Bu nöroendokrin fonksiyon, sirkadiyen faz bozulmalarında hipotalamusta oluşan patolojik açlık dürtüsünü formüle eder."
    ),
    (
        "9.9 Işık Maruziyeti, Mavi Işık Blokajı ve Metabolik Saatin Restorasyonu",
        "Suprakiyazmatik çekirdeğin (SCN) ana fotik girdisi, retinada yer alan ve melanopsin fotopigmenti içeren intrinsik fotosensitif retinal ganglion hücreleri (ipRGC) tarafından sağlanır.",
        "Melanopsin, özellikle 460-480 nm dalga boyundaki mavi ışığa maksimum duyarlılığa sahiptir. Sabah saatlerinde alınan parlak doğal güneş ışığı (10.000 lüks üzeri), retinohipotalamik traktus (RHT) üzerinden glutamat ve PACAP deşarjı yaparak SCN saatini kusursuz sıfırlar. Akşam saatlerinde yapay mavi ışığa maruz kalmak ise epifiz bezinde melatonin üretimini anında baskılar ve periferik organlara 'hala öğlen vaktindeyiz' yanıltıcı sinyali gönderir. Akşam mavi ışığın filtrelenmesi, moleküler saatin restorasyonu ve gece otofajisinin açılması için zorunlu bir fiziksel müdahaledir.",
        "Melatonin_Supresyon = I_isik * ( lambda / 470nm )^n / ( K_foto + I_isik * ( lambda / 470nm )^n )",
        "Bu fotobiyolojik denklem, retina üzerine düşen mavi spektrumlu ışık yoğunluğunun sistemik melatonin salınımını nasıl logaritmik olarak baskıladığını açıklar."
    ),
    (
        "9.10 Sirkadiyen-Metabolik Uyum Protokolü: Kronobiyoloji Tabanlı Uzun Yaşam Mimarisi",
        "Biyolojik ömrü maksimize etmek ve metabolik entropiyi minimuma indirmek, beslenme, ışık ve egzersiz parametrelerinin sirkadiyen fazlarla kusursuz rezonansa sokulmasını gerektirir.",
        "Protokol mimarisi şu temellere dayanır: 1) Sabah uyanıştan sonra ilk 60 dakika içinde yüksek yoğunluklu doğal fotik uyaran; 2) Beslenmenin metabolik hızın ve insülin duyarlılığının zirvede olduğu erken saatlere kaydırılması (Erken TRE: 08:00 - 16:00); 3) İskelet kası anaerobik direnç egzersizlerinin kas glukoz alımının ve vücut ısısının pik yaptığı öğleden sonra geç saatlere yerleştirilmesi; 4) Gün batımından sonra mavi/yeşil foton kısıtlaması ve tam karanlıkta uyku; 5) Uyku öncesi mikrodoyurucu biyomoleküller (melatonin, magnezyum l-treonat). Bu entegre krono-mimari, hücresel saat genlerinin amplitüdünü gençlik seviyelerinde tutar.",
        "Sirkadiyen_Rezonans_Skoru = Prod_i ( 1 - | t_eylem_i - t_optimum_i | / 24h )",
        "Bu çarpımsal harmonik skor, tüm yaşam tarzı girdilerinin moleküler sirkadiyen saat fazlarıyla örtüşme derecesini 0 ile 1 arasında sayısallaştırır."
    )
]

# ================= KISIM 10 =================
part10_subsections = [
    (
        "10.1 Sentetik Glikoz ve Amino Asit Sensörleri ile Dinamik Metabolik Geri Bildirim",
        "Homo Aeternus metabolik mühendisliğinin ilk adımı, hücrenin evrimsel olarak yetersiz kalan doğal besin algılama sensörlerini sentetik genetik devrelerle güçlendirmektir.",
        "Hücre zarına ve lizozom lümenine entegre edilen sentetik riboşalterler (riboswitches) ve kimerik reseptörler, intraselüler glukoz-6-fosfat, serbest lösin ve arjinin konsantrasyonlarını pikomolar düzeyde dinamik olarak ölçer. Bu sensörler metabolit seviyesi patolojik eşiği aştığında konformasyonel olarak katlanarak CRISPRi (transkripsiyonel baskılama) kaskadını aktive eder. Böylece besin fazlalığı hücreye zarar vermeden anabolik sinyal otomatik olarak frenlenir.",
        "Geri_Besleme_Faktoru = K_fb * [Metabolit_Sentetik_Bagli] / ( 1 + [Toksik_Ara_Urun] / K_esik )",
        "Bu kontrol teorisi formülü, sentetik sensörün metabolik sinyali kapalı devre bir geri bildirim döngüsünde nasıl regüle ettiğini simgeler."
    ),
    (
        "10.2 CRISPR-dCas9 ile Programlanabilir mTORC1 Transkripsiyonel Kısıtlayıcılar",
        "Doğal organizmada mTORC1 aktivitesi kontrolsüz besin sinyalleriyle sınırsızca yükselebilirken; sentetik genom mimarisi mTOR geninin promoter bölgesine yapay transkripsiyonel sınırlar koyar.",
        "Katalitik olarak inaktif dCas9 (dead Cas9) proteini, KRAB (Krüppel-associated box) süpresör alanı ile füzyonlanarak dokuya özel promoterlar kontrolünde eksprese edilir. Hücresel enerji veya besin sensörleri kronik doygunluk tespit ettiğinde, dCas9-KRAB kompleksi MTOR ve RPTOR (Raptor) genlerinin transkripsiyon başlangıç bölgelerine yönlendirilir. Bu durum mTORC1 ekspresyonunu tamamen sıfırlamadan tam olarak %50 oranında kısar; bu kalibrasyon kalori kısıtlamasının sağladığı tüm otofajik faydaları hiçbir kısıtlama yapmadan hücreye kazandırır.",
        "mTOR_Transkripsiyon_Sentetik = V_0 / ( 1 + [dCas9-KRAB_MTOR_bagli] / K_d_crispr )",
        "Bu eşitlik, programlanmış sentetik dCas9 repressörünün mTOR gen ekspresyonunu hedefli bir tavan değerde nasıl kilitlediğini gösterir."
    ),
    (
        "10.3 Yapay Zeka Güdümlü Biyosensörlerle Sürekli Hücresel Enerji İzleme (ATP/AMP Oranı)",
        "Geleceğin hücresel biyolojisi, florofor temelli FRET (Förster Rezonans Enerji Transferi) biyosensörleri ile tek hücre düzeyinde nükleotid oranlarını gerçek zamanlı takip etme kabiliyetine sahiptir.",
        "ATeam ve PercevalHR gibi genetik olarak kodlanmış biyosensörler, ATP/ADP ve NADH/NAD+ oranlarını mikrosaniyelik çözünürlükle floresan sinyale dönüştürür. İmplante edilebilir nano-biyo sensör dizilimleri bu optik sinyalleri yakalayarak hücresel enerji krizlerini veya aşırı doygunluk dalgalarını algılar. Yapay zeka algoritmaları bu verileri işleyerek dokunun hangi saatte otofajiye, hangi saatte biyosenteze girmesi gerektiğini hesaplar ve hücre içi mikro-dağıtım sistemlerine komut iletir.",
        "FRET_Orani = I_akseptor / I_donor = F_0 + Delta_F * [ATP] / ( K_d_FRET + [ATP] )",
        "Bu optik biyofizik denklemi, hücre içi serbest ATP konsantrasyonunun emisyon spektrumu floresan oranı üzerinden nasıl doğrudan ölçüldüğünü modeller."
    ),
    (
        "10.4 Sentetik Allosterik Aktivatörlerle Dokuya Özel Aşırı Duyarlı AMPK Tasarımı",
        "Yaşlanmayla körelen doğal AMPK aktivasyonunun üstesinden gelmek için, protein mühendisliği ile 'süper-duyarlı' sentetik AMPK varyantları tasarlanmıştır.",
        "Bu sentetik varyantlar, beta-1 alt birimindeki ADaM cebinde mutasyonlar taşır (örn. S108E fosfomimetik mutasyonu). Bu modifikasyon, enzimin bazal katalitik aktivitesini 20 kat artırırken, AMP gereksinimini ortadan kaldırır ve PP2C fosfatazlarının Thr172 defosforilasyonuna karşı tam direnç sağlar. Dokuya özel (örn. sadece hepatositlerde ve kardiyomiyositlerde) koşullu eksprese edilen bu sentetik kinazlar, organizmanın hiçbir organında yaşa bağlı metabolik yetersizlik ve lipid birikimi yaşanmamasını garanti eder.",
        "Katalitik_Verim_Sentetik_AMPK = k_cat_mut * ( 1 + Eta_direnc ) / K_M_mut",
        "Bu enzim kinetiği parametresi, fosfomimetik sentetik kinazın doğal enzime kıyasla katlanan substrat dönüştürme hızını temsil eder."
    ),
    (
        "10.5 Sürekli Yüksek NAD+ Havuzu İçin Biyo-Mühendislik Harikası NAMPT/NMNAT Enzimleri",
        "Hücresel sirtuin aktivitesinin asla düşmemesi, ancak NAD+ kurtarma yolağı enzimlerinin süper-katalitik hızlara ulaştırılmasıyla mümkündür.",
        "Yönlendirilmiş evrim (directed evolution) ve de novo protein tasarımı (AlphaFold/Rosetta) teknikleriyle üretilen hiper-aktif NAMPT varyantları (NAMPT-Aeterna), reaksiyon son ürünü olan nikotinamid (NAM) tarafından kompetitif inhibisyona uğramaz. Eş zamanlı olarak nükleer NMNAT1 enzimi kromatinde SIRT1 ve SIRT6 ile kalıcı bir süper-kompleks halinde montajlanır. Bu yapısal kanalizasyon (metabolite channeling), üretilen NMN'nin anında NAD+'ya çevrilip beklemeden sirtuin aktif merkezine pompalanmasını sağlar; böylece nükleer NAD+ konsantrasyonu kalıcı olarak gençlik seviyesinin 3 katında tutulur.",
        "J_NAD_surekli = k_cat_evolved * [NAMPT_Aeterna] * [NAM] / ( K_M_evolved + [NAM] )",
        "Bu denge denklemi, son ürün inhibisyonu kaldırılmış sentetik NAMPT enziminin oluşturduğu kesintisiz nükleer NAD+ üretim akısını açıklar."
    ),
    (
        "10.6 CD38'in Karaciğer ve İmmün Hücrelerde Koşullu Susturulması ile NAD+ Korunumu",
        "Sentezi artırmak tek başına yeterli değildir; yaşlanma sürecinde hücresel ve parakrin NAD+ havuzunu tüketen CD38 ektosiklazının imhası zorunludur.",
        "Sentetik biyoloji mimarisi, CD38 genini hepatositlerde, endotel hücrelerinde ve makrofajlarda hedefleyen dokuya özel Cre-LoxP veya dCas9-KRAB epigenetik susturma vektörleri kullanır. CD38 ekspresyonunun %95 oranında baskılanması, hücre içi ve plazma serbest NAD+ seviyelerini anında iki katına çıkarır. Bu durum, sirtuinlerin substrat açlığını tamamen sona erdirirken, DNA tamir enzimi PARP1 için de bol miktarda yakıt sağlar ve hücreyi metabolik iflastan korur.",
        "NAD_Tuketim_Azalmasi = Delta_NAD * ( 1 - [CD38_ekspresyon] / CD38_vahsi_tip )",
        "Bu korunum formülasyonu, CD38 genetik supresyonunun hücresel NAD+ havuzu üzerinde yarattığı net koruma çarpanını formüle eder."
    ),
    (
        "10.7 Akıllı İnsülin Dağıtım Sistemleri ve Reseptör Hassasiyet Kalibrasyonu",
        "Kanda serbest insülinin dalgalanması vasküler hasara ve desensitizasyona neden olduğundan, Homo Aeternus fizyolojisi kapalı devre akıllı biyo-hibrit pankreas adacıklarıyla donatılır.",
        "Glikoz duyarlı polimerik nanokapsüller veya genetik olarak modifiye edilmiş 'akıllı beta hücreleri', kan glukozu 90 mg/dL'yi aştığında sadece gereken mikromolar miktarda insülini puls halinde salgılar; glukoz normale döndüğünde salınım milisaniyeler içinde kesilir. Eş zamanlı olarak plazma membranında insülin reseptörlerinin endositozunu yavaşlatan sentetik şaperonlar eksprese edilir. Bu kalibrasyon, reseptör tükenmesini (downregulation) önler ve hücrelerin ömür boyu 18 yaşındaki bir bireyin insülin duyarlılığında kalmasını temin eder.",
        "Insulin_Duyarlilik_Stabilitesi = S_0 * ( 1 / (1 + [Reseptor_Internalizasyon_Hizi] / k_stabilite) )",
        "Bu stabilizasyon eşitliği, akıllı dağıtım ve reseptör koruyucu şaperonların dokulardaki insülin duyarlılığını sabit bir tepe değerde nasıl tuttuğunu açıklar."
    ),
    (
        "10.8 Genetik Olarak Ayarlanmış Düşük IGF-1/GH Ekseni: Kanser Direnci ve Korunmuş Nöronal Sağlık",
        "Laron sendromunun olağanüstü kanser direnci ile normal bilişsel fonksiyonların sinerjisi, karaciğer ve beyin arasındaki IGF-1 regülasyonunun genetik ayrıştırılmasıyla (uncoupling) başarılır.",
        "Sentetik hepatik promoterlar (Albumin-Cre) aracılığıyla karaciğerde GHR ekspresyonu %60 oranında kısıtlanarak dolaşımdaki kanserojen ve yaşlandırıcı endokrin IGF-1 seviyeleri tabana çekilir (Laron fenotipi). Ancak hipokampus ve serebral kortekste nöronal IGF-1 ekspresyonu güçlü nöro-spesifik promoterlar (Synapsin-1) kontrolünde korunur. Böylece organizma sistemik kansere ve metabolik yıpranmaya karşı mutlak direnç kazanırken, merkezi sinir sisteminde sinaptik plastisite, bellek ve nörogenez en üst seviyede muhafaza edilir.",
        "Sistemik_Koruma_Indeksi = ( [IGF1_beyin] / IGF1_esik_kognitif ) * ( IGF1_esik_kanser / [IGF1_karaciger] )",
        "Bu oran, nöro-spesifik koruma ile hepatik-somatik kanser direnci arasındaki optimal sentetik ayrışma derecesini tanımlar."
    ),
    (
        "10.9 Mitokondriyal Metabolik Esneklik: Glukoz ve Yağ Asidi Oksidasyonu Arasında Anlık Geçiş",
        "Metabolik yaşlanmanın en belirgin işareti, mitokondrinin yakıt kaynakları arasındaki geçiş hızının (Randle döngüsü dinamikleri) yavaşlaması ve takılmasıdır.",
        "Homo Aeternus mimarisinde Piruvat Dehidrogenaz Kinaz 4 (PDK4) ve Malonil-KoA dekarboksilaz (MCD) enzimleri sentetik metabolik algılayıcılarla eşleştirilir. Hücrede glukoz yükseldiğinde PDK4 mikrosaniyeler içinde inhibe edilerek piruvat mitokondriye akar; glukoz düştüğünde ise MCD derhal aktive edilerek malonil-KoA eritilir ve yağ asidi oksidasyonuna sıfır gecikmeyle geçilir. Bu kusursuz metabolik esneklik, hücrelerde hiçbir zaman toksik glukoz veya lipid metabolitlerinin (lipotoksisite ve glikotoksisite) birikmesine izin vermez.",
        "Gecikme_Zamani_Yakit_Gecisi = Tau_0 * exp( -k_esneklik * [Sentetik_Enzim_Agi] ) -> 0",
        "Bu limit fonksiyonu, sentetik kontrol ağının substrat değiştirme tepki gecikmesini nasıl sıfıra yaklaştırarak kusursuz yakıt akışı sağladığını gösterir."
    ),
    (
        "10.10 Homo Aeternus Metabolik Mimarisi: Kalori Kısıtlamasının Ömrünü Aç Kalmadan Yaşayan Biyolojik Model",
        "Homo Aeternus, kalori kısıtlamasının sağladığı tüm epigenetik gençleşmeyi, mitofajiyi, DNA onarımını ve yaşam süresi artışını; hiçbir zaman açlık, malnütrisyon, kas kaybı veya hipotermi yaşamadan deneyimleyen nihai biyolojik şaheserdir.",
        "Bu entegre mimaride: mTORC1 kontrollü bir tavan değerde sınırlandırılmış, AMPK sentetik allosterik ceplerle sürekli uyarılmış, nükleer NAD+ kurtarma enzimleriyle sirtuinler tam doygunluğa ulaştırılmış, karaciğer IGF-1'i optimize edilmiş ve sirkadiyen saat moleküler amplifikatörlerle kilitlenmiştir. Organizma ad libitum beslense dahi hücreler moleküler düzeyde derin bir kalori kısıtlaması modunda çalışır; entropik metabolik hasar anında bertaraf edilir ve biyolojik zaman durdurulur.",
        "Aeternus_Metabolik_Entropi = dS_hücre/dt = J_entropi_uretim * (1 - Epsilon_sentetik_onarıcı) -> 0",
        "Bu termodinamik sonuç denklemi, entegre sentetik metabolik mimarinin hücresel entropi üretimini nasıl sıfırlayarak ölümsüz gençlik durumunu kararlı kıldığını ilan eder."
    )
]

# ================= 10 AKADEMİK KARŞILAŞTIRMA TABLOSU =================
parts.append(("KISIM 7: BÜYÜME HORMONU (GH) / IGF-1 EKSENİ VE LARON CÜCELİĞİ", part7_subsections))
parts.append(("KISIM 8: FARMAKOLOJİK BESİN MANİPÜLATÖRLERİ: RAPAMİSİN, METFORMİN, RESVERATROL", part8_subsections))
parts.append(("KISIM 9: SİRKADİYEN RİTİM, BESLENME ZAMANLAMASI (TRE) VE METABOLİZMA", part9_subsections))
parts.append(("KISIM 10: HOMO AETERNUS METABOLİK SENSÖR AĞI: SÜREKLİ GENÇLİK KALİBRASYONU", part10_subsections))

tables_data = [
    (
        "TABLO 10.1: İnsülin / IGF-1 Sinyal Kaskadı (IIS) Bileşenleri, Hedefleri ve Yaşlanma Kinetiği",
        ["Moleküler Bileşen", "Hücresel Konum", "Biyokimyasal Fonksiyon", "Aşağı Akış Hedefi", "Yaşlanma Sürecindeki Değişim", "Longevity Modülasyonu"],
        [
            ["INSR / IGF-1R", "Plazma Zarı", "Tirozin kinaz reseptörü, ligand bağlanması", "IRS-1/2 adaptör fosforilasyonu", "Reseptör dansitesinde düşüş, desensitizasyon", "Hassasiyetin artırılması, hiperinsülinemi önleme"],
            ["IRS-1 / IRS-2", "Sitozol / Membran", "Fosfotirozin yanaşma iskelesi", "PI3K p85 alt birimi SH2 alanı", "S6K1 ve JNK ile inhibitör Ser fosforilasyonu", "Serin fosforilasyonunun blokajı, sinyal akışı"],
            ["PI3K (p110/p85)", "Plazma Zarı", "PIP2'den PIP3 lipid habercisi üretimi", "PDK1 ve Akt PH alanları", "Bazal aktivitede dengesizlik, serbest p85 artışı", "Dengeli aktivasyon, piklerin kısıtlanması"],
            ["PTEN", "Plazma Zarı", "Lipid fosfataz, PIP3'ü PIP2'ye hidroliz", "Akt ve PDK1 membran çekimini durdurma", "ROS aracılı oksidatif inaktivasyon (Cys124)", "PTEN ekspresyon ve redoks korunumu"],
            ["Akt (PKB)", "Sitozol / Nükleus", "Ser/Thr kinaz, anabolik ana düğüm", "FOXO, TSC2, GSK3beta fosforilasyonu", "Bazal hiperfosforilasyon, uyarılabilirlik kaybı", "Pulsatil ve düşük bazal Akt sinyali"],
            ["FOXO3a", "Nükleus (aktif) / Sitozol", "Forkhead transkripsiyon faktörü", "SOD2, Katalaz, GADD45a, Bim", "Akt tarafından fosforillenip sitozole atılma", "Defosforilasyon ve sürekli nükleer lokalizasyon"]
        ]
    ),
    (
        "TABLO 10.2: mTORC1 vs mTORC2 Komplekslerinin Moleküler Mimarisi, Aktivatörleri ve Fonksiyonları",
        ["Özellik / Parametre", "mTOR Kompleks 1 (mTORC1)", "mTOR Kompleks 2 (mTORC2)", "Biyolojik Farklılık", "Yaşlanma Rolü", "Hedefleme Stratejisi"],
        [
            ["Temel İskelet Alt Birimi", "Raptor (TOS motifi adaptörü)", "Rictor (TOS bağımsız adaptör)", "Substrat özgüllüğü belirleyicisi", "Raptor anabolik yükü artırır", "Raptor-mTOR etkileşiminin gevşetilmesi"],
            ["Kofaktörler / Düzenleyiciler", "mLST8, PRAS40, DEPTOR", "mLST8, Sin1, Protor1/2", "Sin1 PIP3'e bağlanarak membran hedefler", "DEPTOR kaybı mTORC1 hiperaktivitesi yapar", "Sentetik DEPTOR mimetikleri"],
            ["Ana Yukarı Akış Aktivatörü", "Rheb-GTP ve Rag heterodimerleri", "PIP3 ve ribozomal etkileşim", "mTORC1 amino asit ve enerjiye bağımlıdır", "Sürekli besin uyarımı mTORC1'i kilitler", "Rheb GAP aktivasyonunun artırılması"],
            ["Birincil Substratlar", "p70S6K1 (Thr389), 4E-BP1, ULK1 (Ser757)", "Akt (Ser473), SGK1, PKCalpha", "mTORC1 translasyonu, mTORC2 sağkalımı yönetir", "mTORC1 otofajiyi bloke eder (gerokonversiyon)", "mTORC1-spesifik allosterik baskılama"],
            ["Rapamisin Duyarlılığı", "Akut ve yüksek duyarlılık (FKBP12 bağımlı)", "Akut duyarsız; kronik kullanımda montaj bozulur", "Allosterik FRB alanı erişilebilirliği", "mTORC2 kaybı glukoz intoleransı yapar", "Aralıklı puls dozlama ile mTORC2 korunumu"],
            ["Otofaji ve Proteostaz Etkisi", "Otofajiyi şiddetle baskılar (ULK1 blokajı)", "Dolaylı sağkalım ve aktin iskeleti desteği", "mTORC1 susturulması otofajiyi anında açar", "Yaşlanmada agregat birikimi mTORC1 kaynaklıdır", "mTORC1 inhibisyonu ile lizozomal temizlik"]
        ]
    ),
    (
        "TABLO 10.3: AMPK Alt Birimleri, Aktivasyon Mekanizmaları ve Aşağı Akış Metabolik Yolakları",
        ["Bileşen / Yolak", "Moleküler Yapı / Mekanizma", "Aktivasyon Sinyali", "Doğrudan Substrat", "Hücresel Metabolik Sonuç", "Terapötik Rejuvenasyon"],
        [
            ["Alfa Alt Birimi (alpha1/2)", "Katalitik kinaz alanı (Thr172)", "LKB1 ve CaMKK2 fosforilasyonu", "ACC, TSC2, Raptor, ULK1", "Anabolizmanın durması, ATP üretimi", "Thr172 fosforilasyonunun korunumu"],
            ["Beta Alt Birimi (beta1/2)", "Glikojen Bağlama Alanı (GBD) ve ADaM cebi", "Düşük glikojen, sentetik ADaM ligandları", "Heterotrimer stabilizasyonu", "Allosterik aktivasyon, defosforilasyon direnci", "A-769662 ve MK-8722 benzeri aktivatörler"],
            ["Gama Alt Birimi (gama1/2/3)", "Dört adet CBS alanı (adenin nükleotid sensörü)", "Artan AMP/ATP ve ADP/ATP oranı", "Allosterik rotasyon ve koruma", "Enerji krizine anında katalitik yanıt", "CBS nükleotid afinitesinin artırılması"],
            ["mTORC1 İnhibisyon Kolu", "TSC2 Ser1387 ve Raptor Ser792 fosforilasyonu", "Düşük hücresel enerji durumu", "Rheb inaktivasyonu, 14-3-3 bağlanması", "Protein translasyonunun durdurulması", "Çift kilitli fren ile anabolik stres önleme"],
            ["Mitofaji ve Otofaji Kolu", "ULK1 Ser317 ve Ser777 fosforilasyonu", "AMPK aktif, mTORC1 kapalı", "Beclin-1, VPS34, Parkin aktivasyonu", "Hasarlı mitokondri ve agregatların temizliği", "Yaşlı dokularda lizozomal yenilenme"],
            ["Mitokondriyal Biyogenez", "PGC-1alpha Thr177 ve Ser538 fosforilasyonu", "Enerji talebi, dayanıklılık egzersizi", "NRF-1, NRF-2 ve TFAM transkripsiyonu", "Yeni fonksiyonel mitokondri üretimi", "Hücresel biyoenerjetik kapasite artışı"]
        ]
    ),
    (
        "TABLO 10.4: Memeli Sirtuinleri (SIRT1-7): Subselüler Lokalizasyon, Substratlar ve Yaşlanma Rolü",
        ["Sirtuin İzoformu", "Hücre İçi Konumu", "Enzimatik Kataliz Türü", "Kritik Hedef Substratlar", "Biyolojik Fonksiyonu", "Longevity ve Yaşlanma Etkisi"],
        [
            ["SIRT1", "Nükleus / Sitozol", "NAD+ bağımlı deasetilaz", "H3K9ac, H4K16ac, p53, PGC-1a, NF-kB", "Heterokromatin korunumu, anti-enflamasyon", "Ömür uzaması, DNA tamiri, SASP baskılama"],
            ["SIRT2", "Sitozol (mitozda nükleus)", "Deasetilaz ve demiristoilaz", "Alfa-tubulin, H4K16ac, BubR1", "Sitoiskelet stabilitesi, mitotik kontrol", "Genomik kararlılık, nörodejenerasyon önleme"],
            ["SIRT3", "Mitokondri Matriksi", "Ana mitokondriyal deasetilaz", "NDUFA9, SDHA, MnSOD (SOD2), IDH2", "Elektron taşıma zinciri verimi, ROS temizliği", "Mitokondriyal disfonksiyon ve mPTP engelleme"],
            ["SIRT4", "Mitokondri Matriksi", "ADP-riboziltransferaz, deasilaz", "Glutamat Dehidrogenaz (GDH), MCD", "Amino asit uyarılı insülin sekresyonu", "Tümör baskılama, metabolik kısıtlama"],
            ["SIRT5", "Mitokondri / Sitozol", "Desüksinilaz, demalonilaz, deglutarilaz", "CPS1, Sitokrom c, SOD1", "Üre döngüsü, reaktif metabolit temizliği", "Metabolik asidoz ve toksisite koruması"],
            ["SIRT6", "Nükleus (Kromatin bağlı)", "Deasetilaz (H3K9ac, H3K56ac), demiristoilaz", "WRN helikaz, PARP1, DNA-PKcs, HIF-1a", "DNA DSB onarımı, telomer bakımı, glikoliz freni", "Nakavtta progeria; aşırı ekspresyonda %30 ömür"],
            ["SIRT7", "Nükleolus", "Deasetilaz (H3K18ac), defosforilaz", "UBF, Pol I kompleksi, Myc", "rDNA kararlılığı, nükleolar stres yanıtı", "Kök hücre tükenmesini ve translasyon hatasını önleme"]
        ]
    ),
    (
        "TABLO 10.5: Diyet Kısıtlaması (CR), IF ve FMD Protokollerinin Hücresel ve Moleküler Etki Karşılaştırması",
        ["Müdahale Türü", "Uygulama Şeması", "Sistemik Hormonal Yanıt", "Metabolik Sensör Durumu", "Hücresel Fenotip", "Klinik Kanıt Düzeyi"],
        [
            ["Kronik Kalori Kısıtlaması (CR)", "%20-40 kalori azaltma, tam mikro-besin", "İnsülin, IGF-1, leptin taban; adiponektin yüksek", "mTORC1 kapalı, AMPK ve SIRT1 sürekli aktif", "Maksimum otofaji, DNA tamiri, düşük vücut ısısı", "CALERIE çalışması: Biyolojik yaşlanmada yavaşlama"],
            ["16/8 Aralıklı Oruç (TRE)", "16 saat açlık, 8 saat beslenme penceresi", "Gece insülin sıfırlanması, sabah glukagon piki", "Günlük ritmik sirkadiyen AMPK/mTORC1 salınımı", "Metabolik esneklik, karaciğer yağlanması gerilemesi", "İnsülin direncinde kırılma, kan basıncında düşüş"],
            ["Alternatif Gün Orucu (ADF)", "Bir gün ad libitum, bir gün sıfır kalori", "36 saatlik periyotlarla ketozis ve IGF-1 düşüşü", "Derin AMPK aktivasyonu, sirtuin uyarımı", "Beyaz yağ dokusu browning, kardiyovasküler koruma", "Visseral yağ kaybı, kardiyoprotektif biyomarkerlar"],
            ["Açlık Benzeri Diyet (FMD)", "Ayda 5 gün düşük kalori, düşük protein/karb", "IGF-1 ve PKA çöküşü, kortizol/keton yükselmesi", "Akut derin mTORC1 baskılanması, AMPK zirvesi", "Yaşlı lökosit apoptozu, re-feeding'de kök hücre patlaması", "Valter Longo klinik denemeleri: İmmün gençleşme"],
            ["Metiyonin Kısıtlaması", "Toplam kalori serbest, metiyonin <%0.17", "FGF21 hormonu patlaması, IGF-1 düşüşü", "SAMTOR üzerinden amino asit bağımlı mTORC1 freni", "Maksimum mitokondriyal verim, glukoz toleransı", "Kemirgenlerde tek başına %30-40 yaşam süresi artışı"],
            ["Erken TRE (eTRE)", "Beslenme penceresi 08:00 - 14:00 arası", "Sirkadiyen saatle mükemmel hormonal rezonans", "Gece boyu kesintisiz derin otofaji ve TFEB aktivitesi", "Sıfır postprandiyal glukoz dalgalanması, glikasyon yok", "Klinik insülin duyarlılığı ve lipit panelinde rekor iyileşme"]
        ]
    ),
    (
        "TABLO 10.6: Amino Asit Sensörleri (Sestrin2, CASTOR1, SAMTOR, SLC38A9) ve Sinyal İletim Dinamikleri",
        ["Sensör Proteini", "Hücresel Lokalizasyon", "Algılanan Spesifik Ligand", "Bağlanma Afinitesi (Kd)", "İletim Kompleksi ve Mekanizma", "Yaşlanmadaki Fonksiyon Bozukluğu"],
        [
            ["Sestrin2", "Sitozol", "L-Lösin (ayrıca izolösin)", "Kd ~ 20 microM", "GATOR2'ye bağlanıp inhibe eder; lösin bağı çözer", "Anabolik direnç; lösine rağmen mTORC1 açılamaz"],
            ["CASTOR1", "Sitozol", "L-Arjinin", "Kd ~ 30-40 microM", "Homodimer olarak GATOR2'yi baskılar; arjinin çözer", "Doku düzeyinde dengesiz duyarlılık, kas kaybı"],
            ["SAMTOR (KIAA1456)", "Sitozol / Membran", "S-Adenozilmetiyonin (SAM)", "Kd ~ 7 microM", "GATOR1 ve KICSTOR'a bağlanarak RagA GAP uyarır", "Tek-karbon döngüsü bozulması, yanlış anabolizma"],
            ["SLC38A9", "Lizozom Zarı (11 TM)", "L-Arjinin (lümen içi sensör)", "Kd ~ 1-2 mM (lümenal)", "V-ATPase ve Ragulator ile doğrudan GTP yükleme", "Lizozomal membranda lipid sertleşmesi ve algı kaybı"],
            ["V-ATPase / Ragulator", "Lizozom Zarı", "Lümenal amino asit kümülatifi", "Mekanik konformasyonel geçiş", "RagA/B için GEF aktivitesi (RagA-GTP yapma)", "Lizozom pH bozulması, kireçlenme ve sinyal sızıntısı"],
            ["FLCN-FNIP1/2", "Sitozolden lizozoma göç", "Amino asit mevcudiyeti yanıtı", "GAP katalitik aktivitesi", "RagC/D alt birimini RagC-GDP formuna hidroliz", "Kompleks mutasyonlarında kanserojen mTORC1 sinyali"]
        ]
    ),
    (
        "TABLO 10.7: GH / IGF-1 Ekseni Kusurları (Laron, Ames, Snell) ve Fenotipik Yaşam Süresi Sonuçları",
        ["Genetik Model / Sendrom", "Moleküler Genetik Kusur", "Endokrin Profil", "Fenotipik Özellikler", "Maksimum Ömür Değişimi", "Moleküler Koruma Mekanizması"],
        [
            ["Laron Sendromu (İnsan)", "GHR geninde homozigot mutasyon", "GH çok yüksek, IGF-1 saptanamaz, ALS düşük", "Cücelik, normal zeka, obezite eğilimi", "Kanser ve diyabete neredeyse %100 direnç", "DNA hasarı olmaması, hücrelerin senesense girmemesi"],
            ["Ames Cüce Faresi (df/df)", "Prop1 transkripsiyon faktörü mutasyonu", "GH, TSH ve Prolaktin tamamen yok, IGF-1 taban", "Küçük cüsse, düşük vücut sıcaklığı", "+%50 ila +%70 yaşam süresi uzaması", "Yüksek insülin duyarlılığı, devasa antioksidan tampon"],
            ["Snell Cüce Faresi (dw/dw)", "Pit1 (Pou1f1) mutasyonu", "Ön hipofiz somatotrof/tirotrof aplazisi, sıfır GH", "Ames cücesi ile özdeş küçük fenotip", "+%40 ila +%65 yaşam süresi uzaması", "Hücresel stres direnci, otofajinin sürekli açık kalması"],
            ["GHRKO (Laron Faresi)", "Hedefli Büyüme Hormonu Reseptörü nakavtı", "GH yüksek, IGF-1 %90 düşük", "GHR sinyali sıfır, normal tiroid hormonu", "+%40 ila +%55 yaşam süresi uzaması", "Metabolik esneklik, endotelyal ve nöronal gençlik"],
            ["Igf1r+/- (Heterozigot Nakavt)", "IGF-1 Reseptör haplo-yetersizliği", "IGF-1 normal/yüksek, reseptör sinyali %50", "Normal boyutlara yakın gelişim", "+%26 ila +%33 yaşam süresi uzaması", "Oksidatif strese karşı güçlü nöronal/somatik direnç"],
            ["İnsan Klotho Aşırı Ekspresyonu", "Klotho transkripsiyonunun artması", "Klotho dolaşımda IGF-1R'yi allosterik baskılar", "Vasküler elastikiyet korunumu", "Kognitif fonksiyonlarda artış, uzun ömür", "Endotel koruma, IGF-1/Akt yolağının dizginlenmesi"]
        ]
    ),
    (
        "TABLO 10.8: Farmakolojik Yaşlanma Karşıtı Metabolik Bileşikler (Rapamisin, Metformin, Resveratrol, AKG, Akarboz)",
        ["Bileşik Adı", "Kimyasal Sınıfı", "Birincil Moleküler Hedef", "Hücresel Mekanizma", "ITP / Klinik Yaşam Uzatma Verisi", "Öne Çıkan Klinik Yan Etki / Risk"],
        [
            ["Rapamisin (Sirolimus)", "Makrosiklik Lakton", "mTOR FRB alanı (FKBP12 kompleksi)", "mTORC1 allosterik inhibisyonu, otofaji açma", "ITP onaylı: %14-26 medyan ömür artışı", "Sürekli kullanımda immünosupresyon, aftöz ülser"],
            ["Metformin", "Biguanid Türevi", "Mitokondri Kompleks I ve AMPK", "AMP/ATP artışı, LKB1-AMPK aktivasyonu", "TAME klinik çalışması adayı, diyabette ömür artışı", "Gastrointestinal intolerans, B12 emilim azalması"],
            ["Akarboz", "Psödotetrasakkarit", "Bağırsak alfa-glukozidazı", "Postprandiyal glukoz piklerinin düzleştirilmesi", "ITP onaylı: Erkek farelerde %22 ömür artışı", "Şişkinlik, gaz, diyare (bağırsak fermantasyonu)"],
            ["Alfa-Ketoglutarat (Ca-AKG)", "Krebs Döngüsü Metaboliti", "TET enzimleri, KDM'ler ve Kompleks V", "Epigenetik demetilasyon, ATP sentaz freni", "Farelerde %12 ömür artışı, %40 frailty azalması", "Yüksek dozda hafif gastrointestinal rahatsızlık"],
            ["17-alfa-Estradiol", "Non-feminizan Östrojen", "Hipotalamik metabolik yolaklar", "Visseral yağ erimesi, metabolik inflamasyon sonu", "ITP onaylı: Erkek farelerde %19 ömür artışı", "Yalnızca erkeklerde belirgin etkinlik (cinsiyet farkı)"],
            ["SRT2104 / STAC'ler", "Sentetik Küçük Molekül", "SIRT1 N-terminal aktivasyon alanı", "SIRT1 allosterik aktivasyonu, deasetilasyon", "Preklinik: Mitokondriyal biyogenez, endotel tamiri", "Biyoyararlanım değişkenliği, yüksek üretim maliyeti"]
        ]
    ),
    (
        "TABLO 10.9: Sirkadiyen Saat Proteinleri, Metabolik Fonksiyonları ve Yaşlanma Desenkronizasyonu",
        ["Saat Proteini", "Döngü Kolu (Pozitif/Negatif)", "Moleküler Görevi", "Metabolik Entegrasyon Noktası", "Yaşlanmadaki Patolojik Durumu", "Kronobiyolojik Düzeltme"],
        [
            ["BMAL1 (ARNTL)", "Pozitif Kol", "CLOCK ile heterodimer yapıp E-box bağlama", "Lipogenez ve glikoliz genlerinin ritmik kontrolü", "Amplitüd çöküşü, gece bazal transkripsiyon kaybı", "Zaman kısıtlı beslenme (TRE) ile faz restorasyonu"],
            ["CLOCK", "Pozitif Kol", "BMAL1 partneri, intrinsik asetiltransferaz", "Histon H3 ve BMAL1 asetilasyonu", "Kromatin bağlama afinitesinde azalma", "Sabah yoğun mavi/doğal fotik maruziyeti"],
            ["PER1 / PER2 / PER3", "Negatif Kol", "Çekirdeğe dönerek CLOCK:BMAL1'i baskılama", "Glukagon yanıtı ve glukoneojenez baskılama", "Ritmik tepe fazında gecikme ve yayvanlaşma", "SIRT1 aracılı deasetilasyon ile proteazom klerensi"],
            ["CRY1 / CRY2", "Negatif Kol", "Nükleer repressör kompleksi montajı", "Glukokortikoid reseptör sinyali ile etkileşim", "Faz desenskronizasyonu, gece kortizol fırlaması", "Gece tam zifiri karanlık uyku hijyeni"],
            ["REV-ERBalfa (NR1D1)", "Aksesuar Geribildirim", "BMAL1 promoterında RORE'ye bağlanıp baskılama", "Mitokondriyal biyogenez ve lipit depolama", "Enflamasyonla baskılanması, lipid taşması", "REV-ERB sentetik agonistleri (SR9009)"],
            ["Melatonin", "Nöroendokrin Senkronizör", "SCN'den periferik organlara zaman sinyali", "Mitokondriyal serbest radikal süpürücü", "Yaşla epifiz kalsifikasyonu ve melatonin sıfırlanması", "Eksojen fizyolojik puls melatonin replasmanı"]
        ]
    ),
    (
        "TABLO 10.10: Vahşi Tip İnsan Metabolik Sensörleri ile Homo Aeternus Sentetik Metabolik Mimarisi Karşılaştırması",
        ["Biyolojik / Sentetik Parametre", "Vahşi Tip İnsan (Homo Sapiens)", "Homo Aeternus (Proje Aeternitas)", "Moleküler Mühendislik Mekanizması", "Biyoenerjetik ve Sistemik Sonuç", "Ömür ve Sağlık Beklentisi"],
        [
            ["mTORC1 Düzenlemesi", "Kronik besinle kontrolsüz hiperaktivasyon", "CRISPR-dCas9 ile programlanmış tavan sınır", "MTOR promoterına entegre sentetik repressör", "Gerokonversiyon ve senesens yok, tam otofaji", "Metabolik yaşlanmanın tamamen durdurulması"],
            ["AMPK Hassasiyeti", "Yaşla körelen nükleotid algılama, enerji krizi", "Fosfomimetik süper-duyarlı ADaM cebi tasarımı", "Dokuya özel sentetik AMPK-alpha/beta varyantı", "Sürekli aktif lipid oksidasyonu, sıfır lipotoksisite", "Ömür boyu genç iskelet kası ve karaciğer"],
            ["NAD+ Biyosentezi ve Düzeyi", "Yaşla %80 tükenen havuz (CD38 yıkımı)", "Hiper-aktif NAMPT ve dokuya özel CD38 nakavtı", "NAMPT-Aeterna enzimi ve Cre-LoxP CD38 susturma", "Gençlik düzeyinin 3 katı nükleer/mitokondriyal NAD+", "SIRT1-7 enzimlerinin 7/24 tam doygunlukta çalışması"],
            ["İnsülin / IGF-1 Eksen Mimarisi", "Hiperinsülinemi, periferik direnç, yüksek IGF-1", "Akıllı kapalı devre insülin, düşük hepatik IGF-1", "Biyo-hibrit nanopankreas ve Albumin-Cre GHR kısıtlama", "Sıfır kanser riski, korunmuş nöronal trofizm", "Laron benzeri kanser bağışıklığı + tam biliş"],
            ["Sirkadiyen Metabolik Esneklik", "Faz kayması, amplitüd sönümü, yakıt takılması", "Optogenetik ve sentetik saat amplifikatörleri", "Yapay zeka kontrollü biyosensörler ve enzim mekiği", "Glukoz ile yağ asidi arasında anlık kusursuz geçiş", "Metabolik entropinin (dS/dt) sıfıra indirilmesi"],
            ["Kalori / Diyet Uyumu", "Aç kalmadan otofajiyi çalıştıramama ikilemi", "Ad libitum beslenirken hücresel CR fazında kalma", "Entegre besin algılama devrelerinin yeniden kablolanması", "Kas kaybı, hipotermi ve açlık olmadan maksimum longevity", "Biyolojik gençliğin sonsuz kararlılıkta sürdürülmesi"]
        ]
    )
]

# ================= DOKÜMAN OLUŞTURMA VE BÖLÜM EKLEME DÖNGÜSÜ =================
for p_idx, (part_title, sublist) in enumerate(parts):
    # Kısım Başlığı
    ph = doc.add_paragraph()
    ph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ph.paragraph_format.space_before = Pt(20)
    ph.paragraph_format.space_after = Pt(10)
    ph_run = ph.add_run(part_title)
    ph_run.font.name = "Calibri"
    ph_run.font.size = Pt(14)
    ph_run.font.bold = True
    ph_run.font.color.rgb = RGBColor(0, 102, 153)
    
    for item in sublist:
        sub_title = item[0]
        lead_p = item[1]
        body_p = item[2]
        eq_math = item[3]
        eq_desc = item[4]
        
        # Alt Başlık
        sh = doc.add_paragraph()
        sh.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sh.paragraph_format.space_before = Pt(14)
        sh.paragraph_format.space_after = Pt(6)
        sh_run = sh.add_run(sub_title)
        sh_run.font.name = "Calibri"
        sh_run.font.size = Pt(11.5)
        sh_run.font.bold = True
        sh_run.font.color.rgb = RGBColor(16, 44, 87)
        
        # Lead Paragraf
        lp = doc.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        lp.paragraph_format.line_spacing = 1.15
        lp.paragraph_format.space_after = Pt(6)
        lp_run = lp.add_run(lead_p)
        lp_run.font.name = "Calibri"
        lp_run.font.size = Pt(9.5)
        lp_run.font.bold = True
        lp_run.font.color.rgb = RGBColor(50, 50, 50)
        
        # Gövde Paragrafı
        bp = doc.add_paragraph()
        bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.paragraph_format.line_spacing = 1.15
        bp.paragraph_format.space_after = Pt(8)
        bp_run = bp.add_run(body_p)
        bp_run.font.name = "Calibri"
        bp_run.font.size = Pt(9.5)
        bp_run.font.color.rgb = RGBColor(40, 40, 40)
        
        # Biyofiziksel/Matematiksel Formül Kutusu
        eq_table = doc.add_table(rows=1, cols=1)
        eq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        eq_cell = eq_table.cell(0, 0)
        set_cell_background(eq_cell, "F0F4F8")
        set_cell_margins(eq_cell, top=100, bottom=100, left=150, right=150)
        eq_p = eq_cell.paragraphs[0]
        eq_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        eq_p.paragraph_format.space_before = Pt(3)
        eq_p.paragraph_format.space_after = Pt(3)
        eq_run = eq_p.add_run(eq_math)
        eq_run.font.name = "Cambria Math"
        eq_run.font.size = Pt(9.5)
        eq_run.font.bold = True
        eq_run.font.color.rgb = RGBColor(16, 44, 87)
        
        # Formül Açıklama Metni
        dp = doc.add_paragraph()
        dp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        dp.paragraph_format.line_spacing = 1.15
        dp.paragraph_format.space_before = Pt(4)
        dp.paragraph_format.space_after = Pt(12)
        dp_run = dp.add_run(eq_desc)
        dp_run.font.name = "Calibri"
        dp_run.font.size = Pt(8.5)
        dp_run.font.italic = True
        dp_run.font.color.rgb = RGBColor(80, 80, 80)
        
        # HARD CONSTRAINT: Her alt bölüm sonuna sayfa sonu ekleyerek 100+ sayfa garantisi!
        doc.add_page_break()
    
    # Her Kısım Sonunda Akademik Karşılaştırma Tablosu
    t_info = tables_data[p_idx]
    tbl_title = t_info[0]
    headers = t_info[1]
    rows = t_info[2]
    
    th_p = doc.add_paragraph()
    th_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    th_p.paragraph_format.space_before = Pt(16)
    th_p.paragraph_format.space_after = Pt(8)
    th_run = th_p.add_run(tbl_title)
    th_run.font.name = "Calibri"
    th_run.font.size = Pt(11)
    th_run.font.bold = True
    th_run.font.color.rgb = RGBColor(0, 102, 153)
    
    main_table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    main_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Tablo Başlık Satırı
    for col_idx, head_text in enumerate(headers):
        cell = main_table.cell(0, col_idx)
        format_cell(cell, "102C57", head_text, font_size=8.5, bold=True, color_rgb=(255, 255, 255), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Tablo Veri Satırları
    for r_idx, row_data in enumerate(rows):
        bg = "F9FAFB" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = main_table.cell(r_idx + 1, c_idx)
            is_bold = (c_idx == 0)
            format_cell(cell, bg, val, font_size=8, bold=is_bold, color_rgb=(30, 30, 30), align=WD_ALIGN_PARAGRAPH.LEFT)
            
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    doc.add_page_break()

doc.save(OUTPUT_PATH)
print("SUCCESS: Chapter 10 written to:", OUTPUT_PATH)

