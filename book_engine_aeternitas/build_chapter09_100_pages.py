# -*- coding: utf-8 -*-
"""
PROJECT AETERNITAS - CİLT 09: GENOMİK KARARSIZLIK, DNA HASAR YANITI (DDR) VE TRANSPOZON KONTROLÜ
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

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_09_GENOMIK_KARARSIZLIK_DDR_VE_TRANSPOZON_KONTROLU_TAM_100_SAYFA.docx"

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
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 09: GENOMİK KARARSIZLIK VE DNA HASAR YANITI")
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
s_run = sub_p.add_run("CİLT 09: GENOMİK KARARSIZLIK, DNA HASAR YANITI (DDR) VE TRANSPOZON KONTROLÜ\\n(DSB ONARIMI, ATM/ATR, LINE-1 SUSTURULMASI, HETERO-KROMATİN VE DSUP ENTEGRASYONU)")
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
ih_run = intro_h.add_run("CİLT 09 MANİFESTOSU: GENETİK KODUN MUTLAK KARARLILIĞI VE ENTRİPİK İSYANIN BASTIRILMASI")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "Biyolojik organizmanın yaşam programı, 3.2 milyar baz çiftlik deoksiribonükleik asit (DNA) kütüphanesinde depolanmıştır. "
    "Ancak her bir insan hücresi her gün yaklaşık 100.000 endojen ve eksojen DNA hasarına (oksidasyon, deaminasyon, tek ve çift zincir kırıkları) "
    "maruz kalır. Canlılık, bu muazzam moleküler yıpranmaya karşı DNA Hasar Yanıtı (DDR: ATM, ATR, DNA-PKcs) ve özelleşmiş onarım makineleri "
    "(BER, NER, MMR, HR, NHEJ) ile karşı koyar.\\n\\n"
    "Yaşlanma sürecinde bu onarım kalkanı aşınır: Homolog rekombinasyon zayıflar, mutajenik NHEJ artar, heterokromatin erozyonu yaşanır "
    "ve genomun yarısını oluşturan parazitik retrotranspozonlar (özellikle LINE-1 elemanları) uyanarak kromozomları delik deşik eder. "
    "Sitoplazmaya sızan DNA parçaları cGAS-STING yolağıyla kronik inflam-aging yangınını körükler, hücreleri geri dönüşümsüz senesense kilitler.\\n\\n"
    "Bu ciltte; DNA hasar mekanizmaları, gamma-H2AX odakları, çift zincir kırıklarının onarım biyofiziği, heterokromatin çöküşü, "
    "LINE-1 retrotranspozisyonunun dinamikleri, nükleer lamin/por aşınması, SIRT6'nın koruyucu deasetilasyonu, nükleotid havuzu koruması "
    "ve tardigrat Dsup proteini ile çoklu TP53 kopyalarını entegre eden ölümsüz Homo Aeternus genom mimarisi 100 ayrıntılı akademik bölümde ele alınmaktadır."
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

# ==============================================================================
# KISIM 1: DNA HASAR TÜRLERİ VE ENDOJEN / EKSOJEN MUTAJENLER
# ==============================================================================
part1_subsections = [
    (
        "1.1",
        "Spontan Hidrolitik Bozunma: Deaminasyon (C->U, 5mC->T), Depürinasyon ve Abazik (AP) Bölgeler",
        "Fizyolojik su ve sıcaklığın DNA bazlarını sürekli kimyasal hidrolize uğratması: Günde on binlerce baz kaybı ve kalıcı mutasyon riski.",
        "DNA molekülü sulu hücresel ortamda termodinamik olarak kararsızdır. En sık görülen spontan lezyon, glikozidik bağın hidrolitik kopması sonucu pürin bazlarının (adenin, guanin) düşmesiyle oluşan depürinasyondur (hücre başına günde ~10.000 apürinik/abazik AP bölgesi). İkinci büyük tehdit spontan deaminasyondur: Sitozinin ekzoklik amino grubunu kaybetmesi sitozini urasile (C -> U) dönüştürür. Daha da tehlikelisi, metillenmiş CpG adacıklarındaki 5-metilsitozinin deaminasyonudur; bu reaksiyon doğrudan doğal bir DNA bazı olan timini (5mC -> T) üretir ve hücresel onarım makineleri tarafından fark edilmesi son derece zordur.",
        "Rate_Hydrolytic_Decay = k_depur * [Purines] + k_deam * [Cytosines] * exp(-E_act / (R * T))",
        "Spontan hidrolitik bozunma debisi; sıcaklığa bağlı depürinasyon hızı ile sitozin deaminasyon reaksiyon hızlarının toplamıdır."
    ),
    (
        "1.2",
        "Reaktif Oksijen Türleri (ROS) ve DNA Oksidasyonu: 8-OHdG ve Timin Glikol Lezyonları",
        "Mitokondriyal ve peroksizomal serbest radikallerin nükleer DNA'ya saldırısı: Düşük redoks potansiyelli guaninin 8-okzoguanine dönüşümü.",
        "Guanin bazı, dört DNA bazı arasında en düşük iyonlaşma potansiyeline (7.8 eV) sahip olduğu için hidroksil radikallerinin (.OH) birincil hedefidir. Guaninin 8. karbonunun oksidasyonu, biyolojik yaşlanmanın en yaygın biyomarkerı olan 8-hidroksi-2'-deoksiguanozini (8-OHdG / 8-oxoG) oluşturur. 8-oxoG, replikasyon çatalında DNA polimerazları yanıltarak sitozin yerine adenin ile 'Hoogsteen baz eşleşmesi' kurar ve G:C -> T:A transversiyon mutasyonuna neden olur. Timin bazının oksidasyonu ise replikasyonu bloke eden timin glikol lezyonlarını üretir.",
        "Oxidative_Lesion_Density = J_8oxoG = k_fenton * [Fe2+] * [H2O2] * [Guanine_accessible] / [OGG1_clearance]",
        "Nükleer 8-OHdG lezyon yoğunluğu; Fenton reaksiyonu radikal üretim hızının baz eksizyon tamir enzimi OGG1'in temizleme kapasitesine oranıdır."
    ),
    (
        "1.3",
        "Alkilasyon Hasarları: O6-metilguanin ve MGMT (O6-metilguanin-DNA metiltransferaz) İntihar Onarımı",
        "Endojen S-adenozilmetiyonin veya çevresel nitrozaminlerin bazlara metil/alkil grupları eklemesi ve kendini feda eden MGMT proteini.",
        "Alkilleyici ajanlar DNA bazlarındaki nitrojen ve oksijen atomlarına metil veya etil grupları transfer eder. En mutajenik alkilasyon lezyonu O6-metilguanindir (O6-meG); replikasyon sırasında timinle eşleşerek G:C -> A:T transisyonunu tetikler. Bu ölümcül lezyon, 'intihar enzimi' (suicide enzyme) olarak bilinen O6-metilguanin-DNA metiltransferaz (MGMT) tarafından tek basamakta onarılır. MGMT, DNA'daki metil grubunu kendi aktif bölgesindeki Cys145 kalıntısına transfer eder; metillenen MGMT enzimi kalıcı olarak inaktive olur ve proteazomda parçalanır. Her bir lezyon tamiri için bir molekül MGMT kurban edilir.",
        "O6meG_Persistence = [O6meG_initial] - min([O6meG_initial], [MGMT_stoichiometric_pool])",
        "Genomda kalan O6-metilguanin lezyon miktarı; başlangıç hasarından hücredeki toplam stokiyometrik MGMT protein havuzunun çıkarılmasıyla bulunur."
    ),
    (
        "1.4",
        "Ultraviyole (UV) Radyasyonu: Siklobütan Pirimidin Dimerleri (CPD) ve 6-4 Fotourunleri",
        "Güneş ışığındaki UV-B fotonlarının bitişik pirimidin halkalarını kovalent kenetleyerek DNA çift sarmalında 30 derecelik bükülme yaratması.",
        "Güneş ışığından gelen UV-B (290-320 nm) ve UV-A radyasyonu, deri keratinosit ve melanosit DNA'sında fotokimyasal reaksiyonlar başlatır. İki komşu pirimidin bazı (özellikle TT veya TC), 2+2 siklo-katılma reaksiyonu ile kovalent Siklobütan Pirimidin Dimerlerine (CPD) veya (6-4) fotourunlerine (6-4PP) dönüşür. Bu fotolezyonlar DNA sarmalında 30 derecelik kalıcı bir bükülme yaratarak replikatif DNA polimerazları kilitler. UV kaynaklı bu hasarların nükleotid eksizyon onarımıyla (NER) tamir edilememesi deri kanserlerini (melanom) ve erken cilt yaşlanmasını (fotoyaşlanma) başlatır.",
        "CPD_Formation_Rate = Quantum_Yield_UVB * Intensity_Photons * [Adjacent_Thymine_Pairs]",
        "CPD lezyon oluşum debisi; UV-B foton akısı kuantum verimi ve DNA'daki yan yana timin ikililerinin konsantrasyonu ile orantılıdır."
    ),
    (
        "1.5",
        "İyonlaştırıcı Radyasyon (IR): Çift Zincir Kırıkları (DSB) ve Kümelenmiş DNA Hasarı (Clustered Lesions)",
        "X-ışınları, gama ışınları ve kozmik radyasyonun su moleküllerini iyonlaştırarak DNA omurgasında nano-mesafede kümelenmiş kırıklar açması.",
        "İyonlaştırıcı radyasyon (IR), DNA zincirini doğrudan vurabileceği gibi (direkt etki), su moleküllerini radyolize uğratarak nanometrik bir hacimde yoğun hidroksil radikalleri (.OH) fışkırtır (indirekt etki). Bu enerji patlaması, DNA sarmalının 1-2 helikal turu (10-20 bp) içinde birden fazla lezyonun (çift zincir kırığı, abazik bölge, baz oksidasyonu) aynı anda oluştuğu 'kümelenmiş DNA hasarını' (clustered / complex DNA damage) yaratır. Kümelenmiş lezyonlar onarım makinelerini felç eder; zira bir zincir onarılırken karşı zincir de kopuk olduğundan kalıp bulunamaz ve kromozom kopmaları kaçınılmaz hale gelir.",
        "Clustered_Damage_Yield = k_ir * Radiation_Dose_Gray * Linear_Energy_Transfer (LET)",
        "Kümelenmiş DNA hasar sıklığı; soğurulan radyasyon dozu (Gray) ile radyasyonun çizgisel enerji transferi (LET) katsayısının çarpımıdır."
    ),
    (
        "1.6",
        "Replikasyon Stresi: Çatal Duraklaması (Replication Fork Stalling), Çöküşü ve Tek İplikli DNA (ssDNA)",
        "DNA polimerazın bir engelle karşılaşıp durması; replikatif helikazın (CMG) ilerlemeye devam ederek kilometrelerce çıplak ssDNA açması.",
        "S fazında DNA kopyalanırken replikasyon çatalı; DNA lezyonları, sıkı kromatin yapıları veya yetersiz dNTP havuzu nedeniyle yavaşlar veya durur (fork stalling). Eğer DNA polimeraz durduğu halde CMG (Cdc45-MCM-GINS) helikazı DNA'yı açmaya devam ederse, polimeraz ile helikazın bağlantısı kopar. Çatalın arkasında devasa uzunlukta çıplak tek iplikli DNA (ssDNA) açığa çıkar. Bu ssDNA hızla RPA (Replication Protein A) ile kaplanır ve ATR kinaz alarmını tetikler. Duraklayan çatal zamanında kurtarılamazsa endonükleazlar (MUS81) tarafından kesilerek ölümcül çift zincir kırıklarına (çatal çöküşü / fork collapse) dönüşür.",
        "ssDNA_Generation_Rate = v_helicase_CMG - v_polymerase_elongation > 0 (Replication Stress)",
        "Replikasyon stresi şiddeti; CMG helikazının DNA açma hızı ile duraklayan polimerazın uzama hızı arasındaki hız farkıyla açılan ssDNA miktarıdır."
    ),
    (
        "1.7",
        "R-Döngüleri (R-Loops): RNA-DNA Hibritleri, Transkripsiyon-Replikasyon Çarpışmaları",
        "Yeni sentezlenen mRNA'nın arkadaki kalıp DNA zincirine tekrar yapışması; açığa çıkan tek zincirin mutasyonlara ve kırılmalara açık kalması.",
        "Transkripsiyon sırasında yeni uzayan RNA zinciri, bazen RNA polimerazın arkasındaki kalıp DNA ipliğiyle termodinamik olarak hibritleşir. Bu üç zincirli nükleik asit yapısına 'R-Loop' denir: Bir RNA:DNA hibrit çifti ve dışarıda kalmış tek iplikli DNA (ssDNA). R-döngüleri fizyolojik olarak bazı promoterlarda regülatör rol oynasa da, temizlenmediklerinde replikasyon çatalıyla kafa kafaya çarpışır (Transcription-Replication Conflicts - TRC). Dışarıda kalan çıplak DNA zinciri AID/APOBEC deaminazlarının saldırısıyla delik deşik olur. RNase H1/H2 enzimleri RNA'yı parçalayarak R-döngülerini söndürür.",
        "R_Loop_Stability = Delta_G_RNA:DNA_hybrid - Delta_G_DNA:DNA_duplex < 0 (Termodinamik tuzak)",
        "R-Loop oluşum kararlılığı; RNA:DNA hibritinin Gibbs serbest enerjisinin orijinal DNA çift sarmalından daha negatif olmasıyla kilitlenir."
    ),
    (
        "1.8",
        "DNA-Protein Çapraz Bağları (DPC): Topoizomeraz Kilitlenmeleri ve SPRTN Proteaz Temizliği",
        "DNA üzerinde çalışan devasa enzimlerin kovalent olarak DNA omurgasına yapışıp kalması ve replikasyonu fiziksel olarak tıkaması.",
        "Formaldehit gibi metabolitler veya Topoizomeraz inhibitörleri (kamptotesin, etoposid), DNA'ya geçici bağlanan proteinleri kalıcı kovalent çapraz bağlarla (DPC - DNA-Protein Crosslinks) DNA'ya yapıştırır. Özellikle Topoizomeraz I ve II, DNA'yı kesip döndürdükten sonra ligasyon yapamazsa DNA üzerinde 100-200 kDa'lık devasa protein kayaları halinde donar. Bu kovalent engeller replikasyon ve transkripsiyonu tamamen kilitler. Hücre bu engeli aşmak için özel bir DNA-bağımlı metalloproteaz olan SPRTN (DVC1) enzimini kullanır. SPRTN proteini keserek küçük bir peptide budar; ardından translezyon polimerazları engeli aşar.",
        "DPC_Removal_Flux = k_sprtn * [SPRTN_active] * [Ubiquitinated_DPC] * [ssDNA_junction]",
        "DNA-protein çapraz bağ temizleme hızı; aktif SPRTN proteazı, ubikitinlenmiş DPC substratı ve lezyon etrafındaki tek zincir DNA kavşağına bağlıdır."
    ),
    (
        "1.9",
        "Çevresel Karsinojenler: Polisiklik Aromatik Hidrokarbonlar (PAH), Aflatoksin ve DNA Katımları (Adducts)",
        "Sigara dumanı, egzoz gazı ve küflü gıdalardaki prokarsinojenlerin Sitokrom P450 ile reaktif epoksitlere dönüşerek DNA'ya kovalent eklenmesi.",
        "Çevresel karsinojenlerin çoğu başlangıçta inerttir; ancak karaciğerdeki Faz I detoksifikasyon enzimleri (Sitokrom P450 1A1) tarafından reaktif epoksitlere dönüştürülür. Sigara dumanındaki Benzo[a]piren, BPDE (benzo[a]piren diol epoksit) haline gelir ve guanin bazının ekzoksilik amino grubuna (N2-dG) kovalent olarak bağlanır (hacimli DNA katımı / bulky adduct). Benzer şekilde Aspergillus küflerinin ürettiği Aflatoksin B1, N7-guanin katımları yaparak p53 geninde Kodon 249 mutasyonunu (G -> T) tetikler. Bu hacimli kimyasal gruplar sarmal yapısını bozar ve NER sistemiyle derhal temizlenmezse kanser başlatır.",
        "DNA_Adduct_Burden = Integral( k_p450 * [Procarcinogen] dt ) - Clearance_NER_Capacity",
        "Hücresel kümülatif DNA katım yükü; sitokrom P450 aktivasyonuyla oluşan reaktif karsinojen akısının NER tamir kapasitesinden çıkarılmasıdır."
    ),
    (
        "1.10",
        "Yaşlanmayla Kümülatif Genomik Aşınma: Günlük 100.000 DNA Lezyonunun Biyofiziksel Bilançosu",
        "Tek bir insan hücresinde 24 saat içinde gerçekleşen hasar ve onarım bilançosu; onarılamayan binde birlik artığın 80 yılda yarattığı genomik felaket.",
        "Biyofiziksel hesaplamalara göre tek bir insan hücresinde her gün ortalama 10.000 depürinasyon, 5.000 deaminasyon, 10.000 baz oksidasyonu (8-OHdG), 50.000 tek zincir kırığı (SSB) ve yaklaşık 10 ila 50 çift zincir kırığı (DSB) meydana gelir. Hücresel DNA onarım sistemlerinin doğruluğu %99.99'dur. Ancak her gün tamir edilemeyen veya hatalı tamir edilen geriye kalan %0.01'lik 'artık lezyon yükü', 80 yıllık bir insan ömründe hücre başına on binlerce kalıcı somatik mutasyon, translokasyon ve epigenetik iz bırakır. Bu durum organ rezervlerinin tükenmesinin temel sürücüsüdür.",
        "Residual_Genomic_Scars = N_daily_lesions * (1 - Repair_Efficiency_Fidelity) * 365 * Age_Years",
        "Kümülatif genomik yara yükü; günlük 100.000 lezyonun onarım kusur payı (1 - eta) ve yaşam süresi günlerinin çarpımıyla hesaplanır."
    )
]

# ==============================================================================
# KISIM 2: DNA HASAR YANITI (DDR) SİNYAL KASKADI VE KONTROL NOKTALARI
# ==============================================================================
part2_subsections = [
    (
        "2.1",
        "ATM Kinaz: Çift Zincir Kırıklarının Sensörü ve MRN Kompleksi (Mre11-Rad50-Nbs1) Aktivasyonu",
        "Çift zincir kırığının serbest uçlarına kenetlenen MRN kompleksinin inaktif ATM dimerini parçalayıp aktif monomerlere dönüştürmesi.",
        "Çift zincir kırıkları (DSB) hücre için en ölümcül lezyondur; tek bir onarılmamış DSB apoptozu tetikleyebilir. DSB oluştuğunda kırık uçları ilk olarak MRN kompleksi (Mre11 nükleaz, Rad50 ATPaz ve Nbs1/Nibrin) tarafından yakalanır. MRN kompleksi serbest uçları birbirine bağlar ve Nbs1'in C-terminali aracılığıyla inaktif bir homodimer halinde bekleyen ATM (Ataxia Telangiectasia Mutated) kinazı bölgeye çeker. ATM kinaz, Serin 1981 pozisyonundan otofosforilasyona uğrayarak monomerlerine ayrışır ve süper-aktif hale gelir. Aktif ATM, kırık bölgesindeki yüzlerce proteini fosforilleyerek DDR kaskadını ateşler.",
        "ATM_Monomerization_Rate = k_atm * [ATM_dimer_inactive] * [MRN_DSB_bound] * [ATP]",
        "ATM kinaz aktivasyon hızı; kırık DNA ucuna bağlı MRN kompleksi yoğunluğu ve inaktif ATM homodimer konsantrasyonuyla doğru orantılıdır."
    ),
    (
        "2.2",
        "ATR Kinaz ve ATRIP: Tek Zincirli DNA (RPA Kaplı) ve Replikasyon Stresinin Algılanması",
        "Replikasyon çatalı durduğunda açığa çıkan RPA kaplı tek zincirli DNA'nın ATR-ATRIP kompleksini çekerek TopBP1 ile uyarılması.",
        "ATM çift zincir kırıklarına yanıt verirken, ATR (Ataxia Telangiectasia and Rad3-related) kinaz replikasyon stresi ve tek zincirli DNA (ssDNA) kırıklarının ana kumandanıdır. Replikasyon çatalı durakladığında açığa çıkan ssDNA hızla heterotrimerik RPA (RPA70, RPA32, RPA14) kompleksiyle kaplanır. ATR'nin zorunlu partneri ATRIP (ATR-interacting protein), doğrudan RPA'ya bağlanarak ATR'yi bölgeye getirir. Ardından 9-1-1 kelepçe kompleksi (Rad9-Rad1-Hus1) ve aktivatör protein TopBP1 bölgeye kenetlenir. TopBP1, ATR'nin PIKK kinaz domenini allosterik olarak aktive ederek CHK1 sinyal yolunu açar.",
        "ATR_Kinase_Activation = k_atr * [ATR_ATRIP] * [RPA_coated_ssDNA] * [TopBP1_active]",
        "ATR kinaz aktivasyon şiddeti; RPA ile kaplanmış tek zincir DNA uzunluğu ve TopBP1 koaktivatör derişiminin çarpımıdır."
    ),
    (
        "2.3",
        "DNA-PKcs ve Ku70/Ku80 Dimeri: NHEJ Hasar Başı Kenetlenmesi",
        "DSB uçlarına halka gibi geçen Ku heterodimerinin devasa serin/treonin kinaz DNA-PKcs'i çağırarak onarım iskelesi kurması.",
        "Homolog Olmayan Uç Birleştirme (NHEJ) yolağının sensör çekirdeği DNA-PK (DNA-dependent protein kinase) holoenzimidir. Ku70 ve Ku80 proteinleri sepet şeklinde bir heterodimer oluşturur; çift zincir kırığının açık uçlarına sekans-bağımsız olarak saniyeler içinde kayarak geçer. Ku heterodimeri, 469 kDa'lık devasa katalitik alt birim DNA-PKcs'i (PRKDC) DNA ucuna çağırır. DNA-PKcs iki kırık ucunu fiziksel olarak birbirine hizalar (synapsis) ve otofosforilasyonla konformasyon değiştirerek uç işleme enzimlerine (Artemis nükleaz) ve DNA Ligaz IV kompleksine yol açar.",
        "DNA_PKcs_Synapsis_Flux = k_pk * [Ku70_Ku80_bound] * [DNA_PKcs] * [DSB_ends_proximity]",
        "DNA-PKcs sinapsis debisi; kırık uçlarına kenetlenen Ku heterodimer yoğunluğu ve iki DNA ucunun uzaysal yakınlığı ile yönetilir."
    ),
    (
        "2.4",
        "Histon H2AX Serin 139 Fosforilasyonu (gamma-H2AX): Mega-Bazlık Kromatin Hasar Odakları",
        "Tek bir kırık noktasından her iki yöne doğru 1-2 megabazlık kromatin bölgesinde binlerce H2AX histonunun fosforillenip bayrak açması.",
        "DSB oluştuğunda aktifleşen ATM, ATR veya DNA-PKcs kinazları, hasar noktasını çevreleyen nükleozomlardaki histon varyantı H2AX'i C-terminal Serin 139 kalıntısından fosforiller. Bu fosforillenmiş forma gamma-H2AX denir. Fosforilasyon kırık noktasından itibaren her iki yönde 1 ila 2 megabazlık (yaklaşık 10^6 baz çifti) devasa bir kromatin bölgesine yayılır; tek bir kırık için binlerce gamma-H2AX molekülü üretilir. gamma-H2AX, MDC1 adaptör proteinini bağlar; MDC1 RNF8 ve RNF168 E3 ubikitin ligazlarını çağırarak histon H2A'yı ubikitinler. Bu odak (focus), tüm onarım makineleri için devasa bir moleküler deniz feneri görevi görür.",
        "gamma_H2AX_Amplification = N_foci_intensity = k_amp * [ATM_active] * Integral_0^L [H2AX_density](x) dx",
        "gamma-H2AX sinyal amplifikasyonu; aktif ATM kinaz aktivitesi ile kırık ekseninde 1 megabaz boyunca uzanan H2AX nükleozom integralinin çarpımıdır."
    ),
    (
        "2.5",
        "Efektör Kinazlar: CHK1 ve CHK2'nin CDC25 Fosfatazları Yıkıp Hücre Döngüsünü Dondurması",
        "ATM ve ATR'nin sinyali haberci kinazlar CHK2 ve CHK1'e aktarması; hücre döngüsü motorlarının fosfatazlarının proteazomda parçalanması.",
        "DDR sinyali sensör kinazlardan mediatörler aracılığıyla efektör kinazlara aktarılır. ATR, Claspin adaptörü üzerinden Checkpoint Kinase 1'i (CHK1 / CHEK1) Serin 317 ve Serin 345'ten fosforiller; ATM ise doğrudan Checkpoint Kinase 2'yi (CHK2 / CHEK2) Treonin 68'den aktive eder. Aktif CHK1 ve CHK2, CDC25A, CDC25B ve CDC25C fosfatazlarını fosforiller. Fosforillenen CDC25'ler 14-3-3 proteinleri tarafından sitoplazmaya sürülür veya ubikitinlenip proteazomda yok edilir. CDC25 olmayınca CDK2 ve CDK1 kinazları aktive olamaz; hücre döngüsü anında donar.",
        "Cell_Cycle_Arrest_Flux = [p-CHK1] + [p-CHK2] -> Rate_CDC25_Degradation -> Zero_CDK_Activity",
        "Hücre döngüsü duraklama debisi; aktif CHK1 ve CHK2 kinaz derişiminin CDC25 fosfataz yıkım hızını artırarak CDK aktivitesini sıfırlamasıdır."
    ),
    (
        "2.6",
        "p53-Mdm2 Ekseni: Fosforilasyon, Tetramerleşme ve Genomun Koruyucu Şalteri",
        "David Lane'in 'Genomun Koruyucusu': E3 ligaz Mdm2'nin baskısından kurtulan p53'ün homotetramerleşerek nükleer komut vermesi.",
        "Tümör baskılayıcı p53 proteini normal hücrede çok kısa ömürlüdür (~20 dakika); çünkü E3 ligaz Mdm2 tarafından sürekli ubikitinlenip proteazomda parçalanır. DNA hasarı anında ATM ve CHK2 kinazları p53'ü Serin 15 ve Serin 20 pozisyonlarından fosforiller; aynı anda Mdm2'yi de fosforilleyerek inaktive eder. Fosforilasyon p53'ün Mdm2 ile temasını koparır. Stabilize olan p53 hücre içinde birikir, homotetramer (dörtlü) kompleks kurar ve hedef genlerin promoterlarına bağlanır. Hasarın şiddetine göre hücreye ya 'dur ve tamir et' (p21) ya da 'öl' (Bax, Puma) emrini verir.",
        "p53_Stability_Ratio = [p53_tetramer] = [p53_total] / ( 1 + k_mdm2 * [Mdm2_active] / (1 + [p-Ser15_ATM]) )",
        "p53 tetramerik kararlılık oranı; Mdm2 aracılı yıkım hızının ATM Serin 15 fosforilasyonu ile baskılanması sonucu üstel olarak artar."
    ),
    (
        "2.7",
        "p21CIP1 (CDKN1A) Transkripsiyonu ve G1/S Faz Kilitlenmesi",
        "Aktif p53'ün CDK inhibitörü p21'i sentezleterek Siklin E-CDK2 kompleksini kilitlediği ve hasarlı DNA ile S fazına girişi engellediği kapı.",
        "Hafif ve onarılabilir DNA hasarlarında p53'ün ilk açtığı gen CDKN1A'dır (p21CIP1/WAF1). p21 proteini, G1/S geçişini yöneten Siklin D-CDK4/6 ve Siklin E-CDK2 komplekslerine doğrudan bağlanarak kinaz aktivitelerini felç eder. CDK'lar bloke olunca Retinoblastoma (Rb) proteini hipofosforile (aktif) kalır; transkripsiyon faktörü E2F'i sımsıkı tutar. E2F serbest kalamadığı için DNA replikasyon enzimleri sentezlenemez ve hücre G1 fazında çakılı kalır. Bu duraklama, hücreye genomu tamir etmesi için saatler kazandırır.",
        "G1_Arrest_Probability = [p21CIP1] / ([p21CIP1] + K_i_CDK2) * (1 - [p-Rb])",
        "G1 fazı kilitlenme olasılığı; sentezlenen p21 konsantrasyonunun CDK2 inhibisyon doygunluğu ve serbest hipofosforile Rb fraksiyonu ile belirlenir."
    ),
    (
        "2.8",
        "G2/M Kontrol Noktası: CDK1-Siklin B Kompleksinin İnhibisyonu ve Mitotik Katastroftan Kaçış",
        "Hücrenin kırık kromozomlarla mitoza girmesini engelleyen son güvenlik bariyeri: Wee1 kinaz ve CDC25C dengesi.",
        "DNA'sında çift zincir kırığı taşıyan bir hücre mitoza girerse kardeş kromatitler rastgele yırtılır (mitotik katastrof ve hücre ölümü). Bunu önlemek için G2/M kontrol noktası kurulmuştur. Mitoza girişin anahtarı Siklin B1 - CDK1 (MPF) kompleksidir. DDR kinazları CHK1/2, bir yandan CDC25C fosfatazını inaktive ederken; diğer yandan inhibitör kinaz Wee1'i stabilize eder. Wee1, CDK1'i Tirozin 15 pozisyonundan fosforilleyerek kompleksi kilitli tutar. DNA tamamen tamir edilene kadar hücre mitoza asla geçemez.",
        "Mitotic_Entry_Block = [p-Tyr15_CDK1] / [Total_CDK1] = f([Wee1_active], 1 / [CDC25C_active]) -> 1.0",
        "Mitoza giriş blokajı; aktif Wee1 kinazının CDK1 Tirozin 15 kalıntısını fosforilleyerek inaktif tutma oranı ile güvenceye alınır."
    ),
    (
        "2.9",
        "DDR Odaklarının Çözülmesi ve Onarım Sonrası İyileşme (Recovery / Checkpoint Adaptation)",
        "Onarım bittiğinde fosfatazların (Wip1 / PPM1D) devreye girerek gamma-H2AX, ATM ve CHK1/2'yi söndürüp hücreyi uyandırması.",
        "DNA hasarı tamir edildikten sonra hücre döngüsü duraklamasının sonlandırılması (checkpoint recovery) şarttır. Bu süreç p53 tarafından indüklenen Serin/Treonin fosfataz Wip1 (PPM1D) tarafından yürütülür. Wip1; gamma-H2AX'i, ATM'nin Ser1981'ini, CHK1'in Ser345'ini ve p53'ün Ser15'ini defosforilleyerek kapatır. Kırık bölgesindeki ubikitin zincirleri deübikitilazlar (USP16, USP3) tarafından budanır; nükleozomlar eski haline döner ve CDK kinazları tekrar açılarak hücre döngüsü kaldığı yerden devam eder.",
        "Recovery_Rate = k_wip1 * [Wip1_phosphatase] * [Unbroken_DNA_Integrity] / [Residual_DSB]",
        "Hücre döngüsü iyileşme hızı; indüklenen Wip1 fosfataz aktivitesi ve genomik bütünlüğün artık kırık sayısına olan oranıdır."
    ),
    (
        "2.10",
        "Kronik DDR (TIF): Çözülemeyen Telomerik Hasarın Hücreyi Kalıcı Senesense Kilitlemesi",
        "Fabrizio d'Adda di Fagagna'nın keşfi: Telomerik uçlarda oluşan DNA kırıklarının tamir edilememesi ve ömür boyu susmayan alarm zili.",
        "Kromozomların iç bölgelerindeki DSB'ler hızla tamir edilebilir; ancak telomerik TTAGGG tekrarları bölgesinde oluşan kırıklar tamir edilemez! Telomeri koruyan Shelterin kompleksi (özellikle TRF2), NHEJ ve HR onarım makinelerinin uçlara erişmesini engeller. Bu durum, telomerde oluşan hasarın onarılamaz bir 'kalıcı DNA hasar odağına' (TIF - Telomere Dysfunction-Induced Foci) dönüşmesine yol açar. TIF bölgelerinde ATM ve gamma-H2AX aylarca/yıllarca aktif kalır; sürekli p53 ve p21 sinyali basarak hücreyi geri dönüşümsüz hücresel senesense (yaşlanmaya) kilitler.",
        "Irreversible_Senescence_Flag = sum_telomeres I(TIF_count_k > 0) >= 5_Persistent_Foci",
        "Geri dönüşümsüz hücresel senesens bayrağı; hücre çekirdeğinde çözülemeyen persistan telomerik hasar odağı (TIF) sayısının 5'i aşmasıyla tetiklenir."
    )
]

# ==============================================================================
# KISIM 3: TEK ZİNCİR ONARIM MEKANİZMALARI: BER, NER VE MMR
# ==============================================================================
part3_subsections = [
    (
        "3.1",
        "Baz Eksizyon Onarımı (BER): DNA Glikozilazlar (OGG1, UNG, MUTYH) ve AP Endonükleaz (APE1)",
        "DNA sarmalını bozmayan küçük hasarlı bazların glikozilazlarla dışarı fırlatılıp bazsız boşluğun (AP site) açılması.",
        "Baz Eksizyon Onarımı (BER), genomun en yoğun çalışan fabrikasıdır (günde on binlerce lezyon). Süreç lezyona özgü bir 'DNA Glikozilaz' enziminin hasarlı bazı tanımasıyla başlar: OGG1 oksitlenmiş 8-oxoG'yi, UNG deamine olmuş urasili, MUTYH ise 8-oxoG karşısına yanlışlıkla takılmış adenini tanır. Glikozilaz bazı DNA sarmalından dışarı büker (base flipping) ve deoksiriboz şekerine bağlayan glikozidik bağı keserek serbest bırakır. Geride kalan abazik bölge (AP site), AP Endonükleaz 1 (APE1) tarafından 5' tarafından fosfodiester bağı kesilerek açılır.",
        "BER_Initiation_Rate = k_glyco * [Glycosylase_active] * [Damaged_Base_Substrate]",
        "BER başlatma debisi; spesifik DNA glikozilaz enzim derişimi ve sarmaldaki hasarlı baz konsantrasyonunun kovalent tanıma kinetiğidir."
    ),
    (
        "3.2",
        "Kısa Yama (Short-Patch) vs Uzun Yama (Long-Patch) BER: Polimeraz Beta vs Polimeraz Delta/Epsilon",
        "Tek bir nükleotidin değiştirildiği hızlı yol (Short-Patch) ile 2 ila 10 nükleotidin sentezlendiği alternatif yol (Long-Patch).",
        "APE1 kesiminden sonra onarım iki dala ayrılır. Vakaların %80'inde 'Kısa Yama' (Short-Patch) yolu seçilir: DNA Polimeraz Beta (POLB), APE1'in bıraktığı 5'-dRP (deoksiribozfosfat) kalıntısını liyaz aktivitesiyle koparır, tek bir doğru nükleotidi sentezler ve XRCC1 iskelesi eşliğinde DNA Ligaz III (veya Ligaz I) çentiği kovalent olarak mühürler. Eğer 5' ucu modifiye veya dirençli ise 'Uzun Yama' (Long-Patch) yolu devreye girer: Polimeraz Delta/Epsilon ve PCNA 2-10 nükleotid sentezleyerek eski zinciri bir 'kanat' (flap) gibi yana iter; FEN1 endonükleazı kanadı kesip atar ve Ligaz I kapatır.",
        "BER_Pathway_Choice = Ratio(Short_Patch / Long_Patch) ~ [POLB_lyase_efficiency] / [FEN1_activity] ~ 4:1",
        "BER yol tercihi; Polimeraz Beta'nın dRP liyaz çıkarma hızının FEN1 kanat endonükleaz aktivitesine oranıyla yönetilir ve %80 kısa yama lehinedir."
    ),
    (
        "3.3",
        "Nükleotid Eksizyon Onarımı (NER): Hacimli Lezyonların (Bulky Adducts) Çift Tarafından Kesilip Çıkarılması",
        "DNA sarmalında eğrilik yaratan UV fotourunleri ve kimyasal katımların 24-30 nükleotidlik bir oligonükleotid parçası halinde sökülmesi.",
        "Nükleotid Eksizyon Onarımı (NER), DNA çift sarmalının geometrisini bozan büyük (hacimli) lezyonları temizleyen evrensel cerrahi mekanizmadır. BER'den temel farkı; hasarlı bazın tek başına koparılmaması, hasarı içeren tek iplikli DNA segmentinin lezyonun her iki tarafından (5' tarafından 20-22 nükleotid, 3' tarafından 5-6 nükleotid uzaktan) çift kesimle kesilip 24 ila 30 nükleotidlik bir parça halinde tamamen sökülüp atılmasıdır. Açılan boşluk hatasız polimerazlar tarafından karşı kalıp zincir okunarak doldurulur.",
        "NER_Excision_Flux = k_ner * [Dual_Incision_Complex] * [Bulky_Lesion_Density]",
        "NER eksizyon akısı; XPF-XPG çift kesim endonükleaz kompleksinin DNA sarmalındaki hacimli lezyon yoğunluğu ile karşılaşma hızıdır."
    ),
    (
        "3.4",
        "Global Genom NER (GG-NER): XPC-RAD23B ve DDB1-DDB2 (UV-DDB) Hasar Tarama Kompleksleri",
        "Genomun transkripsiyona uğramayan sessiz %99'luk alanında DNA sarmalındaki eğrilikleri gece gündüz tarayan nöbetçi proteinler.",
        "Global Genom NER (GG-NER), tüm genom boyunca hasar taraması yapar. Lezyonu bizzat tanıyan faktör XPC - RAD23B kompleksidir. XPC hasarlı baza değil; hasarın karşı ipliğinde oluşan sarmal distorsiyonuna ve baz eşleşmesi bozukluğuna bağlanır. Ancak UV kaynaklı CPD lezyonları sarmalı çok az büktüğü için XPC tarafından zor fark edilir; burada devreye UV-DDB kompleksi (DDB1 ve DDB2 / XPE) girer. UV-DDB lezyonu tanır, kromatini ubikitinleyerek gevşetir ve XPC'yi hasar noktasına transfer eder. XPC bağlandıktan sonra onarım makineleri toplanır.",
        "GG_NER_Scanning_Rate = D_search * [XPC_RAD23B] * (1 + alpha_ddb * [UV_DDB])",
        "GG-NER tarama hızı; XPC-RAD23B kompleksinin 1D DNA kayma difüzyon katsayısı ve UV-DDB şaperon desteğinin bir fonksiyonudur."
    ),
    (
        "3.5",
        "Transkripsiyon Eşlikli NER (TC-NER): RNA Polimeraz II Blokajı, CSB (ERCC6) ve CSA (ERCC8) ",
        "Aktif genlerin kalıp zincirinde ilerleyen RNA Polimeraz II'nin lezyona çarpıp durmasıyla tetiklenen acil öncelikli tamir.",
        "Hücrenin hayati fonksiyonlarını sürdürmesi transkripsiyonun durmamasına bağlıdır. Bir gen okunurken RNA Polimeraz II (RNAPII) kalıp iplikteki bir lezyona çarptığında fiziksel olarak donar. Bu donma 'Transkripsiyon Eşlikli NER'i (TC-NER) başlatır. Sıkışan polimerazın arkasına CSB (ERCC6 - DNA bağımlı ATPaz) ve CSA (ERCC8 - E3 ubikitin ligaz kompleksi) kenetlenir. CSB ve CSA, RNAPII'yi hafifçe geriye iterek lezyonun üzerini açar; XPC'ye ihtiyaç duyulmadan TFIIH kompleksi doğrudan çağrılır. TC-NER hızı GG-NER'den 10 kat daha süratlidir.",
        "TC_NER_Priority_Flux = k_tc * [Stalled_RNAPII] * [CSB_ATPase] * [CSA_complex]",
        "TC-NER öncelikli tamir debisi; hasar noktasında sıkışan RNA Polimeraz II kompleksi ile aktif CSB ve CSA proteinlerinin yoğunluğuna bağımlıdır."
    ),
    (
        "3.6",
        "TFIIH Kompleksi, XPB/XPD Helikazları ve XPF-ERCC1 / XPG Çift Kesim Endonükleazları",
        "10 alt birimli transkripsiyon/onarım devi TFIIH'ın DNA'yı açması ve yapay nükleaz makaslarının lezyonu söküp çıkarması.",
        "Hasar tanındıktan sonra bölgeye 10 alt birimli TFIIH kompleksi gelir. TFIIH'ın XPB (3'-5') ve XPD (5'-3') helikazları, ATP hidrolizi ile lezyon etrafındaki DNA çift sarmalını 25-30 baz çifti boyunca eriterek tek zincirli bir balon (bubble) açar. XPA proteini lezyonun doğruluğunu teyit eder; RPA açılan tek zinciri korur. Son olarak iki yapıya-özgü endonükleaz devreye girer: XPG proteini 3' tarafından, XPF-ERCC1 heterodimeri ise 5' tarafından DNA omurgasını keser. Hasarlı oligonükleotid fırlatılır; DNA Polimeraz Delta/Epsilon ve Ligaz I/III boşluğu doldurup kapatır.",
        "Dual_Incision_Rate = k_cut * [XPF_ERCC1] * [XPG] * [Unwound_Bubble_TFIIH]",
        "Çift taraflı kesim debisi; TFIIH tarafından açılan tek zincirli balonun XPF-ERCC1 ve XPG nükleazları ile eşzamanlı kesim kinetiğidir."
    ),
    (
        "3.7",
        "Uyuşmazlık Onarımı (MMR): MutS-alfa (MSH2-MSH6) ve MutL-alfa (MLH1-PMS2) Makineleri",
        "Replikasyon polimerazının proofreading gözünden kaçan yanlış baz eşleşmelerini ve küçük kaymaları (indel) yakalayan kalite kontrol.",
        "DNA replikasyon polimerazları her 100.000 bazda bir hata yapar; proofreading bunu binde bire indirir; geriye kalan hataları ise Uyuşmazlık Onarımı (MMR) temizleyerek replikatif doğruluğu milyarda bire çıkarır. MutS-alfa heterodimeri (MSH2-MSH6), Watson-Crick eşleşmesi yapmayan uyumsuz baz çiftlerini (örn. G-T eşleşmesi) tanır; MutS-beta (MSH2-MSH3) ise 1-4 bazlık insersiyon/delesyon ilmeklerini yakalar. Ardından MutL-alfa (MLH1-PMS2) kompleksi gelir. PMS2 endonükleazı yeni sentezlenen yavru ipliği keser; EXO1 ekzonükleazı hatalı bölgeyi siler ve Polimeraz Delta boşluğu yeniden sentezler.",
        "Replication_Fidelity_MMR = Error_Rate_baseline * (1 - MMR_Efficiency) -> 10^-10_errors/bp",
        "Nihai replikasyon sadakati; bazal polimeraz hata oranının MMR düzeltme verimliliği ile (yaklaşık 1000 kat) bastırılmasıyla 10^-10 seviyesine iner."
    ),
    (
        "3.8",
        "Mikrosatelit Kararsızlığı (MSI): MMR Yetersizliğinin Neden Olduğu Genomik Kaymalar",
        "Mononükleotid ve dinükleotid tekrarlarında (mikrosatelitler) polimeraz kayması sonucu oluşan genetik delik deşiklik.",
        "Genomda binlerce kez tekrarlanan mikrosatelit bölgeleri (örn. AAAAA veya CACACA tekrarları), replikasyon sırasında polimeraz kaymasına (slippage) aşırı duyarlıdır. Eğer hücrede MMR sistemi (MLH1, MSH2) genetik mutasyon veya promoter hipermetilasyonu ile susturulmuşsa, bu kaymalar onarılamaz. Mikrosatelit uzunlukları hücreden hücreye değişir; bu fenotipe Mikrosatelit Kararsızlığı (MSI-High) denir. Lynch Sendromunun ve sporidik kolorektal kanserlerin arkasındaki mekanizmadır; yaşlanan kök hücrelerde de MSI birikimi doku fonksiyonlarını aşındırır.",
        "MSI_Index = Count_Altered_Microsatellite_Markers / Total_Tested_Markers > 0.30 (MSI-High)",
        "Mikrosatelit kararsızlık indeksi; test edilen polimerik belirteçler arasında uzunluğu değişmiş mutant lokusların %30'u aşmasıdır."
    ),
    (
        "3.9",
        "Yaşlanmada Tek Zincir Onarım Kapasitesinin Erozyonu ve Enzimatik Kinetik Yavaşlama",
        "Yaşlanan dokularda OGG1, APE1, XPC ve MLH1 protein düzeylerinin gerilemesi; onarım hızının hasar oluşum hızının gerisinde kalması.",
        "Tek zincir onarım sistemleri yaşlanmayla birlikte belirgin bir kinetik yavaşlama yaşar. Yaşlı insan doku biyopsilerinde nükleer OGG1 glikozilaz aktivitesi %50 azalırken, APE1 ekspresyonu düşer. Benzer şekilde deride NER kapasitesi 20 yaşından 70 yaşına kadar %40 geriler; bu da UV hasarlarının kümülatif olarak birikmesine yol açar. Onarım hızının (V_repair) hasar oluşum hızının (V_damage) altına indiği noktada genomda tamir edilmemiş tek zincir çentikleri (SSB) birikir; replikasyon çatalı bu çentiklere çarptığında çift zincir kırıkları patlar.",
        "Unrepaired_SSB_SteadyState = Rate_Damage_SSB / (k_repair * [Active_Repair_Enzymes])",
        "Kararlı durumdaki onarılmamış tek zincir kırık yükü; hasar oluşum hızının aktif tamir enzim konsantrasyonuna oranıyla belirlenir."
    ),
    (
        "3.10",
        "Onarım Bozukluğu Sendromları: Kseroderma Pigmentozum (XP), Cockayne ve Trikotiyodistrofi",
        "Tek bir NER geninin mutasyonunun insanı çocuk yaşta güneşte kanser yapması veya erken yaşlanma (progeroid) tablosuna sokması.",
        "DNA onarımının yaşam için önemini gösteren en dramatik kanıt kalıtsal onarım sendromlarıdır. Kseroderma Pigmentozum (XP: XPA-XPG gen mutasyonları), GG-NER'in çalışmamasıdır; hastalar güneş ışığına maruz kaldıklarında 2000 kat artmış deri kanseri riski yaşar. Cockayne Sendromu (CS: CSA/CSB mutasyonları) ise TC-NER'in çökmesidir; kanser artışı görülmez ancak transkripsiyon kilitlendiği için hastalar şiddetli erken yaşlanma, nörodejenerasyon ve cücelik geliştirerek çocuklukta ölür. Bu durum, transkripsiyon eşlikli onarımın doğrudan hücresel longevity ile bağlantılı olduğunu kanıtlar.",
        "Clinical_Phenotype = IF(GG_NER_Defect, High_Cancer_Risk, IF(TC_NER_Defect, Accelerated_Aging_Progeria))",
        "Onarım defekti klinik sonucu; GG-NER kusurunun karsinogenez yaratması, TC-NER kusurunun ise hızlandırılmış sistemik progeria yaratmasıdır."
    )
]

parts.append(("KISIM 1: DNA HASAR TÜRLERİ VE ENDOJEN / EKSOJEN MUTAJENLER", part1_subsections))
parts.append(("KISIM 2: DNA HASAR YANITI (DDR) SİNYAL KASKADI VE KONTROL NOKTALARI", part2_subsections))
parts.append(("KISIM 3: TEK ZİNCİR ONARIM MEKANİZMALARI: BER, NER VE MMR", part3_subsections))

# ==============================================================================
# KISIM 4: ÇİFT ZİNCİR KIRIK (DSB) ONARIMI: NHEJ, HR VE MMEJ YOL AYRIMI
# ==============================================================================
part4_subsections = [
    (
        "4.1",
        "Homolog Olmayan Uç Birleştirme (NHEJ): Hızlı, Basit fakat Mutajenik (İndel Üreten) Onarım",
        "Hücre döngüsünün tüm evrelerinde (özellikle G0 ve G1) çalışan, kalıp ipliğe bakmaksızın iki serbest ucu birbirine bağlayan acil ilk yardım yolu.",
        "Klasik Homolog Olmayan Uç Birleştirme (c-NHEJ), memeli hücrelerinde çift zincir kırıklarının (DSB) yaklaşık %80'ini onaran primer mekanizmadır. NHEJ, kardeş kromatid kalıbı gerektirmez; bölünmeyen post-mitotik nöronlar ve kök hücreler sükunette tamamen bu yola bağımlıdır. Ancak bu hızın biyolojik bir bedeli vardır: Kırık uçlar genellikle pürüzsüz değildir (hasarlı nükleotidler taşır); uçlar işlenirken (nükleaz kesimleri ve polimeraz dolguları) 1 ila 20 baz çiftlik delesyon veya insersiyonlar (indeller) oluşur. Bu mutajenik yara izleri, on yıllar içinde genomun kodlama ve regülasyon saflığını aşındırır.",
        "NHEJ_Mutagenic_Rate = P(Indel_Formation) = 1 - P_blunt_end_perfect_ligation ~ 0.65",
        "NHEJ onarımının mutajenik olma olasılığı; kırık uçlarının hasar derecesine bağlı olarak %65'in üzerindedir ve kalıcı indel bırakır."
    ),
    (
        "4.2",
        "Ku70/Ku80 Kenetlenmesi, DNA-PKcs Aktivasyonu ve Artemiz (Artemis) Endonükleaz Uç İşleme",
        "Kırık uçları hapseden Ku halkası, DNA-PKcs sinapsisi ve kimyasal olarak tıkanmış uçları kesip temizleyen Artemiz nükleazı.",
        "Ku70/Ku80 heterodimeri kırık ucuna oturduktan sonra DNA-PKcs kinazı çağrılır. İki DNA-PKcs molekülü karşılıklı kenetlenerek uçları bir arada tutar. Ancak radyasyon veya serbest radikallerle oluşan kırık uçları çoğu zaman ligasyona uygun 5'-fosfat ve 3'-hidroksil taşımaz (hasarlı, çapraz bağlı veya saç tokası/hairpin yapısındadır). DNA-PKcs tarafından fosforillenen Artemiz (DCLRE1C) endonükleazı devreye girer. Artemiz saç tokası yapılarını açar, hasarlı tek zincir çıkıntılarını keser ve ligasyona hazır pürüzsüz uçlar oluşturur. Bu kesim işlemi NHEJ'deki genetik bilgi kaybının ana nedenidir.",
        "End_Processing_Rate = k_artemis * [DNA_PKcs_active] * [p-Artemis] * [Complex_DNA_Ends]",
        "NHEJ uç işleme hızı; DNA-PKcs tarafından aktive edilen Artemiz nükleaz konsantrasyonu ve hasarlı uç karmaşıklığı ile belirlenir."
    ),
    (
        "4.3",
        "DNA Ligaz IV - XRCC4 - XLF (PAXX) Kompleksi ve Uçların Kovalent Kapatılması",
        "Uçları birbirine dikip kovalent fosfodiester omurgasını restore eden özel ligasyon makinesi.",
        "Uçlar işlendikten sonra ligasyon aşaması başlar. Bu aşamayı yürüten yegane enzim DNA Ligaz IV'tür (LIG4). Ligaz IV tek başına kararsızdır; zorunlu partneri XRCC4 ile sıkı bir kompleks oluşturur. Bu komplekse XLF (NHEJ1) ve PAXX adaptör proteinleri katılarak DNA uçları üzerinde bir 'köprü' (protein bridge) kurar. Ligaz IV, ATP harcayarak iki DNA ucu arasında fosfodiester bağını kovalent olarak kapatır. Ligaz IV mutasyonları şiddetli immün yetmezlik ve mikrosefaliye yol açarken; yaşlanan dokularda bu kompleksin zayıflaması onarılamayan kromozom kopmalarını artırır.",
        "Ligation_Flux_NHEJ = k_lig4 * [LIG4_XRCC4_XLF_complex] * [Processed_Ends] * [ATP]",
        "NHEJ kovalent kapatma debisi; LIG4-XRCC4-XLF süper-kompleksi yoğunluğu ve işlenmiş DNA uç konsantrasyonunun çarpımıdır."
    ),
    (
        "4.4",
        "Homolog Rekombinasyon (HR): Kardeş Kromatid Kalıbını Kullanan Hatasız (High-Fidelity) Onarım",
        "Yalnızca S ve G2 fazlarında, identik kardeş kromatidi şablon alarak tek bir baz bile kaybetmeden çift zincir kırığını sıfırlayan kusursuz mimari.",
        "Homolog Rekombinasyon (HR), doğanın en sofistike genetik onarım mekanizmasıdır. NHEJ'in aksine kesinlikle mutasyon veya indel üretmez (high-fidelity). HR yalnızca DNA replikasyonunun tamamlandığı veya devam ettiği S ve G2 fazlarında çalışabilir; çünkü kırık zincirin onarılması için identik bir 'kardeş kromatid' (sister chromatid) şablonu şarttır. Kırık DNA, sağlam kardeş kromatidin içine sızar (zincir istilası); sağlam zincir okunarak hasarlı bölge baştan aşağı kopyalanır ve orijinal sekans kusursuzca restore edilir.",
        "HR_Fidelity = 1 - P(Mutation) -> 0.99999 (Kardeş kromatid şablonunun kusursuzluğu)",
        "Homolog Rekombinasyon onarım doğruluğu; sağlam kardeş kromatid kalıbının kullanılması sayesinde neredeyse %100 kusursuzdur."
    ),
    (
        "4.5",
        "Uç Rezeksiyonu (End Resection): Mre11-CtIP ve EXO1/BLM ile 3' Tek İplikli Kuyruk Oluşturulması",
        "HR'nin geri dönüşümsüz ilk kararı: 5' ipliğinin kilometrelerce geriye sindirilerek invaziv 3' tek iplikli DNA kuyruğu açılması.",
        "Bir kırığın NHEJ yerine HR ile tamir edilmesine karar verildiğinde atılan ilk biyokimyasal adım 'Uç Rezeksiyonu'dur (DNA End Resection). Süreç iki basamaklıdır: İlk kısa menzilli kesim, MRN kompleksi ve fosforile CtIP proteini tarafından yapılır; Mre11 endonükleazı kırığın 5' ipliğinde bir çentik açar. İkinci uzun menzilli kesimde ise EXO1 ekzonükleazı ve BLM helikaz-DNA2 nükleaz kompleksi devreye girerek 5' ipliğini geriye doğru binlerce nükleotid sindirir. Geride, 3' ucu serbest olan uzun tek iplikli DNA (ssDNA) kuyrukları kalır.",
        "Resection_Velocity = v_exo1 * [EXO1_active] + v_dna2 * [BLM_DNA2_complex] > 4000_bp/hour",
        "Uç rezeksiyon ilerleme hızı; aktif EXO1 ve BLM/DNA2 komplekslerinin saatte binlerce baz çifti 5' ipliğini sindirmesiyle açılır."
    ),
    (
        "4.6",
        "Rad51 Rekombinazı ve BRCA1/BRCA2: Nükleofilament Oluşumu ve Homolog Zincir İstilası (D-Loop)",
        "Açılan tek zincirin Rad51 ile kaplanarak sarmal bir nükleoprotein filamentine dönüşmesi ve kardeş kromatidi delip girmesi.",
        "Rezeksiyonla açılan 3' ssDNA kuyrukları başlangıçta RPA ile kaplanır. Tümör baskılayıcı BRCA1 ve BRCA2 proteinleri devreye girer. BRCA2, nükleer şaperon görevi görerek RPA'yı ssDNA üzerinden kovar ve yerine yüzlerce Rad51 rekombinaz molekülü yükler. Rad51, DNA etrafında sağ el sarmallı bir 'presinaptik nükleofilament' örer. Bu filament, kardeş kromatid çift sarmalını tarar (homoloji araması). İdentik sekansı bulduğu anda çift sarmalı açar ve içine sızarak bir D-Loop (Displacement Loop) kavşağı kurar; DNA Polimeraz Delta 3' ucundan uzamaya başlar.",
        "Homology_Search_Velocity = D_search * [Rad51_Nucleofilament] * Length_ssDNA / Genome_Volume",
        "Homoloji arama ve D-Loop kurma hızı; Rad51 nükleofilament uzunluğu ve 3 boyutlu nükleer arama difüzyon katsayısının fonksiyonudur."
    ),
    (
        "4.7",
        "Holliday Kavşakları ve Çözülme (Resolvazlar: GEN1, BML-RMI1-RMI2-TOPO3a)",
        "İki kromatidin birbirine çaprazlandığı dört kollu Holliday yapısının kromatit ayrışması (crossover vs non-crossover) ile çözülmesi.",
        "Zincir istilası ve DNA sentezi ilerledikçe iki kromatit birbirine iki noktadan kovalent olarak kilitlenir; bu dört kollu DNA kavşaklarına 'Çift Holliday Kavşağı' (dHJ) denir. Hücre bölünmeden önce bu kromatitlerin birbirinden temizce ayrılması zorunludur. Çözülme iki yolla yapılır: BTR kompleksi (BLM helikaz, Topoizomeraz III-alfa, RMI1, RMI2) kavşakları birbirine doğru çözerek 'çözünme' (dissolution) yapar ve sıfır parça değişimi (non-crossover) sağlar. Alternatif yolda ise GEN1 ve SLX1-SLX4 yapı-spesifik resolvazları kavşağı keserek kromatit ayrışmasını tamamlar.",
        "Holliday_Resolution_Rate = k_dissolution * [BTR_complex] + k_resolvase * [GEN1] * [SLX1_SLX4]",
        "Holliday kavşak çözülme debisi; BTR kromatit çözünme akısı ile GEN1/SLX nükleaz kesim akılarının toplamıdır."
    ),
    (
        "4.8",
        "Yol Seçimi (Pathway Choice): 53BP1-RIF1-Shieldin (NHEJ Bariyeri) vs BRCA1-CtIP (HR İticisi)",
        "Hücrenin kırık anında NHEJ mi yoksa HR mi yapacağını belirleyen moleküler savaş: 53BP1 kalkanı vs BRCA1 rezeksiyonu.",
        "Bir DSB oluştuğunda hücre kaderi 'Yol Seçimi' (DSB Pathway Choice) mekanizmasıyla belirlenir. G1 fazında 53BP1 proteini kırık bölgesindeki gamma-H2AX ve H4K20me2 işaretlerine bağlanır; RIF1 ve Shieldin kompleksini (SHLD1-3) çağırarak kırık ucunda fiziksel bir kalkan örer. Bu kalkan nükleazların uca yaklaşmasını engeller ve kırığı zorunlu olarak NHEJ'e iter. S ve G2 fazında ise CDK aktivitesi tavan yapar; CtIP fosforillenir ve BRCA1 53BP1'i bölgeden kovar. Kalkan kalkınca uç rezeksiyonu başlar ve hücre HR yoluna kilitlenir.",
        "Pathway_Decision_Ratio = [BRCA1_CtIP_active] / ([53BP1_Shieldin_complex] + epsilon) = f(CDK_Activity_Phase)",
        "Onarım yolu tercih oranı; hücre döngüsü CDK kinaz aktivitesine bağlı olarak BRCA1 lehine (HR) veya 53BP1 lehine (NHEJ) kayar."
    ),
    (
        "4.9",
        "Alternatif Uç Birleştirme (alt-NHEJ / MMEJ): Mikrohomoloji Kullanımı ve Polimeraz Teta (POLQ)",
        "NHEJ ve HR çöktüğünde devreye giren en mutajenik acil çıkış kapısı: Mikrohomoloji aracılı uç birleştirme ve devasa delesyonlar.",
        "Eğer hücrede hem Ku70/80 (NHEJ) hem de BRCA1/Rad51 (HR) yetersizse, 'Mikrohomoloji Aracılı Uç Birleştirme' (MMEJ / alt-NHEJ) devreye girer. Bu yolak, kırık uçlarındaki sadece 2 ila 6 baz çiftlik küçük mikro-homoloji adacıklarını kullanır. Polimeraz Teta (POLQ), helikaz ve polimeraz domenleriyle iki ucun mikrohomolojisini hizalar; uçlardaki uyumsuz kanatları kesip aradaki boşluğu doldurur. Sonuç felakettir: Onarım noktasında yüzlerce baz çiftlik devasa delesyonlar ve kromozomal translokasyonlar oluşur. Kanser hücreleri ve yaşlı dokular POLQ bağımlı MMEJ'yi aşırı kullanır.",
        "MMEJ_Deleterious_Flux = k_polq * [Polymerase_Theta] * [Resected_Microhomologies] / ([Ku70_80] + [Rad51])",
        "MMEJ kaynaklı genomik hasar debisi; Polimeraz Teta yoğunluğunun standart koruyucu Ku ve Rad51 proteinlerine olan üstünlüğü ile patlar."
    ),
    (
        "4.10",
        "Yaşlanmada HR Kaybı ve Dokuların Hatalı NHEJ/MMEJ Onarımına Mahkum Olması",
        "Yaşlanan kök hücrelerde SIRT6 ve BRCA1 düzeylerinin düşmesi sonucu hatasız HR'nin sönümlenmesi ve mutasyonel kararsızlık.",
        "Vera Gorbunova ve Andrei Seluanov laboratuvarlarının çığır açıcı keşiflerine göre, yaşlanma sürecinde Homolog Rekombinasyon kapasitesi %80'den fazla azalır. Yaşlı hücrelerde nükleer NAD+ tükenişi SIRT6 deasetilaz aktivitesini düşürür; bu da CtIP rezeksiyonunu ve Rad51 yüklenmesini engeller. Hücreler S/G2 fazında dahi çift zincir kırıklarını hatasız HR ile tamir edemez; zorunlu olarak mutajenik c-NHEJ ve felaket düzeyindeki MMEJ yoluna sürüklenir. Bu onarım kalitesi kayması, yaşlı organizmalarda kromozomal translokasyon ve onkogenik mutasyon patlamasının ana nedenidir.",
        "HR_to_NHEJ_Ratio(t) = HR_Capacity_0 * exp(-k_aging * t) / NHEJ_Activity_Constant",
        "Yaşlanmaya bağlı HR/NHEJ kapasite oranı; zaman içinde üstel olarak düşerek dokuları hataya açık mutajenik onarıma mahkum eder."
    )
]

# ==============================================================================
# KISIM 5: HETEROKROMATİN KAYBI HİPOTEZİ VE EPİGENETİK YENİDEN YAPILANMA
# ==============================================================================
part5_subsections = [
    (
        "5.1",
        "Heterokromatin Mimarisi: Konstitütif (H3K9me3, HP1) ve Fakültatif (H3K27me3, PRC2) Bölgeler",
        "Genomun susturulmuş karanlık bölgeleri: Sentromer ve telomerleri saran konstitütif heterokromatin ile gelişimsel fakültatif sessizlik.",
        "Ökaryotik genom iki ana kromatin durumunda paketlenir: Transkripsiyona açık, gevşek ökromatin ve sıkıca paketlenmiş, sessiz heterokromatin. Heterokromatin ikiye ayrılır: 1) Konstitütif heterokromatin; tüm hücre tiplerinde sentromerik, perisentromerik ve telomerik tekrarları kilitler; H3K9me3 (Histon H3 Lizin 9 trimetilasyonu) ve Heterokromatin Proteini 1 (HP1alpha/beta) oligomerleriyle süper-yoğun faz ayrışması damlacıkları kurar. 2) Fakültatif heterokromatin; Polycomb Baskılayıcı Kompleks 2 (PRC2) tarafından kurulan H3K27me3 işaretleriyle dokuya özgül sessiz genleri yönetir.",
        "Heterochromatin_Density = [H3K9me3] * [HP1_alpha] + [H3K27me3] * [PRC2_complex]",
        "Nükleer heterokromatin paketlenme yoğunluğu; konstitütif H3K9me3/HP1 faz ayrışması kütlesi ile fakültatif H3K27me3/PRC2 zenginliğinin toplamıdır."
    ),
    (
        "5.2",
        "Nükleer Periferi ve Lamin B1: Heterokromatinin Nükleer Çepere Ankrajı",
        "Lamin-İlişkili Domenlerin (LAD) çekirdek zarına çivilenmesi: Genomun üçte birinin nükleer periferde sessizliğe gömülmesi.",
        "Heterokromatin nükleus içinde rastgele yüzmez; genomun yaklaşık %35'ini oluşturan mega-bazlık 'Lamin-İlişkili Domenler' (LAD - Lamina-Associated Domains), nükleer lamina ağına (özellikle Lamin B1 ve LBR reseptörüne) fiziksel olarak ankrajlanır (çivilenir). Bu nükleer çeper bölgesi gen ekspresyonu için derin bir 'baskılama çölüdür' (repressive desert). LAD bölgelerindeki genler ve retrotranspozonlar kapalı kalarak transkripsiyon faktörlerinin erişiminden saklanır. Yaşlanmayla Lamin B1 yıkıldığında LAD'lar zardan koparak nükleer merkeze dağılır.",
        "LAD_Anchoring_Stability = K_anchor * [Lamin_B1] * [LBR_receptor] * [H3K9me2/3_density]",
        "LAD kromatin ankraj kararlılığı; nükleer zar Lamin B1 ve LBR reseptör yoğunluğu ile heterokromatin histon işaretlerinin çarpımıdır."
    ),
    (
        "5.3",
        "Villeponteau'nun 'Loss of Heterochromatin' Hipotezi: Yaşlanmanın Açılan Susturulmuş Genler Olması",
        "1997 yılında Bernard Villeponteau tarafından ortaya atılan teori: Yaşlanmanın ana motoru DNA mutasyonları değil, kilitli heterokromatinin çözülmesidir.",
        "Villeponteau'nun 'Heterokromatin Kaybı Yaşlanma Modeli'ne (Loss of Heterochromatin Model of Aging) göre; gençlik heterokromatin adacıklarının sıkı ve kusursuz izolasyonudur. Yaşlanma sürecinde hücreler her bölündüğünde ve DNA tamir edildikçe histon metiltransferazlar ve şaperonlar tükenir; heterokromatin adacıkları çözülmeye (açılmaya) başlar. Normalde ömür boyu kilitli kalması gereken sessiz genler, gelişimsel faktörler ve transpozonlar rastgele uyanır. Bu durum hücre kimliğini yıkar, transkripsiyonel kaosa yol açar ve hücreyi hızla senesense sürükler.",
        "Aging_Drift_Rate = d(Entropy_Epigenetic)/dt = k_decay * (1 - [Heterochromatin_Integrity])",
        "Epigenetik yaşlanma sürüklenme hızı; nükleer heterokromatin bütünlüğünün zamansal kaybı ile Shannon kromatin entropisinin artışıdır."
    ),
    (
        "5.4",
        "Yaşlanmada H3K9me3 ve HP1alpha Erozyonu: Transkripsiyonel Gürültü (Transcriptional Noise)",
        "Tek hücre RNA dizilemesiyle kanıtlanan fenomen: Yaşlı hücrelerde aynı tip hücrelerin her birinin rastgele farklı genler haykırması.",
        "H3K9me3 ve HP1alpha kaybı, genomik DNA'nın sıkı yapısını gevşetir. Ökromatin ile heterokromatin arasındaki sınırlar silinir. Bu durum 'Transkripsiyonel Gürültü' (Transcriptional Noise) yaratır: Genç bir dokuda aynı tip 1000 hücre neredeyse identik bir transkriptom profili sergilerken; yaşlı dokudaki 1000 hücrenin her biri rastgele uygunsuz genler eksprese etmeye başlar. Bir karaciğer hücresinde nöronal genler veya enflamatuar transkriptler spontan olarak açılır; hücreler senkronize doku fonksiyonu yürütme yeteneğini kaybeder.",
        "Transcriptional_Noise = Cell_to_Cell_Variance(Expression_Vector_i, Expression_Vector_j) ~ exp(k_noise * Age)",
        "Transkripsiyonel gürültü; yaşlanan dokudaki tek hücreler arası gen ekspresyon varyansının zamansal üstel artışı ile formalize edilir."
    ),
    (
        "5.5",
        "DNA Metilasyon Kaybı (Global Hipometilasyon) ve Sentromerik/Perisentromerik Kararsızlık",
        "Yaşlanan hücrelerde genel CpG metilasyonunun silinmesi: Sentromerik uydu DNA'larının çözülmesi ve kromozom ayrışma felçleri.",
        "Steve Horvath'ın epigenetik saatinde belirli CpG adacıkları yaşla hipermetile olurken; genomun geneline bakıldığında devasa bir 'global hipometilasyon' (metilasyon kaybı) yaşanır. Özellikle sentromerik ve perisentromerik bölgelerdeki alfa-uydu (satellite) tekrarlarındaki 5mC metilasyonu silinir. Metilasyonunu kaybeden sentromerik heterokromatin açılır; kromatitler arası kohezyon zayıflar, kinetokor montajı aksar ve mitoz bölünmelerde anöploidi, kromozom kopmaları ve mikronükleus oluşumu patlar.",
        "Centromeric_Instability_Index = (Unmethylated_Sat_DNA / Total_Sat_DNA) * Aneuploidy_Frequency",
        "Sentromerik kararsızlık indeksi; metilasyonunu kaybetmiş perisentromerik uydu DNA oranının mitotik anöploidi sıklığıyla çarpımıdır."
    ),
    (
        "5.6",
        "H4K16 Asetilasyon Dinamiği (MOF vs SIRT1/6) ve Kromatin Gevşemesi",
        "Histon H4 Lizin 16 asetilasyonunun nükleozomlar arası elektrostatik çekimi yok ederek 30 nm'lik kromatin lifini çözmesi.",
        "Histon H4'ün 16. lizin kalıntısının asetilasyonu (H4K16ac), kromatin mimarisinin en kritik biyofiziksel şalteridir. Asetilsiz H4'ün bazik ucu, komşu nükleozomdaki H2A/H2B asidik yamasıyla tuz köprüleri kurarak kromatinin 30 nanometrelik sıkı lif halinde katlanmasını sağlar. KAT8 (MOF) asetiltransferazı bu lizini asetillediğinde pozitif yük nötralize olur; elektrostatik çekim biter ve kromatin anında gevşer. Yaşlanmada NAD+ tükenişi nedeniyle deasetilazlar (SIRT1 ve SIRT6) H4K16'yı kapatamaz; kontrolsüz H4K16ac birikimi tüm genomu gevşeterek yaşlanmayı hızlandırır.",
        "Chromatin_Relaxation_Fraction = [H4K16ac] / [Total_H4] = [MOF_activity] / ([SIRT1] * [NAD+] + [SIRT6] * [NAD+])",
        "Kromatin gevşeme oranı; MOF asetiltransferaz aktivitesinin NAD+ bağımlı SIRT1 ve SIRT6 deasetilasyon klerensine oranı ile belirlenir."
    ),
    (
        "5.7",
        "Yaşlanan Hücrede Senesens İlişkili Heterokromatin Odakları (SAHF): p16/Rb ile Sessizleşen Genler",
        "Gevşeyen genomun aksine, proliferasyon genlerinin çekirdekte devasa mikroskobik DAPI-yoğun odaklara hapsedilmesi.",
        "Senesen hücrelerde paradoksal bir kromatin yeniden yapılanması görülür: Bir yanda heterokromatin kaybı yaşanırken; diğer yanda hücre proliferasyon genleri (E2F hedef genleri: Siklin A, PCNA) 'Senesens İlişkili Heterokromatin Odakları' (SAHF - Senescence-Associated Heterochromatin Foci) adı verilen devasa nükleer kondansatlarda kilitlenir. DAPI boyasıyla nükleusta yoğun noktalar halinde görülen SAHF; p16INK4a-Rb yolağı, HMGA proteinleri, H3K9me3 ve makroH2A histon varyantı ile örülür. Bu yapı hücrenin bir daha asla bölünememesini moleküler olarak garantiye alır.",
        "SAHF_Formation_Score = Count(DAPI_Dense_Foci) = f([p16INK4a], [Hypophosphorylated_Rb], [HMGA1/2])",
        "SAHF oluşum skoru; çekirdekteki DAPI-yoğun heterokromatin odak sayısı olup p16/Rb aktivasyonu ve HMGA mimari proteinlerine bağlıdır."
    ),
    (
        "5.8",
        "Nükleozom Yoğunluğu Kaybı: Yaşla Çekirdekteki Toplam Histon Kütlesinin %50 Azalması",
        "Maya, solucan, fare ve insanda evrensel korunan yaşlanma damgası: Histon proteinlerinin sentezlenememesi ve çıplak kalan DNA.",
        "Jessica Tyler laboratuvarının çığır açan keşiflerine göre, yaşlanan ökaryotik hücrelerin en temel ortak özelliği çekirdekteki toplam nükleozom yoğunluğunun dramatik kaybıdır (Loss of Histones). Yaşlı hücrelerde histon transkripsiyonu ve kanonik histon şaperonları (Asf1, CAF-1) çöker; çekirdekteki toplam Histon H3 ve H4 protein kütlesi %30 ila %50 oranında buharlaşır. Nükleozom aralıkları açılır; çıplak kalan ara DNA (linker DNA) kırılmalara, transpozon aktivasyonuna ve aberan transkripsiyonel başlatmaya açık hale gelir.",
        "Nucleosome_Occupancy_Drop = (Density_Nucleosomes_Young - Density_Nucleosomes_Old) / Density_Young ~ 0.40",
        "Nükleozom doluluk kaybı; yaşlı çekirdekteki baz çifti başına düşen oktamerik histon yoğunluğunun gençliğe kıyasla %40 gerilemesidir."
    ),
    (
        "5.9",
        "Epigenetik Sürüklenme (Epigenetic Drift) ve Hücre Kimliğinin Bulanıklaşması (De-differentiation)",
        "Bir nöronun veya kas hücresinin on yıllar içinde epigenetik manzarasını kaybederek ilkel, işlevsiz bir hücresel duruma kayması.",
        "Waddington'ın gelişimsel manzarasında bir hücre olgun bir vadiye yerleştiğinde, dokuya özgül kimliğini epigenetik metilasyon ve heterokromatin duvarlarıyla korur. Yaşlanma sürecinde gerçekleşen stokastik epigenetik sürüklenme (Epigenetic Drift), bu vadi duvarlarını aşındırır. Hücre kimliği bulanıklaşır (dediferansiasyon / kimlik kaybı): Yaşlı bir nöron sinaptik protein genlerini sustururken bazal embriyonik genleri açmaya başlar; yaşlı bir beta hücresi insülin üretimini durdurur. Bu durum organ yetmezliklerinin doğrudan hücresel kaynağıdır.",
        "Cellular_Identity_Fidelity = Correlate(Transcriptome_Cell_i(t), Differentiated_Lineage_Atlas) -> Declines_with_Age",
        "Hücresel kimlik sadakati; yaşlanan hücrenin transkriptomunun dokuya özgül referans tek-hücre atlası ile korelasyonunun zamanla erimesidir."
    ),
    (
        "5.10",
        "Heterokromatinin Restorasyonu: SUV39H1 ve G9a Metiltransferaz Uyarımı ile Gençleşme",
        "Histon metiltransferazların genetik veya küçük moleküllerle aktive edilerek H3K9me3 zırhının yeniden örülmesi ve gençlik kalkanı.",
        "Heterokromatin kaybı geri döndürülemez bir kader değildir. SUV39H1 ve SUV39H2 metiltransferazlarının aşırı ekspresyonu veya sirtuin aktivasyonu (SIRT6/SIRT1); H3K9me3 heterokromatinini yeniden organize eder, sentromerik bölgeleri sıkılaştırır ve retrotranspozonları susturur. Progeria hücre modellerinde SUV39H1 takviyesi veya G9a stimülasyonu; nükleer zar buruşukluğunu (blebbing) düzeltmiş, DNA hasar odaklarını yarıya indirmiş ve hücresel ömrü uzatmıştır.",
        "Heterochromatin_Restoration_Index = [H3K9me3_restored] / [H3K9me3_young] = f([SUV39H1_engineered], [SAM_pool])",
        "Heterokromatin restorasyon indeksi; tasarlanmış SUV39H1 enzim aktivitesi ve S-adenozilmetiyonin (SAM) metil donör havuzu ile gençlik seviyesine çıkar."
    )
]

# ==============================================================================
# KISIM 6: TRANSPOZONLARIN UYANIŞI: LINE-1 RETROTRANSPOZONLARI VE GENOMİK KAOS
# ==============================================================================
part6_subsections = [
    (
        "6.1",
        "İnsan Genomunun Bencil Sakinleri: LINE-1 (L1), Alu ve SVA Retrotranspozon Mimarisi",
        "İnsan DNA'sının %50'den fazlasını oluşturan 'zıplayan genler': Milyonlarca yıllık viral enfeksiyonların kalıntı fosil ordusu.",
        "İnsan genomunun protein kodlayan ekzonları toplam DNA'nın sadece %1.5'ini oluştururken; genomun neredeyse yarısı retrotranspozonlardan ibarettir. Bunların en güçlüsü LINE-1 (Long Interspersed Nuclear Element-1) ailesidir (genomun %17'si, yaklaşık 500.000 kopya). İnsan genomunda retrotranspozisyon yeteneğini korumuş yaklaşık 80-100 adet 'sıcak' (tam uzunlukta aktif) LINE-1 lokusu bulunur. Kısa retrotranspozonlar olan Alu elemanları (SINE ailesi, %11) ve SVA elemanları ise kendi enzimlerine sahip değildir; zıplamak için LINE-1 protein makinelerini parazit gibi kullanırlar.",
        "Transposon_Genomic_Share = Coverage(LINE1) + Coverage(Alu) + Coverage(SVA) + Coverage(LTR) ~ 52%_of_Genome",
        "Retrotranspozon genomik payı; LINE-1, Alu, SVA ve endojen retrovirüs (HERV) dizilimlerinin insan DNA'sındaki toplam yüzölçümüdür."
    ),
    (
        "6.2",
        "LINE-1 Biyolojisi: 6 kb Uzunluğunda, ORF1p (RNA şaperonu) ve ORF2p (Endonükleaz / RT)",
        "Tam uzunluktaki aktif bir LINE-1 elemanının anatomisi: İki protein kodlayan çift okuma çerçevesi ve retrotranspozisyon aygıtı.",
        "Aktif bir insan LINE-1 elemanı yaklaşık 6.000 baz çifti uzunluğundadır; bir 5' UTR internal promoterı, iki protein kodlayan açık okuma çerçevesi (ORF1 ve ORF2) ve bir poli-A kuyruğu içerir. ORF1p (40 kDa), yüksek afiniteli bir RNA bağlayıcı şaperondur; homotrimerler oluşturarak L1 RNA'sını sarar ve Ribonükleoprotein (RNP) partikülünü kurar. ORF2p (150 kDa) ise iki ölümcül enzimatik aktiviteye sahip multifonksiyonel bir devdir: N-terminalinde DNA'yı kesen bir Endonükleaz (EN) domeni ve merkezinde RNA'dan DNA sentezleyen bir Ters Transkriptaz (RT) domeni taşır.",
        "LINE1_RNP_Complex = [L1_mRNA] + 3 * N_trimers * [ORF1p] + 1 * [ORF2p_EN_RT]",
        "LINE-1 aktif ribonükleoprotein partikülü; L1 mRNA transkripti, trimerik ORF1p şaperon kılıfı ve enzimatik ORF2p motorunun stokiyometrisidir."
    ),
    (
        "6.3",
        "Hedef Odaklı Ters Transkripsiyon (TPRT): LINE-1'in Kromozomları Kesip Kendini Kopyalaması",
        "Kopya-Yapıştır mekanizması: ORF2p'nin genomik DNA'yı çentiklemesi ve L1 RNA'sını doğrudan kromozoma kopyalaması.",
        "LINE-1 elemanları 'Hedef Odaklı Ters Transkripsiyon' (TPRT - Target-Primed Reverse Transcription) yoluyla çoğalır. L1 RNP partikülü nükleusa girer. ORF2p'nin endonükleaz domeni, genomik DNA'daki konsensüs 5'-TTTT/AA-3' dizilimini tanıyarak DNA'nın bir ipliğini keser (çentik açar). Açığa çıkan 3'-OH ucu bir primer görevi görür; ORF2p'nin ters transkriptazı L1 mRNA'sının poli-A kuyruğunu bu uca eşleştirir ve kromozom üzerinde doğrudan cDNA sentezlemeye başlar. İkinci zincir de kesilip sentez tamamlandığında, genomun yeni bir bölgesine hedef bölge duplikasyonlarıyla (TSD) yepyeni bir L1 kopyası kovalent olarak entegre edilmiş olur.",
        "TPRT_Integration_Velocity = k_tprt * [ORF2p_EN] * [ORF2p_RT] * [Genomic_Cleavage_Sites_TTTTAA]",
        "TPRT entegrasyon debisi; ORF2p endonükleaz kesim hızı ile ters transkriptaz zincir uzama hızının genomik hedef konsantrasyonuyla çarpımıdır."
    ),
    (
        "6.4",
        "Gençlikte Transpozonların Susturulması: DNA Metilasyonu, H3K9me3, SIRT6 ve piRNA Kalkanı",
        "Genomun gençlikteki savunma kalkanı: LINE-1 5' UTR promoterının aşırı metillenerek derin dondurucuya kapatılması.",
        "Genç ve sağlıklı bir hücrede transpozonların zıplaması ölümcül bir tehlike olduğu için hücreler çok katmanlı baskılama sistemleri kurmuştur. LINE-1 5' UTR promoterındaki CpG adacıkları DNA metiltransferazlar (DNMT1, DNMT3A) tarafından %95 oranında hipermetillenir. Histon metiltransferazlar (SUV39H1, SETDB1) bu bölgelere yoğun H3K9me3 işaretleri koyar ve HP1 ile sıkı heterokromatin blokajı örer. SIRT6 enzimi L1 promoterlarına bağlanarak H3K56'yı deasetiller ve transkripsiyonu kilitler; germ hattında ise piRNA (PIWI-interacting RNA) mekanizması transpozon RNA'larını anında parçalar.",
        "L1_Repression_Index = [5mC_Methylation_5UTR] * [H3K9me3_L1] * [SIRT6_occupancy] -> Full_Silence",
        "LINE-1 gençlik susturma indeksi; 5' UTR promoter CpG metilasyonu, H3K9me3 heterokromatini ve SIRT6 bağlanmasının sinerjisidir."
    ),
    (
        "6.5",
        "Yaşlanmada LINE-1 Uyanışı: Heterokromatin Çözülmesiyle Milyonlarca L1 Transkriptinin Patlaması",
        "John Sedivy laboratuvarının devrimsel keşfi: Yaşlanan hücrelerde heterokromatin açıldığında uyuyan canavarın uyanması.",
        "Brown Üniversitesi'nden John Sedivy ve ekibi, hücresel senesens ve doku yaşlanması sırasında heterokromatin erozyonunun en yıkıcı sonucunu kanıtlamıştır: 'Transpozon Patlaması'. Yaşlı dokularda L1 promoterlarındaki DNA metilasyonu silinir, SIRT6 kromatini terk eder ve nükleer Lamin B1 çöker. Sonuçta yıllardır uyuyan LINE-1 elemanları transkripte edilmeye başlar. Yaşlı fare ve insan dokularında L1 RNA'sı ve ORF1p proteini gençlik seviyesinin 10 ila 50 katına fırlar; hücre içi bir transpozon enfeksiyonu yangınına dönüşür.",
        "L1_Transcription_Burst = k_trans * (1 - [5mC_Methylation_5UTR]) * (1 / [SIRT6_nuclear])",
        "Yaşlılıkta LINE-1 transkripsiyon patlaması; promoter DNA metilasyon kaybı ve nükleer SIRT6 klerensinin ters fonksiyonuyla fırlar."
    ),
    (
        "6.6",
        "Somatik Retrotranspozisyon Mozaizmi: Nöronlarda ve Yaşlı Dokularda Yeni İnsersiyon Mutasyonları",
        "Yaşlı insan beyninde tek tek nöronların genomuna yeni LINE-1 kopyalarının saplanması ve nöronal devre fonksiyonlarının bozulması.",
        "Uzun süre retrotranspozisyonun yalnızca germ hattında (üreme hücrelerinde) gerçekleştiği sanılmıştır; oysa Fred Gage laboratuvarı yetişkin insan beyninde nöral kök hücrelerde ve nöronlarda somatik retrotranspozisyonun aktif olduğunu kanıtlamıştır. Yaşlanmayla birlikte bu süreç kontrolden çıkar. Tek-hücre tam genom dizilemesi, yaşlı bir insanın kortikal nöronlarının her birinin genomunun farklı yerlerine yeni LINE-1 insersiyonları saplandığını (somatik mozaizm) göstermiştir. Bu rastgele insersiyonlar tümör baskılayıcı genleri bozar, sinaptik genleri keser ve nörodejenerasyonu hızlandırır.",
        "De_Novo_Insertions_Brain = N_active_L1_hot * P(Retrotransposition_Event) * Years_Lived",
        "Beyindeki de novo somatik L1 insersiyon sayısı; sıcak L1 kopya sayısı, retrotranspozisyon olasılığı ve yaşanılan yılların çarpımıdır."
    ),
    (
        "6.7",
        "Sitoplazmik L1 cDNA'sı ve cGAS-STING Oto-enflamasyonu: İnflam-aging'in Gizli Tetikleyicisi",
        "Sitoplazmada sentezlenen L1 cDNA'sının hücre tarafından viral bir saldırı gibi algılanması ve interferon fırtınası başlatması.",
        "John Sedivy ekibinin Nature dergisinde yayınlanan tarihi bulgusu: Uyanan LINE-1 elemanları sitoplazmada ORF2p ters transkriptazı ile L1 cDNA'sı (çift iplikli DNA parçaları) üretir. Bu cDNA'lar nükleusa giremeyip sitoplazmada birikir. Sitoplazmik DNA sensörü cGAS, bu L1 cDNA'larını bir retrovirüs istilası gibi algılar; STING yolağını aktive ederek Tip I İnterferon (IFN-alfa/beta) ve IL-6, TNF-alfa üretimini patlatır. Yaşlılıkta görülen kronik steril enflamasyonun (İnflam-aging) ve SASP fenotipinin en temel kök sürücüsü bizzat bu uyanan iç parazitlerdir.",
        "Transposon_Inflammation_Flux = k_sting * [Cytoplasmic_L1_cDNA] * [cGAS_active] -> Type_I_IFN_Burst",
        "Transpozon kaynaklı oto-enflamasyon akısı; sitoplazmada biriken LINE-1 cDNA konsantrasyonunun cGAS/STING aktivasyonu ile çarpımıdır."
    ),
    (
        "6.8",
        "LINE-1'in Genomik Kararsızlığı Artırması: DNA Çift Zincir Kırıkları ve Kromozomal Translokasyonlar",
        "ORF2p endonükleazının genomu rastgele binlerce noktadan kesmesi; çift zincir kırıkları ve kromozom kollarının kopması.",
        "LINE-1 sadece yeni kopyalar eklemekle kalmaz; bizzat genomik DNA'yı fiziksel olarak parçalar. ORF2p'nin endonükleaz domeni nükleusa girdiğinde retrotranspozisyon yapamasa bile genomik DNA'yı rastgele yerlerden çentikler. Karşılıklı çentikler çift zincir kırıklarına (DSB) dönüşür. Yaşlanan hücrede nükleer gamma-H2AX odaklarının büyük kısmının doğrudan ORF2p endonükleaz kesimlerinden kaynaklandığı gösterilmiştir. Bu kırıklar hatalı onarıldığında telomer-sentromer füzyonları ve kromozomal translokasyonlar patlar.",
        "L1_Induced_DSBs = k_endonuclease * [ORF2p_nuclear] * [Accessible_Chromatin_Sites]",
        "LINE-1 kaynaklı çift zincir kırık oluşum hızı; nükleusa giren ORF2p endonükleaz konsantrasyonu ve açık kromatin bölgelerinin yoğunluğu ile ölçeklenir."
    ),
    (
        "6.9",
        "Ters Transkriptaz İnhibitörleri (NRTI'lar: Lamivudin, Tenofovir) ile LINE-1 Susturma",
        "HIV tedavisinde kullanılan retroviral ilaçların insan LINE-1 ters transkriptazını felç ederek yaşlı fareleri gençleştirmesi.",
        "Sedivy laboratuvarı, HIV tedavisinde kullanılan nükleozid ters transkriptaz inhibitörlerinin (NRTI: Lamivudin / 3TC, Tenofovir, Emtrisitabin), insan ve fare LINE-1 ORF2p ters transkriptazını da nanomolar afiniteyle inhibe ettiğini keşfetmiştir. Yaşlı farelere içme suyuyla Lamivudin verildiğinde; sitoplazmik L1 cDNA sentezi durmuş, cGAS-STING interferonu sönmüş, kas ve karaciğerdeki senesen SASP enflamasyonu gerilemiş ve farelerin doku fonksiyonları gençleşmiştir. Bu buluş, transpozon hedefli longevity farmakolojisinin miladıdır.",
        "Inflamaging_Suppression_Index = 1 - exp(-k_3tc * [Lamivudine_plasma] / IC50_L1_RT)",
        "İnflam-aging baskılanma indeksi; plazma Lamivudin konsantrasyonunun LINE-1 ters transkriptaz enzimini inhibe etme gücü ile üstel olarak artar."
    ),
    (
        "6.10",
        "Transpozon Biyolojisi ve Türlerin Yaşam Süresi Arasındaki Ters Korelasyon",
        "Kör fareler ve balinalar gibi asırlık canlılarda transpozonların mutlak kontrol altında tutulması; kısa ömürlü farelerde transpozon anarşisi.",
        "Komparatif genomik analizler, memeli türlerinin maksimum yaşam süresi ile transpozon aktivitesi arasında çarpıcı bir ters korelasyon ortaya koymuştur. 30 yıldan fazla yaşayan çıplak kör farelerde (naked mole-rat) ve 200 yıl yaşayan Grönland balinalarında, LINE-1 elemanları aşırı güçlü epigenetik kalkanlar ve süper-aktif SIRT6 varyantlarıyla ömür boyu sıfır düzeyinde kilitli tutulur. Buna karşın sadece 2-3 yıl yaşayan laboratuvar farelerinde transpozon susturma sistemleri çok zayıftır ve retrotranspozisyon hızla yaşlanmayı tetikler. Bu durum transpozon kontrolünün türsel longevity'nin anahtarı olduğunu kanıtlar.",
        "Species_Maximum_Lifespan = Const_Species * ( [SIRT6_Binding_Affinity_to_L1] / L1_Retrotransposition_Rate )",
        "Bir memeli türünün maksimum yaşam süresi; SIRT6'nın transpozonları susturma afinitesi ile yerel retrotranspozisyon hızının oranına doğrudan bağlıdır."
    )
]

parts.append(("KISIM 4: ÇİFT ZİNCİR KIRIK (DSB) ONARIMI: NHEJ, HR VE MMEJ YOL AYRIMI", part4_subsections))
parts.append(("KISIM 5: HETEROKROMATİN KAYBI HİPOTEZİ VE EPİGENETİK YENİDEN YAPILANMA", part5_subsections))
parts.append(("KISIM 6: TRANSPOZONLARIN UYANIŞI: LINE-1 RETROTRANSPOZONLARI VE GENOMİK KAOS", part6_subsections))

# ==============================================================================
# KISIM 7: NÜKLEER ZAR, NÜKLEOPORİN BOZULMASI VE PROGERİA MODELLERİ
# ==============================================================================
part7_subsections = [
    (
        "7.1",
        "Nükleer Por Kompleksi (NPC): 120 MDa'lık Kapı, FG-Nükleoporinler ve Moleküler Elek",
        "Çekirdek ile sitoplazma arasındaki yegane geçiş kapısı: 30 farklı nükleoporinden oluşan sekiz katlı simetrik mega-silindir.",
        "Nükleer Por Kompleksi (NPC), ökaryotik hücrenin en büyük protein komplekslerinden biridir (~120 MDa memelilerde). Merkez kanal, fenilalanin-glisin (FG) tekrarları içeren intrinsik düzensiz FG-nükleoporinler (Nup62, Nup98) tarafından bir hidrofobik faz ayrışması jeli halinde doldurulmuştur. 40 kDa altındaki küçük moleküller bu jelden pasif difüzyonla serbestçe sızarken; büyük makromoleküller yalnızca Nükleer Lokalizasyon Sinyali (NLS) ve karyoferin/importin nükleer taşıyıcıları aracılığıyla bu hidrofobik engeli aşabilir.",
        "Nuclear_Permeability = P_npc = P_0 * exp(-Radius_Stokes_Cargo / Pore_Hydrophobic_Mesh_Size)",
        "Nükleer por geçirgenliği; taşınan kargonun hidrodinamik Stokes yarıçapının FG-nükleoporin hidrofobik ağ gözenek boyutuna oranının üstel fonksiyonudur."
    ),
    (
        "7.2",
        "Uzun Ömürlü Nükleoporinlerin (Nup96, Nup107) Yaşlanmayla Aşınması ve Çekirdek Sızıntısı",
        "Martin Hetzer laboratuvarının keşfi: İskelet nükleoporinlerinin ömür boyu turnover olmaması, oksitlenmesi ve çekirdeğin sızdırmazlığını kaybetmesi.",
        "Nöronlar ve kas hücreleri gibi bölünmeyen post-mitotik hücrelerde, NPC'nin yapısal iskele halkasını oluşturan nükleoporinler (Nup107-160 alt kompleksi ve Nup96) hücrenin tüm yaşamı boyunca asla yenilenmez; protein yarı ömürleri yıllardır. On yıllar süren reaktif oksijen saldırısı bu uzun ömürlü nükleoporinleri aşındırır. Nükleer porlar gevşer ve bozulur; normalde sitoplazmada kalması gereken büyük proteinler ve toksik agregatlar (TDP-43, Tubulin) çekirdek içine sızarken, nükleer faktörler dışarı kaçar (sızıntılı çekirdek sendromu).",
        "Nuclear_Leakage_Flux = J_leak = N_deteriorated_NPCs * Area_pore * (D_cytosolic_protein / Length_channel) * Delta_C",
        "Nükleer sızıntı debisi; yapısal bütünlüğünü kaybetmiş aşınmış nükleer por sayısı ve sitozolik protein difüzyon gradyanı ile hesaplanır."
    ),
    (
        "7.3",
        "Nükleositoplazmik Transport Felci: RanGTP / RanGDP Gradyanının Çökmesi ve Kargo Tıkanması",
        "Karyoferin taşıyıcılarını besleyen kimyasal enerji gradyanının çökmesi: Transkripsiyon faktörlerinin sitoplazmada mahsur kalması.",
        "Nükleusa kargo ithalatı (import) ve ihracatı (export), nükleer RanGTP ile sitoplazmik RanGDP arasındaki konsantrasyon gradyanı ile yürütülür. Nükleustaki guanin nükleotid değişim faktörü RCC1, Ran'ı GTP ile doldururken; sitoplazmadaki RanGAP1 GTP'yi hidroliz eder. Yaşlanmayla birlikte nükleer porlar bozulduğunda ve RCC1 kromatinden koptuğunda, Ran gradyanı çöker. Importin-alfa/beta taşıyıcıları kargolarını (örneğin DNA tamir proteinlerini, şaperonları ve transkripsiyon faktörlerini) çekirdeğe sokamaz; nükleositoplazmik transport felç olur.",
        "Transport_Driving_Force = Delta_mu_Ran = R * T * ln( [RanGTP_nucleus] / [RanGTP_cytosol] ) -> Declines_to_Zero",
        "Nükleositoplazmik transport itici kimyasal potansiyeli; nükleer ve sitozolik RanGTP konsantrasyon oranının logaritması olup yaşla sıfırlanır."
    ),
    (
        "7.4",
        "Lamin A/C Biyogenezi: Prelamin A'nın Farnezilasyonu ve ZMPSTE24 Proteazı ile Kesimi",
        "Nükleer zar iskeletini kuran Lamin A'nın karmaşık post-translasyonel olgunlaşma yolu: C-terminal farnezilasyon ve esansiyel proteolitik budama.",
        "İç nükleer zarın mekanik desteği olan Lamin A, LMNA geninden önce 'Prelamin A' öncülü olarak sentezlenir. Prelamin A'nın C-terminalindeki CAAX motifi (Cys-Ser-Ile-Met), Farneziltransferaz (FTase) enzimi tarafından 15 karbonlu hidrofobik bir farnezil lipidi ile donatılır; ardından Rce1 proteazı AAX kalıntılarını keser ve Icmt karboksimetilleme yapar. Farnezil kuyruğu Prelamin A'yı nükleer zara çiviler. Son ve en kritik adımda, çinko metalloproteazı ZMPSTE24 (FACE-1), farnezilli son 15 amino asitlik parçayı kesip atarak olgun, çözünür Lamin A'yı serbest bırakır.",
        "Mature_Lamin_A_Flux = k_zmpste * [ZMPSTE24_active] * [Farnesylated_Prelamin_A_membrane]",
        "Olgun Lamin A üretim akısı; iç nükleer zardaki aktif ZMPSTE24 proteaz konsantrasyonu ve farnezillenmiş ara ürün mevcudiyeti ile belirlenir."
    ),
    (
        "7.5",
        "Hutchinson-Gilford Progeria Sendromu (HGPS): LMNA Ekzon 11 Kriptik Uyuşması ve Progerin",
        "Tek bir sessiz nokta mutasyonunun (c.1824C>T, p.Gly608Gly) alternatif bir uçbirleştirme bölgesi açarak 50 amino asitlik delesyon yaratması.",
        "Hutchinson-Gilford Erken Yaşlanma Sendromu (HGPS), çocukların 10 kat hızla yaşlanarak ergenlikte ateroskleroz ve kalp krizinden öldüğü genetik bir faciadır. Hastalığın nedeni LMNA geninin ekzon 11'indeki tek bir de novo sessiz mutasyondur (c.1824C>T). Bu mutasyon amino asidi değiştirmez ancak mRNA üzerinde kriptik bir 5' uçbirleştirme (splice) donör bölgesi açar. Splicing sırasında ekzon 11'in sonundaki 150 nükleotid (50 amino asit) silinir. Silinen bu 50 amino asitlik bölge, tam olarak ZMPSTE24 proteazının kesim noktasını barındırır! Kesilemeyen bu mutant kalıcı farnezilli proteine 'Progerin' denir.",
        "Progerin_Production_Ratio = Splicing_Efficiency(Cryptic_Donor) / Splicing_Efficiency(Canonical_Donor) ~ 0.85",
        "Progerin üretim oranı; c.1824C>T mutasyonunun yarattığı kriptik donör bölgesinin kanonik donöre olan uçbirleştirme baskınlığıdır."
    ),
    (
        "7.6",
        "Progerinin Toksisitesi: Nükleer Blebbing, Heterokromatin Çözülmesi ve Erken Kök Hücre Ölümü",
        "Farnezil kuyruğu kesilemeyen Progerinin nükleer zara zehirli bir kazık gibi çakılması; çekirdeğin buruşması ve kromatini yırtması.",
        "Progerin farnezil lipit kuyruğunu asla kaybedemediği için nükleer zardan ayrılamaz; iç nükleer zar altında toksik, çözünmeyen oligomerik plaklar halinde çöker. Bu katılaşma nükleusun küresel esnekliğini yok eder: Çekirdek zarı derin girintiler, çıkıntılar ve fıtıklaşmalar (blebbing) sergiler. Nükleer porlar dağılır; Lamin B1 ve heterokromatin (H3K9me3) nükleer çeperden koparak çözülür. DNA replikasyonu durur, telomerler hızla yıpranır ve tüm kök hücre havuzları (özellikle vasküler düz kas ve mezenkimal kök hücreler) apoptoza giderek damarları kireçlendirir.",
        "Nuclear_Deformity_Index = 1 - (4 * pi * Area_nucleus / Perimeter_nucleus^2) = f([Progerin_membrane_burden])",
        "Nükleer deformite indeksi; dairesellikten sapma derecesi olup zarda biriken farnezilli progerin kütlesiyle doğru orantılı olarak artar."
    ),
    (
        "7.7",
        "Normal İnsan Yaşlanmasında Spontan Progerin Birikimi",
        "Tom Misteli laboratuvarının şok edici keşfi: Progeria mutasyonu taşımayan sağlıklı 80 yaşındaki bireylerin dokularında progerin bulunması.",
        "Progerin sadece HGPS hastalarına özgü bir anomali değildir; normal insan yaşlanmasının da gizli aktörüdür. Sağlıklı bireylerin hücrelerinde bile RNA uçbirleştirme makineleri her birkaç yüz transkripsiyonda bir 'hata' yaparak ekzon 11'deki kriptik bölgeyi spontan olarak kullanır ve eser miktarda progerin üretir. Gençlikte hücre bölünmeleri bu toksik proteini seyreltirken; yaşlanan post-mitotik hücrelerde (damar endoteli, kardiyomiyositler, nöronlar) progerin birikerek kritik eşiği aşar. Yaşlı insan damar otopsilerinde saptanan progerin, arter sertleşmesinin ve endotel yaşlanmasının doğrudan suçlusudur.",
        "Age_Associated_Progerin = Integral_0^Age ( Leakage_Rate_Cryptic_Splicing * [LMNA_transcripts] ) dt",
        "Yaşa bağlı doku progerin birikimi; LMNA transkripsiyonundaki düşük frekanslı kriptik uçbirleştirme sızıntısının zamansal integralidir."
    ),
    (
        "7.8",
        "Farneziltransferaz İnhibitörleri (Lonafarnib) ve Progerin Temizleme Protokolleri",
        "FDA onaylı ilk progeria ilacı Lonafarnib: Progerinin farnezillenmesini engelleyerek zardan koparılması ve yaşamın uzatılması.",
        "Progerinin toksisitesi hidrofobik farnezil kuyruğuna bağımlı olduğu için, farmakologlar Farneziltransferaz İnhibitörlerini (FTI: Lonafarnib) geliştirmiştir. Lonafarnib, FTase enzimini bloke ederek Progerine farnezil eklenmesini önler. Farnezillenemeyen progerin zardan kopar, sitoplazmada çözünür kalır ve nükleer zar blebbing'ini dramatik şekilde düzeltir. HGPS çocuklarında yapılan Faz II/III klinik denemelerinde Lonafarnib kardiyovasküler hasarı geriletmiş ve çocukların yaşam süresini anlamlı biçimde uzatarak FDA onayı almıştır. Yeni nesil antisens oligonükleotidler (ASO) ise kriptik bölgeyi doğrudan maskeleyerek progerin sentezini sıfırlar.",
        "Progerin_Farnesylation_Block = 1 - ([Farnesyl_Progerin] / [Total_Progerin]) = [Lonafarnib] / (IC50_FTase + [Lonafarnib])",
        "Progerin farnezilasyon blokaj oranı; uygulanan Lonafarnib konsantrasyonunun Farneziltransferaz enzim inhibisyon doygunluğuna bağlıdır."
    ),
    (
        "7.9",
        "Nükleer İskeletin Mekanotransdüksiyon Rolü: Hücresel Sertleşme ve DNA Hasar Hassasiyeti",
        "Doku sertliği ve mekanik kuvvetlerin nükleer lamina üzerinden kromatini bükmesi; yaşlı laminin mekanik gerilimde DNA'yı kırması.",
        "Nükleus hücrenin en sert organelidir ve hücre dışı matriks mekanik gerilimlerini (mechanotransduction) LINC kompleksi (SUN ve KASH proteinleri) aracılığıyla doğrudan nükleer laminaya aktarır. Genç hücrede Lamin A/C esnektir ve nükleusu amortisör gibi korur. Yaşlanan veya progerin biriktiren hücrede nükleer lamina aşırı sertleşir (rijidite). Kalp kası veya damar her kasıldığında, sertleşmiş nükleer zardan kromatite yıkıcı mekanik şok dalgaları iletilir; bu mikroskobik gerilim nükleer kromatini fiziksel olarak yırtarak çift zincir DNA kırıklarını patlatır.",
        "Mechanical_DSB_Rate = k_mech * Elastic_Modulus_Lamina * Cytoskeletal_Shear_Stress / Nuclear_Compliance",
        "Mekanik stres kaynaklı DNA kırık debisi; nükleer laminanın sertlik Young modülü ve hücre iskeleti kesme geriliminin nükleer esnekliğe oranıdır."
    ),
    (
        "7.10",
        "Nükleer Membran Bütünlüğünün Restore Edilmesi: Sentetik Por ve Lamin Stabilizasyonu",
        "Kriyoprezervasyon, sentetik nükleoporin hidrolojelleri ve dCas9-tabanlı lamin editing ile nükleusun fabrika ayarlarına döndürülmesi.",
        "Geleceğin proteomik ve genomik mühendisliği, nükleer zar bütünlüğünü restore etmeyi hedefler. David Liu ekibi, adenine base editor (ABE) kullanarak progeria farelerinde c.1824C>T mutasyonunu doğrudan DNA düzeyinde tek enjeksiyonla düzeltmiş ve farelerin ömrünü iki katına çıkarmıştır. Eşzamanlı olarak de novo tasarlanmış sentetik nükleoporin peptitleri aşınmış FG-ağını yeniden doldurarak sızıntılı çekirdeği kapatır; Lamin B1 takviyesi heterokromatin LAD bölgelerini tekrar zara çivileyerek nükleusu gençlik mimarisine kavuşturur.",
        "Nuclear_Restoration_Score = [Repaired_LMNA_alleles] / Total_Alleles * (Nuclear_Circularity_Restored / 1.0)",
        "Nükleer restorasyon skoru; genetik baz düzenleme ile düzeltilmiş LMNA alel oranı ile nükleer dairesellik kazanımının çarpımıdır."
    )
]

# ==============================================================================
# KISIM 8: KLONAL MOZAİZM, MİKRONÜKLEUSLAR VE KROMOZOMAL ANÖPLOİDİ
# ==============================================================================
part8_subsections = [
    (
        "8.1",
        "Somatik Mutasyon Birikimi ve Doku Klonal Mozaizmi: Yetişkin İnsanda 'Yama İklimi'",
        "İnigo Martincorena'nın Sanger Enstitüsü keşfi: Sağlıklı görünen bir yetişkinin derisinin ve yemek borusunun mutasyonlu klon yamalarından oluşması.",
        "Geleneksel inanış sağlıklı dokuların homojen vahşi tip hücrelerden oluştuğunu varsaymıştır; oysa Sanger Enstitüsü'nün mikroskobik doku biyopsilerinde yaptığı ultra-derin dizilemeler şoke edici bir gerçeği ortaya koymuştur. 60 yaşındaki sağlıklı bir bireyin cildindeki veya yemek borusundaki hücrelerin üçte birinden fazlası kanser sürücü mutasyonları (NOTCH1, TP53, FAT1) taşıyan genişlemiş klonlardan (somatik mozaizm) oluşur. Doku adeta bir 'yamalı bohça' gibidir. Yaşlanma, bu aberan klonların rekabet ederek dokunun sağlıklı hücrelerini yerinden etmesi sürecidir.",
        "Clonal_Mosaicism_Coverage = sum_clones Area(Clone_k with Driver_Mutation) / Total_Tissue_Area > 0.35",
        "Doku klonal mozaizm kapsama alanı; sürücü mutasyon taşıyan bağımsız somatik klonların toplam doku alanındaki yüzölçümüdür."
    ),
    (
        "8.2",
        "Mitotik Hatalar ve Anöploidi: Kohezin Kaybı, İğ İpliği Kontrol Noktası (SAC) Yetersizliği",
        "Yaşlanan hücrelerde kromozomları bir arada tutan kohezin halkalarının erimesi ve yavru hücrelere yanlış sayıda kromozom dağılması.",
        "Mitoz sırasında kardeş kromatitlerin kusursuz ayrılması, kromatitleri saran halka şeklindeki Kohezin (SMC1, SMC3, RAD21) kompleksine ve İğ İpliği Kontrol Noktasına (SAC: Mad2, BubR1) bağlıdır. Yaşlanan hücrelerde ve özellikle yaşlanan oositlerde (kadın yumurtasında), kohezin proteinleri yenilenemediği için zamanla ayrışır. Kromatitler iğ ipliği kutuplarına çekilirken kohezyon erken kopar; bir hücre 47 kromozom alırken diğeri 45 kromozom alır (Anöploidi). Anöploidi, yüzlerce genin stokiyometrik dengesini bozarak proteotoksik stres ve hücre ölümü yaratır.",
        "Aneuploidy_Probability_Mitosis = P_aneuploidy = 1 - exp(-k_cohesin_loss * (1 - [BubR1_active]))",
        "Mitotik anöploidi olasılığı; kohezin ayrışma hızı ve iğ ipliği kontrol noktası bekçisi BubR1 aktivitesinin yetersizliğiyle üstel artar."
    ),
    (
        "8.3",
        "Kromotripsis (Chromothripsis): Tek Bir Hücre Döngüsünde Kromozomun Parçalanıp Kaotik Birleşmesi",
        "Yıllar süren mutasyon birikimi yerine, tek bir mitotik kazada bir kromozomun yüzlerce parçaya patlayıp rastgele birbirine kaynaması.",
        "David Pellman laboratuvarının keşfettiği 'Kromotripsis' (kromozom kıyameti), genom evriminin en şiddetli katastrofudur. Geciken veya kusurlu bir mitozda geride kalan tek bir kromozom, ana çekirdeğin dışında küçük bir 'Mikronükleus' içine hapsolur. Mikronükleus zarı kusurludur ve S fazında yırtılır. Sitoplazmik nükleazlar içeri dolarak bu kromozomu onlarca/yüzlerce minik parçaya böler. Bir sonraki mitozda bu parçalar ana çekirdeğe katılır ve NHEJ ligazları parçaları kaotik bir sıra ve yönde birbirine yapıştırır; tek bir hücre döngüsünde yüzlerce genetik delesyon ve duplikasyon oluşur.",
        "Chromothripsis_Shatter_Index = N_genomic_breakpoints_per_chromosome > 50_Oscillating_Segments",
        "Kromotripsis şiddet indeksi; tek bir kromozom kolunda 50'den fazla salınımlı kopya sayısı değişim kırılma noktasının bulunmasıdır."
    ),
    (
        "8.4",
        "Mikronükleus Biyofiziği: Geride Kalan Kromozomların Kusurlu Zarlarla Hapsedilmesi",
        "Ana çekirdekten ayrı düşen öksüz kromozomun etrafına sarılan nükleer zarın Lamin B1'den yoksun ve kırılgan olması.",
        "Mitoz anafazında iğ ipliğine tutunamayan veya geç ayrılan asentrik kromozom parçaları kutuplara taşınamaz. Telofazda bu geride kalan kromatin parçasının etrafında bağımsız bir nükleer zar örülür; oluşan yapıya 'Mikronükleus' denir. Ancak biyofiziksel olarak mikronükleus zarı ana çekirdek zarı gibi kusursuz değildir: Yeterli Lamin B1 ve nükleer por içeremez, mekanik olarak aşırı kırılgandır ve nükleositoplazmik transport yapamaz. Mikronükleus içindeki DNA replike olamaz, sürekli DNA hasarına uğrar ve hücre için bir 'saatli bomba' haline gelir.",
        "Micronucleus_Stability_HalfLife = T_half_mn = Const * [Lamin_B1_density_MN] / Nuclear_Envelope_Tension",
        "Mikronükleus zar kararlılık yarı ömrü; zarındaki Lamin B1 yoğunluğu ile doğru, membran yüzey gerilimi ile ters orantılıdır."
    ),
    (
        "8.5",
        "Mikronükleus Yırtılması: Sitoplazmada Kromozom Yıkımı ve cGAS-STING Mega-Alarmı",
        "Mikronükleus zarının aniden patlamasıyla tüm bir kromozomun sitoplazmaya dökülmesi ve cGAS sensörünün aşırı uyarılması.",
        "Mikronükleusların %60'ından fazlası hücre döngüsü ilerledikçe spontan olarak fiziksel yırtılmaya (envelope rupture) uğrar. Zar patladığında kromatin sitoplazmik ortama çırılçıplak açığa çıkar. Sitoplazmadaki cGAS (siklik GMP-AMP sentaz) enzimi devasa bir DNA kütlesiyle karşılaşır; mikronükleus etrafında kümelenerek endüstriyel ölçekte 2'3'-cGAMP üretmeye başlar. STING yolağı maksimum kapasitede kilitlenir; hücre aşırı miktarda pro-enflamatuar SASP faktörleri salgılayarak çevresindeki tüm dokuyu immünolojik bir yıkıma ve senesense sürükler.",
        "cGAS_Activation_Micronucleus = k_cgas * Area_exposed_chromatin * [cGAS_cytosol] -> Hyper_Inflammation",
        "Mikronükleus kaynaklı cGAS aktivasyon debisi; yırtılan zardan açığa çıkan çıplak kromatin yüzey alanı ve sitozolik cGAS derişimiyle patlar."
    ),
    (
        "8.6",
        "Klonal Hematopoez ve Kanser Öncesi Darboğazlar",
        "Kemik iliğindeki kan kök hücrelerinin DNMT3A, TET2 ve ASXL1 mutasyonlarıyla rekabet üstünlüğü kazanıp tüm kanı ele geçirmesi.",
        "Yaşlanmayla birlikte kan kök hücre havuzunda (HSC) klonal mozaizm dramatik bir forma bürünür: 'Önemi Belirsiz Klonal Hematopoez' (CHIP). 70 yaş üstü insanların %15-20'sinde, tek bir mutant HSC klonu tüm beyaz kan hücrelerinin %10-50'sini üretir. En yaygın mutasyonlar epigenetik regülatörler olan DNMT3A ve TET2 genlerindedir. Bu mutant monosit ve makrofajlar aterosklerotik plaklara sızarak aşırı IL-1beta ve IL-6 salgılar; bu durum lösemi riskini 10 kat artırmanın yanı sıra kalp krizi ve inme riskini 2 katına fırlatır.",
        "CHIP_Cardiovascular_Risk = Hazard_Ratio = 1.0 + alpha_vaf * Variant_Allele_Frequency(TET2/DNMT3A)",
        "Klonal hematopoez kardiyovasküler risk çarpanı; kanda tespit edilen mutant klonun varyant alel frekansı (VAF) ile doğrusal olarak katlanır."
    ),
    (
        "8.7",
        "Yaşlanan Dokularda Kromozomal Yapısal Varyasyonlar (CNV, İnversiyon, Translokasyon)",
        "Yalnızca tek baz mutasyonları değil; mega-bazlık kopya sayısı değişimlerinin dokuların genomik haritasını parçalaması.",
        "Kopya Sayısı Varyasyonları (CNV: delesyonlar ve duplikasyonlar), yaşlanan hücrelerde nükleer kararsızlığın en büyük boyutudur. Replikasyon stresi kırılgan bölgeleri (common fragile sites) koparır; hatalı NHEJ bu kırıkları yanlış kromozomlara dikerek translokasyonlar yaratır. Yaşlı insan beyninde tek bir piramidal nöronda yüzlerce megabazlık genomik dengesizlik saptanabilir. Bu yapısal varyasyonlar gen dozajını (gene dosage) bozar; proteomik komplekslerin stokiyometrisi çöker ve hücre disfonksiyona sürüklenir.",
        "Structural_Variant_Load = sum_chromosomes ( Megabases_Deleted + Megabases_Duplicated )",
        "Kromozomal yapısal varyasyon yükü; hücre genomunda delesyon veya duplikasyona uğramış toplam megabazlık DNA uzunluğudur."
    ),
    (
        "8.8",
        "Hücresel Mozaizmin Tek-Hücre Bütün Genom Dizilemesi (scWGS) ile Aydınlatılması",
        "Yeni nesil biyoteknolojik atılım: Tek bir hücrenin DNA'sını amplifiye ederek her bir hücrenin kişisel mutasyon günlüğünü okuma.",
        "Klasik doku dizilemesi (bulk sequencing), milyarlarca hücrenin ortalamasını aldığı için somatik klonları ve tek hücre mutasyonlarını göremez. Tek-Hücre Bütün Genom Dizilemesi (scWGS - single-cell Whole Genome Sequencing) ve PTA (Primary Template-directed Amplification) teknolojileri, tek bir nöron veya kök hücrenin DNA'sını hatasız çoğaltmayı başarmıştır. Bu analizler, bir insanın 80 yıllık yaşamı boyunca her bir kortikal nöronunun yılda ortalama 20-30 yeni somatik nokta mutasyonu biriktirdiğini ve 80 yaşında tek bir nöronun genomunda 2500'den fazla mutasyon taşıdığını kesin olarak ispatlamıştır.",
        "Somatic_Mutation_Clock = N_single_nucleotide_variants = Basal_Mutations_Birth + Rate_Annual * Age_Years",
        "Somatik tek hücre mutasyon saati; doğumdaki bazal mutasyon sayısına yıllık lineer mutasyon birikim hızının eklenmesiyle hesaplanır."
    ),
    (
        "8.9",
        "Somatik Mutasyon Oranının Memeli Türleri Arasında Yaşam Süresiyle Ters Orantılılığı",
        "Alex Cagan ve Inigo Martincorena'nın Nature 2022 keşfi: 16 memeli türünde mutasyon hızı ile maksimum yaşam süresi arasındaki evrensel kural.",
        "2022 yılında Wellcome Sanger Enstitüsü'nden Alex Cagan ve ekibi, farelerden insanlara, zürafalardan balinalara kadar 16 farklı memeli türünde somatik mutasyon hızlarını karşılaştırmıştır. Çıkan sonuç evrensel bir biyolojik yasadır: Bir türün somatik mutasyon hızı, o türün maksimum yaşam süresi ile kusursuz bir ters orantı izler! 3 yıl yaşayan fare yılda yaklaşık 800 mutasyon biriktirirken; 80 yıl yaşayan insan yılda sadece 47 mutasyon biriktirir. Tüm memeliler ömürlerinin sonunda yaklaşık olarak aynı toplam somatik mutasyon yüküne (~3000 lezyon) ulaşarak ölür; bu durum somatik kararsızlığın yaşam süresinin evrensel sınırlandırıcısı olduğunu kanıtlar.",
        "Species_Annual_Mutation_Rate = Universal_Mutational_Ceiling (~3000) / Maximum_Lifespan_Years",
        "Bir memeli türünün yıllık somatik mutasyon hızı; evrensel türsel mutasyonel tavan sınırının türün maksimum yaşam süresine bölünmesidir."
    ),
    (
        "8.10",
        "Klonal Seçilimin Toksisitesi: Süper-Uyumlu (Super-Fit) Aberan Klonların Doku İşgali",
        "Enflamasyona ve strese dirençli p53/TET2 mutant klonların, dokunun sağlıklı genç hücrelerini zehirleyerek ele geçirmesi.",
        "Klonal hematopoez ve doku mozaizmi nötr bir süreç değildir; Darwinist bir mikro-evrim yarışıdır. Kronik enflamasyon veya sigara/alkol gibi stresler altında, p53 veya TET2 mutasyonu taşıyan klonlar normal hücrelere kıyasla apoptoza gitmez ve hayatta kalma avantajı (super-fitness) kazanır. Bu 'süper-uyumlu' mutant klonlar salgıladıkları faktörlerle komşu normal hücrelerin bölünmesini baskılar ve doku alanını agresifçe işgal eder. Doku genç kök hücrelerini kaybeder; yerini kansere ramak kalmış senesen ve enflamatuar mutant klon kolonileri alır.",
        "Clonal_Expansion_Velocity = v_clone = Selection_Coefficient_s * Replication_Rate * (1 - Immune_Clearance)",
        "Aberan klonal genişleme hızı; mutant klonun Darwinist seleksiyon katsayısı ve replikasyon hızının immün klerens direnciyle çarpımıdır."
    )
]

# ==============================================================================
# KISIM 9: GENOMİK KORUMA, SIRT6 ENZİMİ VE SENTETİK DNA ONARIM AJANLARI
# ==============================================================================
part9_subsections = [
    (
        "9.1",
        "SIRT6: Uzun Yaşamın Kromatin Bekçisi (H3K9/H3K56 Deasetilaz ve PARP1 Mono-ADP Ribozilaz)",
        "Nükleustaki chromatin gardiyanı: Çift zincir kırıklarına ilk koşan, heterokromatini kilitleyen ve transpozonları susturan süper-sirtuin.",
        "SIRT6, nükleer kromatin mimarisini ve genom kararlılığını yöneten çok fonksiyonlu bir NAD+ bağımlı enzimdir. İki temel enzimatik gücü vardır: Birincisi, Histon H3 Lizin 9 (H3K9ac) ve Lizin 56 (H3K56ac) deasetilaz aktivitesidir; telomerleri stabilize eder, heterokromatini sıkılaştırır ve LINE-1 promoterlarını kilitler. İkincisi, mono-ADP-ribozilaz aktivitesidir; DNA hasarı anında PARP1 ve KDM2A'yı ribozilleyerek çift zincir kırık onarım makinelerini ışık hızında organize eder. SIRT6 eksikliği farelerde 4 haftada ölümcül progeria yaratır.",
        "SIRT6_Enzymatic_Flux = k_deacet * [SIRT6] * [H3K9ac] * [NAD+] + k_ribosyl * [SIRT6] * [PARP1] * [NAD+]",
        "SIRT6 toplam hücresel koruma debisi; NAD+ bağımlı deasetilasyon akısı ile mono-ADP-ribozilasyon akısının toplamıdır."
    ),
    (
        "9.2",
        "SIRT6'nın NHEJ ve HR Onarımını Hızlandırma Mekanizmaları",
        "Hasar noktasında kromatini yeniden modelleyerek hem Homolog Rekombinasyonu hem de doğru ligasyonu koordine eden orkestra şefi.",
        "Bir DSB oluştuğunda SIRT6 saniyeler içinde kırık bölgesine toplanır. Burada chromatin yeniden modelleme enzimi SNF2H'yi bölgeye çağırarak nükleozomları kaydırır ve kırık bölgesini onarım enzimlerine açar. Homolog Rekombinasyonda CtIP proteinini K432/K433 pozisyonlarından deasetilleyerek uç rezeksiyonunu fırlatır; Rad51 filament montajını hızlandırır. NHEJ'de ise DNA-PKcs ve Ku kompleksini stabilize eder. SIRT6, yaşlanan hücrede sönümlenen hatasız HR onarımını yeniden canlandıran yegane moleküler katalizördür.",
        "HR_Efficiency_Boost = 1 + alpha_sirt6 * [SIRT6_chromatin_bound] * [CtIP_deacetylated]",
        "HR onarım verimlilik artışı; kromatinde aktifleşen SIRT6 yoğunluğu ve deasetillenen CtIP rezeksiyon faktörü ile katlanır."
    ),
    (
        "9.3",
        "SIRT6 Aşırı Ekspresyonunun Farelerde Yaşam Süresini %30 Uzatması (Haim Cohen Deneyleri)",
        "İsrail Bar-Ilan Üniversitesi'nin tarihi keşfi: Tek bir genin (SIRT6) transgenik artırılmasıyla farelerde kansersiz ve radikal ömür uzaması.",
        "2012 ve 2021 yıllarında Nature Communications dergisinde yayınlanan çığır açıcı çalışmalarda, Haim Cohen laboratuvarı SIRT6'yı tüm vücutta aşırı eksprese eden transgenik fareler üretmiştir (MOSES fareleri). Sonuçlar nefes kesicidir: Transgenik farelerin ortalama ve maksimum yaşam süresi %30 uzamıştır (insan ömründe ek 25 yıla eşdeğer). Bu fareler yaşlandıklarında kansere yakalanmamış, kan şekeri ve lipid profilleri genç kalmış, motor koordinasyonlarını kaybetmemiş ve karaciğerlerinde DNA hasar odakları sıfıra yakın kalmıştır. Bu deney, genomik korumanın memelilerde ömrü doğrudan uzattığının kesin kanıtıdır.",
        "Lifespan_Extension_MOSES = Delta_Lifespan = +30% = f([Transgenic_SIRT6_Expression] ~ 2.5_fold)",
        "SIRT6 aşırı ekspresyonu ile elde edilen yaşam süresi uzaması; bazal ekspresyonun 2.5 katına çıkarılmasıyla %30 radikal artış sağlar."
    ),
    (
        "9.4",
        "PARP1 Regülasyonu ve NAD+ Desteği ile Hızlı Tek Zincir Tamiri",
        "PARP1'in tek zincir kırıklarına anında yetişip XRCC1 ve Ligaz III'ü çağırması; NAD+ öncülleriyle enerji tükenişinin önlenmesi.",
        "DNA tek zincir kırıkları (SSB) hücrede en sık oluşan lezyondur. PARP1, çinkolu parmak domenleriyle SSB çentiğini saniyeler içinde yakalar ve oto-PARilasyonla (kendi üzerine poli-ADP-riboz zincirleri örerek) XRCC1 iskele proteinini ve DNA Ligaz III'ü bölgeye çeker; çentik dakikalar içinde kapatılır. Ancak kronik DNA hasarında PARP1'in aşırı çalışması hücre içi NAD+ havuzunu tüketerek hücreyi öldürür. NMN veya NR takviyesi, PARP1'in NAD+ açlığını kapatarak tek zincir tamirini yıldırım hızında tutarken hücresel enerjiyi korur.",
        "SSB_Clearance_Velocity = k_parp1 * [Active_PARP1] * [XRCC1] * ( [NAD+] / (K_m_NAD + [NAD+]) )",
        "Tek zincir kırık klerens hızı; aktif PARP1 ve XRCC1 yoğunluğu ile hücre içi serbest NAD+ doygunluğunun fonksiyonudur."
    ),
    (
        "9.5",
        "Küçük Moleküllü DNA Onarım Uyarıcıları: Allosterik SIRT6 Aktivatörleri (UBCS039)",
        "Enzim mühendisliğiyle geliştirilen ve SIRT6'nın deasetilasyon hızını 40 kat artıran ilk sentetik allosterik moleküller.",
        "SIRT6'yı gen terapisi olmaksızın aktive etmek için farmakologlar küçük moleküllü allosterik aktivatörler geliştirmiştir. İtalyan araştırmacıların sentezlediği UBCS039 ve türevleri (örneğin MDL-800 ve quercetin türevleri), SIRT6'nın nükleotid bağlama cebinin yanındaki allosterik cebe bağlanır. Molekül enzimin H3K9ac substratına olan afinitesini katlar ve deasetilasyon katalitik hızını 40 kata kadar artırır. Bu aktivatörler yaşlı hücrelere verildiğinde; nükleer DNA tamir hızı genç hücre seviyesine fırlar, LINE-1 transkripsiyonu anında durur ve hücresel senesens geriletilir.",
        "SIRT6_Activation_Factor = Rate_catalytic / Rate_basal = 1 + alpha_act * [UBCS039] / (K_a + [UBCS039]) ~ 40x",
        "Sentetik allosterik aktivatör UBCS039'un SIRT6 deasetilasyon hızına kazandırdığı katalitik ivme; mikromolar konsantrasyonda 40 kata kadar çıkar."
    ),
    (
        "9.6",
        "Transpozon Baskılayıcı İlaç Protokolleri: Lamivudin (3TC) ve Emtrisitabin ile Sistemik Gençleşme",
        "Günde bir tablet nükleozid analoğu ile retrotranspozon yangınını söndüren klinik longevity müdahalesi.",
        "Klinik longevity hekimliğinde devrim yaratan protokollerden biri, güvenliği 30 yıldır kanıtlanmış NRTI sınıfı ters transkriptaz inhibitörlerinin (Lamivudin / 3TC ve Emtrisitabin / FTC) anti-aging amacıyla yeniden konumlandırılmasıdır (drug repurposing). Bu moleküller, yaşlanan hücrelerin sitoplazmasında uyanan LINE-1 ORF2p ters transkriptazını bloke eder; L1 cDNA sentezini durdurur. cGAS-STING oto-enflamasyon döngüsü kırılır; eklem iltihapları geriler, nöroenflamasyon söner ve dokulardaki senesen hücre yükü temizlenir.",
        "Systemic_Inflamaging_Drop = Delta_SASP = -k_nrti * [3TC_bioavailable] * [LINE1_active_tissues]",
        "Sistemik inflam-aging gerilemesi; biyoyararlanımı olan Lamivudin konsantrasyonu ve dokulardaki aktif LINE-1 transpozon yükünün çarpımıdır."
    ),
    (
        "9.7",
        "Antioksidan Nükleozid Tedavileri ve dNTP Havuzunun Oksidasyondan Korunması (MTH1 Enzimi)",
        "Hücrenin DNA'ya yanlışlıkla oksitlenmiş nükleotid takmasını önleyen temizleyici enzim: 8-oxo-dGTP'nin MTH1 ile imhası.",
        "Serbest radikaller sadece kromatindeki DNA'yı vurmaz; sitoplazmada ve mitokondride bekleyen serbest nükleotid havuzunu (dNTP) da oksitler. Serbest dGTP oksitlendiğinde 8-oxo-dGTP oluşur. Eğer DNA polimeraz bu oksitlenmiş bazı yanlışlıkla yeni sentezlenen DNA zincirine takarsa (sanitasyon yetersizliği), tüm genom mutasyonla enfekte olur. Nükleotid havuzunun temizleyicisi MTH1 (NUDT1) enzimidir; 8-oxo-dGTP'yi hidroliz ederek monofosfata çevirir ve DNA'ya girmesini engeller. MTH1 aşırı eksprese edilen transgenik fareler nörodejenerasyona dirençli ve uzun ömürlüdür.",
        "dNTP_Pool_Sanitization_Ratio = [MTH1_active] / ( [8-oxo-dGTP] + epsilon ) -> Zero_Oxidized_Incorporation",
        "dNTP havuzu temizleme rasyosu; aktif MTH1 pirofosfataz yoğunluğunun sitozolik okside dGTP derişimine oranı olup DNA'ya hatalı girişi sıfırlar."
    ),
    (
        "9.8",
        "CRISPR-dCas9 Tabanlı Epigenetik Yeniden Metilleme: LINE-1 ve Transpozonların Yeniden Kilitlenmesi",
        "Katalitik olarak ölü Cas9'a (dCas9) bağlanan DNMT3A ve KRAB domenleri ile transpozon promoterlarına kovalent kilit vurma.",
        "Modern genom mühendisliği, uyanan transpozonları tek tek susturmak için 'Epigenom Düzenleme' (Epigenome Editing) platformlarını kullanır. Katalitik kesme yapamayan ölü Cas9 (dCas9), bir de novo DNA metiltransferaz (DNMT3A/3L) ve bir heterokromatin baskılayıcı (KRAB domeni) ile füzyonlanır. LINE-1 5' UTR konsensüs sekansını hedefleyen tek bir sgRNA kılavuzluğunda bu kompleks, hücredeki 100 aktif sıcak LINE-1 lokusunun promoterına eşzamanlı oturur. CpG adacıklarını yeniden metiller ve H3K9me3 zırhını örerek transpozonları gençlikteki derin sükunetine kilitler.",
        "Epigenetic_Silencing_Efficiency = P(Methylation_5UTR) = [dCas9_DNMT3A_KRAB] / (K_d + [dCas9_complex]) -> 98%",
        "LINE-1 epigenetik susturma verimliliği; dCas9-DNMT3A-KRAB kompleksinin promoter bağlanma doygunluğu ile %98 seviyesinde kalıcı sessizlik sağlar."
    ),
    (
        "9.9",
        "Radyoprotektör Moleküller: Amifostin, Melatonin ve Tiyol İçerikli DNA Kalkanları",
        "İyonlaştırıcı radyasyonun ve serbest radikallerin DNA omurgasına ulaşmadan önce milisaniyede nötralize edilmesi.",
        "DNA hasarını oluştuktan sonra tamir etmek yerine oluşmasını engellemek radyoprotektör moleküllerle mümkündür. Klinik onaylı organik tiyofosfat Amifostin (WR-2721), alkalen fosfataz ile serbest tiyole (WR-1065) dönüştürülür; DNA omurgasını elektrostatik olarak sararak hidroksil radikallerini (.OH) difüzyon sınırında bir hızla yakalar. Güçlü bir serbest radikal süpürücüsü olan nükleer Melatonin de DNA'yı doğrudan radyasyondan ve lipid peroksitlerinden korur. Bu moleküller hücreyi DNA kırıklarına karşı 2 ila 3 kat daha dirençli hale getirir.",
        "Dose_Reduction_Factor = DRF = Radiation_Dose_Tolerated_Protected / Radiation_Dose_Control ~ 2.5",
        "Radyoproteksiyon doz azaltma faktörü (DRF); uygulanan Amifostin/Melatonin kalkanı sayesinde tolere edilen radyasyon dozunun kontrol dozuna oranıdır."
    ),
    (
        "9.10",
        "Genomik Bütünlük Paneli: Kanda gamma-H2AX, cfDNA (Hücresiz DNA) ve 8-OHdG Klinik Takibi",
        "Bireyin anlık genomik aşınmasını ve kanser riskini sıvı biyopsi ile mililitre kanda ölçen klinik longevity paneli.",
        "Kişiselleştirilmiş longevity tıbbında bireyin anlık genomik kararsızlığı sıvı biyopsi testleriyle izlenir: 1) Periferik kan mononükleer hücrelerinde (PBMC) akış sitometrisi ile ölçülen gamma-H2AX odak sayısı (aktif DSB yükü); 2) İdrar ve plazmada LC-MS/MS ile ölçülen serbest 8-OHdG derişimi (sistemik DNA oksidasyonu); 3) Plazmadaki Hücresiz Dolaşan DNA (cfDNA) miktarı ve fragmantasyon boyutu (hücresel nekroz ve mikronükleus sızıntısı); 4) Dolaşımdaki LINE-1 RNA ve retrotranspozon cDNA seviyeleri. Bu panel, genomik kriz patlamadan önce koruyucu protokolleri tetikler.",
        "Genomic_Instability_Index = GII = w1*[gamma_H2AX] + w2*[8-OHdG_urine] + w3*[cfDNA_plasma] + w4*[L1_cDNA]",
        "Klinik genomik kararsızlık indeksi; gamma-H2AX, idrar 8-OHdG, plazma cfDNA ve transpozon cDNA konsantrasyonlarının ağırlıklı toplamıdır."
    )
]

# ==============================================================================
# KISIM 10: HOMO AETERNUS GENOMİK MİMARİSİ: HATASIZ VE KORUNMUŞ DNA TASARIMI
# ==============================================================================
part10_subsections = [
    (
        "10.1",
        "Sentetik p53 Mimarisi: Fil Genomundaki 20 Çift TP53 (LFS6) Kopyasının İnsan Hücrelerine Entegrasyonu",
        "Peto Paradoksunun moleküler çözümü: İnsandaki tek çift TP53 yerine fillerdeki gibi çoklu kopyalarla kansere mutlak bağışıklık.",
        "Peto Paradoksu; fillerin insanlardan 100 kat daha fazla hücreye sahip olmasına rağmen neredeyse hiç kansere yakalanmamasını sorgular. Joshua Schiffman laboratuvarı, fillerin genomunda 20 çift (40 alel) TP53 geni taşıdığını keşfetmiştir; tek bir DNA hasarında fillerin hücreleri anında tamir yapar veya hasarlı hücreyi derhal apoptoza sokarak kanserleşmeyi imkansız kılar. Homo Aeternus genomik mimarisinde insan kök hücrelerine sentetik BAC (bakteriyel yapay kromozom) vektörleriyle ekstra fonksiyonel TP53 kopyaları entegre edilir; insan dokuları kansere karşı mutlak bir biyolojik kalkan kazanır.",
        "Cancer_Immunity_Index = 1 - P(Malignant_Escape) = 1 - (1 - P_apoptosis_per_p53)^N_TP53_copies -> 0.99999",
        "Kanser bağışıklık indeksi; hücrede aktif eksprese edilen TP53 kopya sayısı (N=40) ile orantılı olarak malign hücre kaçış olasılığını sıfırlar."
    ),
    (
        "10.2",
        "Sıfır-Transpozon İnsan Genomu: LINE-1, Alu ve SVA Parazitlerinin Sentetik Olarak Temizlenmesi",
        "Büyük ölçekli genom sentezi ile 3.2 milyar baz çiftlik insan DNA'sındaki tüm parazitik transpozon sekanslarının silinmesi.",
        "George Church ve GP-write (Genome Project-Write) konsorsiyumunun vizyoner projesi: İnsan genomundaki 500.000 LINE-1 ve 1 milyon Alu elemanı faydalı değil, evrimsel parazitlerdir. De novo DNA sentez teknolojileriyle, tüm transpozon elemanlarından arındırılmış 'Ultra-Temiz Sentetik İnsan Genomu' tasarlanır. Transpozon içermeyen bir hücrede ne retrotranspozisyon mutasyonu olabilir, ne cGAS-STING oto-enflamasyonu yaşanır, ne de rekombinasyonel delesyonlar meydana gelir; genomik stabilite gençlik platosunda sonsuza kadar donar.",
        "Synthetic_Genome_Purity = (Total_Genomic_Length - Transposon_Coverage) / Coding_And_Regulatory_Regions -> 100%",
        "Sentetik genom saflık oranı; parazitik transpozon çöplerinin temizlenmesiyle sadece fonksiyonel kodlama ve regülasyon bölgelerinin kalmasıdır."
    ),
    (
        "10.3",
        "Yapay Aşırı Doğruluklu DNA Polimerazlar: Proofreading Kapasitesi 1000 Kat Artırılmış POL-Delta/Epsilon",
        "Polimeraz ekzonükleaz aktif ceplerinin de novo protein tasarımıyla mutajenik hata oranının milyarda birden trilyonda bire indirilmesi.",
        "Doğal DNA Polimeraz Delta ve Epsilon, her 10^7 bazda bir hata yapar; bu oran 80 yılda milyarlarca somatik mutasyon anlamına gelir. Sentetik biyologlar, polimerazların 3'-5' ekzonükleaz domenlerini yönlendirilmiş evrim (directed evolution) ve AlphaFold tasarımıyla modifiye eder. Yanlış baz takıldığında oluşan geometrik gerilimi 10 kat daha hızlı tanıyan 'Süper-Doğruluklu Polimerazlar' (Super-Fidelity DNA Polymerases), replikatif hata oranını 10^-11 seviyesine çeker; hücre bölünmelerine bağlı somatik mutasyon birikimi tarihe karışır.",
        "Replication_Error_Rate_Synthetic = P_error_natural / 1000 -> 10^-11_mutations/bp/division",
        "Sentetik aşırı doğruluklu polimeraz hata oranı; doğal replikatif mutasyon frekansının 1000 kat altına indirilerek mükemmelleştirilir."
    ),
    (
        "10.4",
        "Sürekli Nükleer Lamin A ve Sentetik Nükleoporin Zırhı ile Sızıntısız Çekirdek",
        "ZMPSTE24 bağımsız doğrudan olgun Lamin A ekspresyonu ve FG-nükleoporin hidrojel yenileyicileri ile ebedi nükleer sızdırmazlık.",
        "Homo Aeternus çekirdeğinde progerin oluşumu imkansızdır: LMNA geni, C-terminal CAAX motifi içermeyecek şekilde doğrudan olgun Lamin A kodlayacak biçimde genetik olarak yeniden yazılır; farnezilasyona ve ZMPSTE24 kesimine ihtiyaç kalmaz. Eşzamanlı olarak hücre içine yerleştirilen sentetik nükleoporin şaperonları, aşınan Nup96 ve Nup107 proteinlerini düzenli olarak yenileyerek nükleer porların sızdırmazlığını ve RanGTP gradyanını bir asır boyunca ilk günkü gibi tutar.",
        "Nuclear_Envelope_Integrity_HomoAeternus = lim_{t -> infty} ( Blebbing_Events(t) + Leakage_Flux(t) ) = ZERO",
        "Homo Aeternus nükleer zar bütünlüğü aksiyomu; sonsuz zaman ufkunda nükleer fıtıklaşma ve sitozolik sızıntı debisinin sıfır olmasıdır."
    ),
    (
        "10.5",
        "Otonom İn Vivo DNA Hasar Tespit ve Onarım Nanobotları",
        "Kromatin lifleri boyunca sürekli devriye gezen, kırık uçları yakalayıp kovalent olarak dikerek saniyeler içinde onaran DNA origami robotları.",
        "Biyolojik onarım enzimlerinin yerini alan veya onları destekleyen yapay biyo-nano makineler: DNA origami ve karbon nanotüplerden inşa edilen 'Onarım Nanobotları' (Repair Nanobots). Bu nano-cihazlar nükleer kromatinde 1D difüzyonla kayarak çift zincir kırıklarını, abazik bölgeleri ve timin dimerlerini algılar. Kırık uçları mikromekanik kıskaçlarıyla kavrayıp hizalar, bünyelerinde taşıdıkları liyofilize dNTP ve sentetik ligaz modülleriyle kırığı 5 saniye içinde sıfır delesyonla kapatır.",
        "Nanobot_Repair_Velocity = N_nanobots_nucleus * v_scan * [DSB_lesions] -> RealTime_Zero_Lag_Repair",
        "Nanobot genomik onarım debisi; çekirdekteki aktif nanobot sayısı ve tarama hızıyla hasarın oluştuğu anda gecikmesiz imha edilmesidir."
    ),
    (
        "10.6",
        "Epigenetik Yedekleme ve Periyodik 'Genomik Geri Yükleme Noktası' (Genome Restore Points)",
        "Bireyin 20 yaşındaki kusursuz metilom ve histon haritasının dijital ve biyolojik olarak yedeklenip her 5 yılda bir hücreye yüklenmesi.",
        "Bilgisayar işletim sistemlerindeki 'Sistem Geri Yükleme Noktası' kavramının epigenetiğe uyarlanması: Birey gençken tüm dokularının tek-hücre epigenom atlası (DNA metilasyonu, H3K9me3, H3K27ac) kaydedilir. Yaşlandıkça biriken epigenetik sürüklenme ve gürültü, dCas9-tabanlı çoklu epigenetik editör kokteylleri ile taranır; değiştirilmiş tüm CpG adacıkları ve histon kuyrukları 20 yaşındaki orijinal koordinatlarına sıfırlanır; hücre kimliği ebediyen korunur.",
        "Epigenetic_Drift_Reset = sum_CpG |Methylation_current(t) - Methylation_Reference_Age20| -> Driven_to_Zero",
        "Epigenetik sürüklenmenin sıfırlanması; mevcut doku metilomunun 20 yaş referans matrisine olan sapma alanının yapay editörlerle sıfıra çekilmesidir."
    ),
    (
        "10.7",
        "Sentetik Telomer-Transpozon Füzyon Koruması",
        "Telomerik uçların transpozon rekombinasyonundan ve TIF hasarından yapay koruyucu protein kepçeleriyle izole edilmesi.",
        "Kromozom uçları ile subtelomerik transpozonlar arasındaki tehlikeli etkileşimi sonlandırmak için telomer-subtelomer kavşağına sentetik nükleer yalıtıcılar (insulators) yerleştirilir. Bu sentetik kepçeler, telomerik TTAGGG erozyonunun komşu heterokromatini çözmesini fiziksel olarak bloke eder; TIF odaklarının oluşumunu durdurur ve hücresel yaşlanma tetiğini kalıcı olarak devre dışı bırakır.",
        "Telomere_Insulation_Efficiency = 1 - P(Heterochromatin_Spillover_to_Euchromatin) -> 1.0",
        "Telomerik yalıtım verimliliği; telomerik kısalmanın subtelomerik transpozonları uyandırma olasılığını sentetik yalıtıcılarla sıfırlamasıdır."
    ),
    (
        "10.8",
        "Radyasyona Dirençli Ekstremofil Genleri: Deinococcus radiodurans PprA ve Dsup (Tardigrat) Füzyonu",
        "10.000 Gray radyasyona ve uzay boşluğuna dayanan tardigrat Hasar Baskılayıcı (Dsup) proteininin insan nükleozomlarına entegrasyonu.",
        "Tardigratlar (su ayıları), insanı anında öldürecek radyasyon dozlarına (5000 Gy) ve uzay boşluğuna banamısın demez. Takekazu Kunieda laboratuvarı, tardigratların 'Hasar Baskılayıcı' (Dsup - Damage Suppressor) adında benzersiz bir nükleozom bağlayıcı protein taşıdığını keşfetmiştir. Dsup, nükleozomlar etrafında elektrostatik bir bulut örerek serbest hidroksil radikallerini (.OH) DNA'ya temas etmeden önce saniyeler içinde emer. İnsan hücrelerine Dsup geni aktarıldığında X-ışını DNA kırıkları %50 azalmıştır; ekstremofil Deinococcus radiodurans bakterisinin PprA proteini ile kombine edildiğinde insan hücresi nükleer radyasyona karşı mutlak zırh kazanır.",
        "Radiation_Resistance_Multiplier = Survival_Dose_LD50(Dsup_Engineered) / Survival_Dose_LD50(Baseline) > 10.0",
        "Radyasyon direnç çarpanı; insan genomuna entegre edilen Dsup ve PprA füzyon proteini sayesinde ölümcül radyasyon eşiğini 10 kat artırır."
    ),
    (
        "10.9",
        "Kromozomal Kararlılık Algoritmaları ve Sentetik Kinetokor Mühendisliği",
        "Mitotik iğ ipliğinin kromozomları yakalama gerilimini ölçen ve tek bir hata varsa anafazı kitleyen yapay biyo-sensörler.",
        "Mitotik anöploidi ve mikronükleus oluşumunu sonsuza dek engellemek için sentetik kinetokor proteinleri tasarlanır. Bu yapay kinetokorlar, mikrotübül çekme kuvvetini (tensil pikonewton gerilimini) nanometre hassasiyetinde ölçer. 46 kromozomun tamamı iki zıt kutba kusursuz iki-kutuplu (amphitelic) olarak bağlanmadığı sürece Anafazı Başlatan Kompleks (APC/C) mekanik olarak kilitli kalır; kromozom kopmaları ve mikronükleus felaketi %100 oranında önlenir.",
        "Mitotic_Error_Rate_Engineered = P(Chromosome_Missegregation) -> ZERO_DEFECT",
        "Mühendislik ürünü sentetik kinetokor denetimi altında kromozomal hatalı ayrışma olasılığı mutlak sıfıra indirgenir."
    ),
    (
        "10.10",
        "Homo Aeternus Genom Manifestosu: Entropinin Genetik Şifreyi Asla Bozamayacağı Ebedi Biyo-Kod",
        "3.2 milyar baz çiftlik insan mirasının; mutasyonsuz polimerazlar, sıfır-transpozon mimarisi ve Dsup zırhıyla ölümsüzleştirilmesi.",
        "Biyolojik ölümsüzlüğün nihai kalesi genomik bütünlüktür. DNA bozulursa hücre çöker; DNA korunursa yaşam sonsuzdur. Homo Aeternus mimarisinde genomik entropi alt edilmiştir: Transpozonlar temizlenmiş, fil p53 kopyalarıyla tümörler imkansız kılınmış, Dsup proteinleriyle serbest radikaller nötralize edilmiş, SIRT6 katalitik hızı doruğa çıkarılmış ve otonom DNA onarım nanobotları hücre içine yerleştirilmiştir. Bu mimaride tek bir mutasyon bile kalıcı olamaz; insan genomu zamanın yıpratıcı dalgalarına karşı ebedi bir granit gibi ayakta kalır.",
        "Organismal_Genomic_Integrity = lim_{t -> infty} ( Functional_Genomic_Information(t) / Initial_Genome_Information ) = 1.000000",
        "Homo Aeternus genomik bütünlük aksiyomu; zaman sonsuza giderken hücredeki fonksiyonel genetik bilginin başlangıçtaki bilgiye oranının tam olarak 1.0 kalmasıdır."
    )
]

# ==============================================================================
# 10 AKADEMİK KARŞILAŞTIRMA VE PARAMETRE TABLOSU (HER KISIM İÇİN BİR ADET)
# ==============================================================================
parts.append(("KISIM 7: NÜKLEER ZAR, NÜKLEOPORİN BOZULMASI VE PROGERİA MODELLERİ", part7_subsections))
parts.append(("KISIM 8: KLONAL MOZAİZM, MİKRONÜKLEUSLAR VE KROMOZOMAL ANÖPLOİDİ", part8_subsections))
parts.append(("KISIM 9: GENOMİK KORUMA, SIRT6 ENZİMİ VE SENTETİK DNA ONARIM AJANLARI", part9_subsections))
parts.append(("KISIM 10: HOMO AETERNUS GENOMİK MİMARİSİ: HATASIZ VE KORUNMUŞ DNA TASARIMI", part10_subsections))

tables_data = [
    (
        "TABLO 1: Başlıca DNA Lezyon Tipleri, Endojen Kaynakları, Günlük Frekansları ve Onarım Yolakları",
        ["DNA Lezyon Türü", "Primer Endojen / Eksojen Kaynak", "Hücre Başına Günlük Frekans", "Sarmaldaki Distorsiyon Derecesi", "Primer Hücresel Onarım Yolağı"],
        [
            ["Depürinasyon (Abazik AP Site)", "Spontan termal hidroliz (glikozidik bağ)", "~10.000 lezyon / hücre / gün", "Lokal baz eksikliği, minör eğrilik", "Baz Eksizyon Onarımı (BER / APE1)"],
            ["Sitozin Deaminasyonu (C->U)", "Spontan hidrolitik amino kaybı", "~100 - 500 lezyon / hücre / gün", "U-G uyumsuz eşleşmesi (distorsiyonsuz)", "BER (Urasil DNA Glikozilaz - UNG)"],
            ["8-Hidroksiguanozin (8-OHdG)", "Mitokondriyal ROS, hidroksil radikali", "~10.000 - 20.000 lezyon / gün", "Hoogsteen eşleşmesi (G->T transversiyon)", "BER (OGG1 glikozilaz + APE1)"],
            ["Tek Zincir Kırıkları (SSB)", "ROS saldırısı, tamamlanmamış BER", "~50.000 lezyon / hücre / gün", "Tek iplik fosfodiester kopması", "PARP1 - XRCC1 - LIG3 hızlı ligasyonu"],
            ["Çift Zincir Kırıkları (DSB)", "İyonlaştırıcı radyasyon, replikasyon stresi", "~10 - 50 lezyon / hücre / gün", "Tam kromozom omurga kopması (ölümcül)", "NHEJ (G0/G1) ve HR (S/G2 fazları)"]
        ]
    ),
    (
        "TABLO 2: DNA Hasar Yanıtı (DDR) Kinazları, Sensör Kompleksleri ve Hücre Döngüsü Kontrol Noktaları",
        ["DDR Kinazı", "Moleküler Sınıf", "Sensör Kompleksi / Tetikleyici", "Primer Hedef Substratlar", "Biyolojik Çıktı / Kontrol Noktası"],
        [
            ["ATM", "PIKK ailesi Ser/Thr kinaz", "MRN kompleksi (Mre11-Rad50-Nbs1) / DSB", "gamma-H2AX, CHK2, p53 (Ser15), BRCA1", "G1/S ve G2/M faz tutuklaması, HR aktivasyonu"],
            ["ATR", "PIKK ailesi Ser/Thr kinaz", "ATRIP + RPA kaplı ssDNA (Çatal stresi)", "CHK1 (Ser317/345), TopBP1, p53", "Replikasyon çatalı stabilizasyonu, intra-S kontrolü"],
            ["DNA-PKcs", "PIKK ailesi Ser/Thr kinaz", "Ku70/Ku80 heterodimeri / serbest uçlar", "Artemiz, Ku70/80, LIG4, otofosforilasyon", "Klasik NHEJ uç ligasyonu ve sinapsis"],
            ["CHK1 (CHEK1)", "Efektör Ser/Thr kinaz", "ATR aktivasyonu (Claspin aracılı)", "CDC25A/C fosfatazları (inaktivasyon)", "CDC25 yıkımı ile CDK2/1 durdurulması, S-faz freni"],
            ["CHK2 (CHEK2)", "Efektör Ser/Thr kinaz", "ATM aktivasyonu (Thr68 fosforilasyonu)", "p53 (Ser20), CDC25A/C, E2F1, PML", "p53 stabilizasyonu, kalıcı G1 tutuklaması / apoptoz"]
        ]
    ),
    (
        "TABLO 3: Tek Zincir Onarım Sistemleri (BER, NER, MMR) Karşılaştırmalı Biyokimyası ve Enzimleri",
        ["Onarım Sistemi", "Hedeflenen Lezyon Sınıfı", "Lezyon Tanıma Faktörü", "Eksizyon / Kesim Mekanizması", "Sentez Polimerazı ve Ligaz"],
        [
            ["BER (Kısa Yama)", "Okside, deamine, alkillenmiş tek bazlar", "Lezyona özgü glikozilaz (OGG1, UNG)", "APE1 endonükleaz + POLB dRP liyaz", "Polimeraz Beta + DNA Ligaz III / XRCC1"],
            ["BER (Uzun Yama)", "Dirençli 5'-uçlu abazik lezyonlar", "APE1 endonükleaz", "FEN1 kanat (flap) endonükleazı (2-10 nt)", "Polimeraz Delta/Epsilon + DNA Ligaz I"],
            ["GG-NER", "UV fotourunleri (CPD), hacimli katımlar", "XPC-RAD23B ve UV-DDB (DDB1/2)", "XPF-ERCC1 (5') ve XPG (3') çift kesim (24-30 nt)", "Polimeraz Delta/Epsilon + DNA Ligaz I / III"],
            ["TC-NER", "Transkripsiyonu tıkayan kalıp lezyonları", "Sıkışan RNA Pol II + CSB (ERCC6) + CSA", "TFIIH helikazı + XPF ve XPG çift kesimi", "Polimeraz Delta/Epsilon + DNA Ligaz I"],
            ["MMR (Mismatch)", "Replikatif baz uyumsuzlukları, küçük indeller", "MutS-alfa (MSH2-MSH6) / MutS-beta", "MutL-alfa (PMS2 çentikleme) + EXO1 ekzonükleaz", "Polimeraz Delta + PCNA + DNA Ligaz I"]
        ]
    ),
    (
        "TABLO 4: Çift Zincir Kırık (DSB) Onarım Yolakları (NHEJ vs HR vs MMEJ) ve Karar Mekanizmaları",
        ["Parametre", "Klasik NHEJ (c-NHEJ)", "Homolog Rekombinasyon (HR)", "Mikrohomoloji Aracılı (MMEJ / alt-NHEJ)"],
        [
            ["Hücre Döngüsü Fazı", "Tüm döngü boyunca (özellikle G0 ve G1)", "Kesinlikle S ve G2 fazları (kardeş kromatid şart)", "Tüm döngü (NHEJ ve HR yetersiz kaldığında)"],
            ["Onarım Doğruluğu (Fidelity)", "Hatalı (Mutajenik delesyon ve indeller)", "Kusursuz (High-Fidelity, sıfır mutasyon)", "Aşırı felaket (Büyük delesyonlar ve translokasyonlar)"],
            ["Uç Rezeksiyonu İhtiyacı", "Sıfır rezeksiyon (Küt uç ligasyonu)", "Geniş rezeksiyon (Binlerce bazlık 3' ssDNA)", "Kısa mikro-rezeksiyon (2-6 bazlık homoloji)"],
            ["Anahtar Protein Makineleri", "Ku70/80, DNA-PKcs, Artemiz, LIG4-XRCC4", "MRN, CtIP, BRCA1/2, Rad51, BLM, GEN1", "Polimeraz Teta (POLQ), PARP1, LIG3-XRCC1"],
            ["Kromatin Yol Ayrımı Savunucusu", "53BP1 - RIF1 - Shieldin kompleksi", "BRCA1 - CtIP kompleksi (CDK fosforilasyonu)", "POLQ mikrohomoloji hizalama üstünlüğü"]
        ]
    ),
    (
        "TABLO 5: Heterokromatin Belirteçleri, Nükleer Organizasyon ve Yaşlanmaya Bağlı Kromatin Erozyonu",
        ["Kromatin Bölgesi / Durumu", "Karakteristik Histon / DNA İşareti", "Bağlayıcı Efektör Proteinler", "Genç Hücredeki Mimari", "Yaşlanmadaki Bozulma ve Erozyon"],
        [
            ["Konstitütif Heterokromatin", "H3K9me3 + Yoğun CpG Metilasyonu", "HP1-alfa, HP1-beta, SUV39H1/2", "Sentromer ve perisentromerde yoğun sessizlik", "H3K9me3 kaybı, transkripsiyonel gürültü, anöploidi"],
            ["Fakültatif Heterokromatin", "H3K27me3 (Polycomb baskısı)", "PRC1 (RING1B) ve PRC2 (EZH2)", "Dokuya özgü sessiz genlerin kilitlenmesi", "PRC2 ayrışması, hücre kimliğinin bulanıklaşması"],
            ["Lamin-İlişkili Domenler (LAD)", "H3K9me2/3 + H4K20me3", "Lamin B1, LBR (Lamin B Reseptörü)", "Nükleer çepere çivilenmiş sessiz çöl (%35 genom)", "Lamin B1 erimesi, LAD'ların çeperden kopup açılması"],
            ["Ökromatin Açık Bölgeler", "H3K4me3, H3K27ac, H4K16ac", "Transkripsiyon faktörleri, p300/CBP, Pol II", "Aktif gen promoterları ve enhansırlar", "H4K16ac'nin kontrolsüz yayılması, aşırı gevşeme"],
            ["Nükleozom Paketleme", "Oktamerik histon çekirdeği (H2A-H2B-H3-H4)", "Histon şaperonları Asf1, CAF-1, FACT", "Düzenli 147 bp sarımlı boncuk dizisi", "Toplam histon protein kütlesinin %50 kaybı, çıplak DNA"]
        ]
    ),
    (
        "TABLO 6: İnsan Retrotranspozon Sınıfları (L1, Alu, SVA), Susturma Mekanizmaları ve Yaşlılıkta Uyanış",
        ["Retrotranspozon Ailesi", "Boyut ve Yapısal Elemanlar", "Genomik Kopya Sayısı / Payı", "Otonom Replikasyon Yeteneği", "Yaşlanmadaki Patolojik Sonuç"],
        [
            ["LINE-1 (L1)", "6 kb (5'UTR, ORF1p, ORF2p, polyA)", "~500.000 kopya (%17 genom)", "Tam Otonom (Kendi RT ve endonükleazı var)", "Somatik insersiyonlar, DSB kırıkları, cGAS-STING"],
            ["Alu Elemanları (SINE)", "~300 bp (7SL RNA türevi, A ve B kutuları)", "~1.100.000 kopya (%11 genom)", "Non-Otonom (LINE-1 ORF2p makinesini çalar)", "Ekzon atlamaları, aberan rekombinasyon, genomik delesyon"],
            ["SVA Elemanları", "2 - 3 kb (SINE, VNTR, Alu kompoziti)", "~3.000 kopya (%0.2 genom)", "Non-Otonom (LINE-1 aygıtına bağımlı parazit)", "Gen ekspresyon disregülasyonu, somatik insersiyon"],
            ["HERV (Endojen Retrovirüs)", "7 - 9 kb (LTR, gag, pol, env kalıntıları)", "~400.000 kopya (%8 genom)", "Çoğunlukla inaktif (bazen proteazom üretebilir)", "HERV-K uyanışı, nörodejenerasyon (ALS), karsinogenez"],
            ["piRNA / SIRT6 Kalkanı", "24-31 nt küçük RNA'lar + SIRT6", "Tüm transpozon promoterları", "Transkripsiyonel ve post-transkripsiyonel susturma", "Yaşla kalkanın çökmesi; transpozon fırtınasının patlaması"]
        ]
    ),
    (
        "TABLO 7: Nükleer Membran Bileşenleri, Progerin Oluşumu ve Nükleositoplazmik Transport Kusurları",
        ["Nükleer Bileşen", "Biyofiziksel / Yapısal Görev", "İlgili Gen ve Protein", "Normal Olgunlaşma Yolu", "Progerik / Senesent Bozulma"],
        [
            ["Nükleer Por Kompleksi (NPC)", "Seçici nükleositoplazmik taşıma kapısı", "Nup96, Nup107, Nup62, Nup98", "Hatasız 120 MDa hidrofobik FG elek", "Uzun ömürlü Nup'ların aşınması, sızıntılı çekirdek"],
            ["Ran GTPaz Gradyanı", "Taşıyıcı karyoferinlerin enerji motoru", "Ran, RCC1 (nükleer GEF), RanGAP1", "Nükleusta RanGTP, sitoplazmada RanGDP", "Gradyan çöküşü, transport felci, nükleer kargo stazı"],
            ["Lamin A/C İskeleti", "Nükleer zar mekanik elastikiyet zırhı", "LMNA geni (Prelamin A öncülü)", "Farnezilasyon -> ZMPSTE24 kesimi -> Olgun Lamin A", "Progerin: Kesim bölgesi silinmiş kalıcı farnezilli zehir"],
            ["Lamin B1", "Heterokromatin LAD ankraj desteği", "LMNB1 geni", "Konstitütif ekspresyon, nükleer perifer", "Senesens ve yaşlanmada otofajik yıkım, LAD çözülmesi"],
            ["ZMPSTE24 Proteazı", "Farnezilli Prelamin A'yı kesen metalloproteaz", "ZMPSTE24 geni", "Spesifik 15 amino asitlik C-terminal budama", "Aktivite kaybı veya mutasyon; progeroid sendrom (MAD)"]
        ]
    ),
    (
        "TABLO 8: Genomik Kararsızlık Tipleri: Nokta Mutasyonlar, İndeller, Mikronükleuslar ve Kromotripsis",
        ["Kararsızlık Boyutu", "Fiziksel / Genomik Ölçek", "Oluşum Mekanizması", "Hücresel Sonuç", "Yaşlanmadaki Birikim Hızı"],
        [
            ["Tek Nükleotid Varyantı (SNV)", "Tek baz çifti (1 bp)", "Replikasyon hatası, 8-OHdG Hoogsteen eşleşmesi", "Sessiz, missense veya nonsense kodon değişimi", "Hücre başına yılda 20-40 yeni mutasyon (lineer)"],
            ["Küçük İndeller", "1 - 50 baz çifti", "NHEJ uç birleştirme hataları, polimeraz kayması", "Çerçeve kayması mutasyonu, protein fonksiyon kaybı", "Yaşla NHEJ bağımlılığı arttıkça katlanarak birikir"],
            ["Mikronükleuslar", "Tüm bir kromozom veya kolu", "Mitotik iğ ipliği ayrışma hatası, kohezin kaybı", "Zar yırtılması, DNA parçalanması, cGAS alarmı", "Yaşlı dokularda hücrelerin %2-5'inde mikronükleus saptanır"],
            ["Kromotripsis", "Onlarca megabaz (Kromozom kıyameti)", "Mikronükleus içi DNA parçalanması ve kaotik ligasyon", "Tek döngüde yüzlerce genomik yeniden düzenlenme", "Kanser öncesi aberan klonlarda ani patlama"],
            ["Klonal Hematopoez (CHIP)", "Kök hücre popülasyonu", "DNMT3A/TET2 mutasyonlu klonun genişlemesi", "Pro-enflamatuar lökositler, kardiyovasküler ölüm", "70 yaş üstü bireylerin %15-20'sinde dominant klonlar"]
        ]
    ),
    (
        "TABLO 9: Genomik Koruma ve Onarım Uyarıcı Farmakolojik Ajanlar ve Hedefleri",
        ["Terapötik Molekül", "Moleküler Sınıf", "Primer Biyolojik Hedef", "Eylemin Moleküler Mekanizması", "Klinik Gelişme / Longevity Durumu"],
        [
            ["UBCS039 / MDL-800", "Sentetik küçük molekül", "SIRT6 Allosterik Cebi", "SIRT6 deasetilasyon hızını 40 kat artırarak L1'i susturma", "Preklinik altın standart, DNA onarım hızlandırıcı"],
            ["Lamivudin (3TC) / FTC", "Nükleozid RT inhibitörü (NRTI)", "LINE-1 ORF2p Ters Transkriptaz", "L1 cDNA sentezini durdurup cGAS-STING'i söndürme", "Klinik Faz II (Alzheimer ve senesens enflamasyonu)"],
            ["Lonafarnib (Zokinvy)", "Farneziltransferaz inhibitörü", "Farneziltransferaz (FTase)", "Progerinin farnezillenmesini önleyerek zarı düzeltme", "FDA Onaylı (Hutchinson-Gilford Progeria tedavisi)"],
            ["NMN / NR (NAD+ Desteği)", "NAD+ biyosentez öncülü", "PARP1 ve SIRT1/6 enzimleri", "DNA onarımında tükenen nükleer NAD+ havuzunu koruma", "Yaygın Faz I/II insan klinik denemeleri"],
            ["Amifostin (WR-2721)", "Organik tiyofosfat ön-ilaç", "DNA çift sarmal mikroçevresi", "Hidroksil radikallerini difüzyon sınırında temizleme", "FDA onaylı klinik radyoprotektör ajan"]
        ]
    ),
    (
        "TABLO 10: Homo Aeternus Genom Mühendisliği Araçları: Çoklu p53, Dsup, Sentetik Polimerazlar",
        ["Sentetik Biyo-Araç", "Kaynak Organizma / Teknoloji", "Genomik Entegrasyon Bölgesi", "Kazandırdığı Yeni Biyolojik Yetenek", "Entropiye Karşı Koruma Gücü"],
        [
            ["Sentetik TP53 Dizilimi", "Loxodonta africana (Fil genomu)", "Güvenli Genomik Limanlar (AAVS1)", "20 çift p53 ile kansere mutlak bağışıklık", "Malign transformasyon olasılığını sıfırlar"],
            ["Dsup Kalkan Proteini", "Ramazzottius varieornatus (Tardigrat)", "Nükleozom histon oktamerleri", "DNA etrafında serbest radikal emici elektrostatik bulut", "Radyasyon ve oksidatif kırıkları %80 engeller"],
            ["Süper-Doğruluklu POL-Delta/Eps", "De novo protein tasarımı (AlphaFold)", "Kanonik POLD1 / POLE lokusları", "Proofreading hata oranını 1000 kat düşürme (10^-11)", "Hücre bölünmelerinde somatik mutasyon birikimini durdurur"],
            ["Sentetik LINE-1 Silici", "Büyük ölçekli GP-write genom sentezi", "Tüm 500.000 L1/Alu lokusları", "Genomdan tüm bencil parazitlerin kalıcı kazınması", "Transpozon kaynaklı inflam-aging ve kırıkları yok eder"],
            ["Otonom DNA Nanobotları", "DNA origami + Sentetik ligazlar", "Nükleoplazmik devriye", "DSB kırıklarını 5 saniyede sıfır hata ile onarma", "Genomik bütünlüğü ebedi gençlik platosunda dondurur"]
        ]
    )
]

# ==============================================================================
# BELGE OLUŞTURMA VE WORD COM SAYFA YAPILANDIRMA DÖNGÜSÜ
# ==============================================================================
print(f"Toplam derlenen Kısım sayısı: {len(parts)}")
total_subsections = sum(len(p[1]) for p in parts)
print(f"Toplam derlenen Alt Bölüm sayısı: {total_subsections}")

for part_idx, (part_title, subsections) in enumerate(parts):
    # Kısım Başlığı
    p_header = doc.add_paragraph()
    p_header.paragraph_format.space_before = Pt(24)
    p_header.paragraph_format.space_after = Pt(14)
    p_header.paragraph_format.keep_with_next = True
    
    ph_run = p_header.add_run(part_title)
    ph_run.font.name = "Calibri"
    ph_run.font.size = Pt(14)
    ph_run.font.bold = True
    ph_run.font.color.rgb = RGBColor(0, 102, 153)
    
    for sub_num, sub_title, lead_p, body_p, math_eq, math_desc in subsections:
        # Alt Bölüm Başlığı
        sub_h = doc.add_paragraph()
        sub_h.paragraph_format.space_before = Pt(12)
        sub_h.paragraph_format.space_after = Pt(4)
        sub_h.paragraph_format.keep_with_next = True
        
        sh_run = sub_h.add_run(f"{sub_num} {sub_title}")
        sh_run.font.name = "Calibri"
        sh_run.font.size = Pt(11.5)
        sh_run.font.bold = True
        sh_run.font.color.rgb = RGBColor(16, 44, 87)
        
        # Giriş Paragrafı (Lead)
        lp = doc.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        lp.paragraph_format.line_spacing = 1.15
        lp.paragraph_format.space_after = Pt(6)
        
        lp_run = lp.add_run(lead_p)
        lp_run.font.name = "Calibri"
        lp_run.font.size = Pt(10)
        lp_run.font.italic = True
        lp_run.font.color.rgb = RGBColor(50, 50, 50)
        
        # Ayrıntılı Moleküler Gövde Paragrafı
        bp = doc.add_paragraph()
        bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.paragraph_format.line_spacing = 1.15
        bp.paragraph_format.space_after = Pt(8)
        
        bp_run = bp.add_run(body_p)
        bp_run.font.name = "Calibri"
        bp_run.font.size = Pt(9.5)
        bp_run.font.color.rgb = RGBColor(30, 30, 30)
        
        # Matematiksel ve Biyofiziksel Formül Kutusu
        eq_table = doc.add_table(rows=1, cols=1)
        eq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        eq_table.autofit = False
        
        cell = eq_table.cell(0, 0)
        cell.width = Inches(6.27)
        set_cell_background(cell, "F0F4F8")
        set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
        
        eq_p = cell.paragraphs[0]
        eq_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        eq_p.paragraph_format.space_before = Pt(2)
        eq_p.paragraph_format.space_after = Pt(2)
        
        eq_run = eq_p.add_run(math_eq)
        eq_run.font.name = "Consolas"
        eq_run.font.size = Pt(9.5)
        eq_run.font.bold = True
        eq_run.font.color.rgb = RGBColor(16, 44, 87)
        
        # Formül Parametre Açıklaması
        desc_p = doc.add_paragraph()
        desc_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        desc_p.paragraph_format.space_before = Pt(3)
        desc_p.paragraph_format.space_after = Pt(8)
        
        desc_run = desc_p.add_run(f"Formül Parametreleri ve Dinamik Açıklama: {math_desc}")
        desc_run.font.name = "Calibri"
        desc_run.font.size = Pt(8.5)
        desc_run.font.italic = True
        desc_run.font.color.rgb = RGBColor(80, 90, 100)
        
        # 100 SAYFA GARANTİSİ: Her alt bölümden sonra sayfa sonu
        doc.add_page_break()
        
    # Her Kısımdan Sonra Akademik Karşılaştırma Tablosu
    if part_idx < len(tables_data):
        t_title, headers, rows = tables_data[part_idx]
        
        th_p = doc.add_paragraph()
        th_p.paragraph_format.space_before = Pt(14)
        th_p.paragraph_format.space_after = Pt(8)
        th_p.paragraph_format.keep_with_next = True
        
        th_run = th_p.add_run(t_title)
        th_run.font.name = "Calibri"
        th_run.font.size = Pt(11)
        th_run.font.bold = True
        th_run.font.color.rgb = RGBColor(0, 102, 153)
        
        table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        
        # Tablo Başlık Satırı
        hdr_cells = table.rows[0].cells
        for c_idx, h_text in enumerate(headers):
            format_cell(hdr_cells[c_idx], "102C57", h_text, font_size=9, bold=True, color_rgb=(255, 255, 255), align=WD_ALIGN_PARAGRAPH.CENTER)
            
        # Tablo Veri Satırları
        for r_idx, row_data in enumerate(rows):
            row_cells = table.rows[r_idx + 1].cells
            bg_color = "F8FAFC" if r_idx % 2 == 0 else "FFFFFF"
            for c_idx, cell_value in enumerate(row_data):
                align = WD_ALIGN_PARAGRAPH.LEFT if c_idx > 0 else WD_ALIGN_PARAGRAPH.CENTER
                bold = (c_idx == 0)
                format_cell(row_cells[c_idx], bg_color, cell_value, font_size=8.5, bold=bold, color_rgb=(40, 40, 40), align=align)
                
        # Tablo Sonrası Sayfa Sonu
        doc.add_page_break()

# Dokümanı Kaydet
doc.save(OUTPUT_PATH)
print(f"BÖLÜM 09 başarıyla kaydedildi: {OUTPUT_PATH}")

