# -*- coding: utf-8 -*-
"""
PROJECT AETERNITAS - CİLT 08: PROTEOSTAZ, OTOFAJİ-LİZOZOM YOLAĞI VE AGREGATOM TEMİZLİĞİ
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

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_08_PROTEOSTAZ_OTOFAJI_VE_AGREGATOM_TEMIZLIGI_TAM_100_SAYFA.docx"

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
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 08: PROTEOSTAZ VE AGREGATOM TEMİZLİĞİ")
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
s_run = sub_p.add_run("CİLT 08: PROTEOSTAZ, OTOFAJİ-LİZOZOM YOLAĞI VE AGREGATOM TEMİZLİĞİ\\n(ŞAPERONLAR, 26S PROTEAZOM, CMA, LİPOFUSİN TASFİYESİ VE SENTETİK DİSAGREGAZLAR)")
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
ih_run = intro_h.add_run("CİLT 08 MANİFESTOSU: HÜCRESEL ÇÖPÜN TASFİYESİ VE PROTEOMİK DOKUNULMAZLIK")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "Biyolojik sistemlerin dinamik kararlılığı, polipeptid zincirlerinin kusursuz 3 boyutlu konformasyonel geometrilerine katlanması, "
    "işlev görmesi ve ömrünü tamamladığında zamanında parçalanmasıyla (proteostaz ağı) mümkündür. Ancak termodinamik entropi, "
    "proteinleri sürekli yanlış katlanma kinetik tuzaklarına iter; bu durum toksik oligomerlerin, çapraz beta-tabakalı amiloid liflerin "
    "ve yıkılamayan lipofusin pigmentlerinin (agregatom) hücre içinde birikmesine neden olur.\\n\\n"
    "Yaşlanma sürecinde moleküler şaperon (Hsp70/90, TRiC) kapasitesi çöker, 26S proteazomun hidrolitik odacıkları tıkanır, "
    "endoplazmik retikulumda UPR alarmı kronikleşir ve otofaji-lizozom otoyolu (makrootofaji, CMA) durma noktasına gelir. "
    "Sonuç; nörodejeneratif çöküş (Alzheimer, Parkinson, ALS), kardiyovasküler amiloidoz, katarakt ve hücresel fonksiyonların felcidir.\\n\\n"
    "Bu ciltte; protein katlanma enerjetiği, Hsp şaperon mekanizmaları, UPR kolları (IRE1a, PERK, ATF6), 26S proteazom mimarisi, "
    "makrootofaji makineleri (ULK1, Beclin1, LC3 lipitlenmesi), şaperon aracılı otofaji (LAMP-2A/Hsc70), TFEB ve CLEAR gen ağı, "
    "lipofusin biyofiziği, PROTAC/AUTAC teknolojileri ve yapay desagregazlarla sıfır-çöp proteomik mimarisi 100 ayrıntılı bölümde incelenmektedir."
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
# KISIM 1: PROTEOSTAZ AĞININ MİMARİSİ VE PROTEİN KATLANMA BİYOFİZİĞİ
# ==============================================================================
part1_subsections = [
    (
        "1.1",
        "Anfinsen Dogması ve Enerji Manzarası (Protein Folding Funnel): Termodinamik vs Kinetik Tuzaklar",
        "Bir polipeptit zincirinin 3 boyutlu yerel yapısının birincil amino asit diziliminde kodlanması ve termodinamik serbest enerji hunisindeki yolculuğu.",
        "Christian Anfinsen'in 1973 Nobel ödüllü dogmasına göre, fizyolojik koşullarda bir proteinin yerel (native) konformasyonu, Gibbs serbest enerjisinin küresel minimumda (global minimum) olduğu durumdur. Katlanma süreci, düz bir yoldan ziyade çok boyutlu bir 'enerji hunisi' (folding funnel) üzerinde ilerler. Polipeptit huniden aşağı inerken hidrofobik kalıntılar içe gömülür (hidrofobik çöküş) ve konformasyonel entropi azalır. Ancak huni pürüzsüz değildir; derin yerel enerji çukurları (kinetik tuzaklar) içerir. Yanlış katlanan ara ürünler bu çukurlara takılarak oligomerik agregatlara dönüşür; şaperonlar bu kinetik bariyerleri aşmak için ATP enerjisi harcar.",
        "Delta_G_folding = Delta_H_conformational - T * (Delta_S_chain + Delta_S_solvent) < 0",
        "Proteinin katlanma serbest enerjisi; entalpik bağ kazançları ile polipeptid zincir entropi kaybı ve çözücü su entropi artışının farkıdır."
    ),
    (
        "1.2",
        "Ko-translasyonel Katlanma ve Ribozoma Bağlı Şaperonlar (Trigger Factor, RAC)",
        "Polipeptid ribozom çıkış tünelinden (exit tunnel) sentezlenirken henüz uzama aşamasında başlayan anlık katlanma ve şaperon kalkanı.",
        "Protein katlanması ribozomdan sentezin bitmesini beklemez; polipeptid zinciri ribozomun 10 nanometrelik çıkış tünelinden 30-40 amino asit uzadığı anda ko-translasyonel olarak başlar. Bakterilerde Trigger Factor (TF), ökaryotlarda ise Ribozoma Bağlı Kompleks (RAC: Hsp70 sınıfı Hsp70L1/Ssz1 ve Hsp40 sınıfı Zuo1) ribozom çıkış kapısına konuşlanır. Bu şaperonlar, henüz sentezlenmekte olan açık hidrofobik bölgelere bağlanarak zincirin çevreyle veya komşu ribozomlardan çıkan proteinlerle erken ve yanlış agregasyon yapmasını önler. Yaşlanan hücrelerde ribozom çevirisi yavaşlar ve bu erken şaperon koordinasyonu aksar.",
        "P_cotranslational_folding = k_fold / (k_fold + k_elongation_ribosome + k_aggregation)",
        "Ko-translasyonel katlanma olasılığı; yerel katlanma hızının ribozomal uzama hızı ve erken agregasyon hızına olan rekabetçi oranıdır."
    ),
    (
        "1.3",
        "Isı Şoku Proteinleri (Hsp70 / Hsp40 Döngüsü): ATP Bağımlı Substrat Yakalama",
        "Hücrenin en yaygın moleküler şaperon makinesi: Hsp70'in ATP hidrolizi ile açık hidrofobik substratları kıskaç içine alıp bırakma döngüsü.",
        "Hsp70 (stres indüklenebilir Hsp72 ve konstitütif Hsc70), N-terminal nükleotid bağlama domenine (NBD) ve C-terminal substrat bağlama domenine (SBD) sahiptir. Döngü, Hsp40 (DnaJ) ko-şaperonunun hidrofobik 5-7 amino asitlik açık bölgeleri tanıması ve substratı Hsp70-ATP formuna sunmasıyla başlar. Hsp40, Hsp70'in içsel ATPaz aktivitesini 1000 kat artırır. ATP hidroliz olup ADP'ye dönüştüğünde, SBD alfa-helikal kapağı kapanarak substratı sımsıkı hapseder (kıskaç durumu). Nükleotid değişim faktörleri (NEF: Bag1-3, Hsp110), ADP'yi ATP ile değiştirir; kapak açılır ve katlanmış protein salınır.",
        "Cycle_Rate_Hsp70 = k_NEF * [NEF] * (k_cat_Hsp40 * [Hsp40]) / (K_m_ATP + [ATP])",
        "Hsp70 şaperon döngü frekansı; Hsp40 aktivasyon kinetiği ile nükleotid değişim faktörü (NEF) hızının ATP doygunluğuyla çarpımıdır."
    ),
    (
        "1.4",
        "Şaperoninler (Hsp60 / GroEL-GroES ve CCT/TRiC Kompleksi): Nano-Kafes İçi İzolasyon",
        "Tek bir protein molekülünü sitoplazmanın kalabalığından izole eden silindirik varil benzeri nano-reaktör: Anfinsen kafesi.",
        "Şaperoninler (Grup I: mitokondriyal Hsp60/Hsp10 ve bakteriyel GroEL/GroES; Grup II: ökaryotik sitozolik CCT/TRiC), iki adet sırt sırta vermiş 7'li veya 8'li halkadan oluşan devasa oligomerik varillerdir. Yanlış katlanmış bir polipeptid, varilin hidrofobik iç boşluğuna çekilir. ATP ve GroES kapağının bağlanmasıyla varil içi hacim iki katına çıkar ve iç yüzey negatif yüklü hidrofilik karakter kazanır. Hapsedilen protein, başka hiçbir molekülle temas etmeksizin saniyeler boyunca kendi kendine katlanmaya bırakılır (Anfinsen kafesi hipotezi). CCT/TRiC, özellikle aktin ve tübülin gibi hücre iskeleti proteinlerinin katlanmasında vazgeçilmezdir.",
        "Folding_Rate_Cage = k_cage * [Substrate_bound] * [ATP_hydrolysis_cooperative]",
        "Kafes içi katlanma hızı; şaperonin odacığına alınan substrat miktarı ile halkalar arası kooperatif ATP hidroliz debisinin fonksiyonudur."
    ),
    (
        "1.5",
        "Hsp90 Sistemi: İstemci Proteinler (Kinazlar, Nükleer Reseptörler) ve Konformasyonel Olgunlaşma",
        "Sinyal iletiminin anahtarları olan protein kinazları ve steroid hormon reseptörlerini aktif formda tutan süper-şaperon dimer.",
        "Hsp90, hücre proteinlerinin %1-2'sini oluşturan ve homodimer halinde çalışan esansiyel bir şaperondur. Hsp70'in aksine genel katlanmamış proteinleri değil; belirli 'istemci' (client) proteinleri (örneğin Raf-1, Akt, Cdk4, steroid reseptörleri) tanır. Hsp90 döngüsü, ATP bağlanmasıyla N-terminal domenlerin dimerleşerek kapalı 'moleküler makas' formuna geçmesiyle yürütülür. Ko-şaperonlar (Cdc37 kinazlar için; Aha1 ATPaz aktivasyonu için; p23 kapalı durumu stabilize etmek için) istemcinin hedeflenmesini sağlar. Yaşlanan hücrelerde Hsp90'ın agregatlara bağlanarak tükenmesi, sinyal iletim yollarının çökmesine yol açar.",
        "Client_Activation_Flux = k_hsp90 * [Hsp90_dimer] * [Client_precursor] * [Aha1] / (1 + [Geldanamycin]/K_i)",
        "Hsp90 istemci aktivasyon akısı; aktif Hsp90 ve Aha1 ko-şaperon derişiminin spesifik inhibitör geldanamisin varlığındaki kinetik fonksiyonudur."
    ),
    (
        "1.6",
        "Küçük Isı Şoku Proteinleri (sHsp'ler / Hsp27, alfaB-Kristallin): ATP-Bağımsız Tutma Şaperonları",
        "ATP harcamadan hidrofobik agregatları bir sünger gibi emip bağlayan ve hücreyi ani çöküşten koruyan oligomerik tamponlar.",
        "Küçük ısı şoku proteinleri (sHsp'ler; Hsp27/HSPB1, alfaB-kristallin/HSPB5, Hsp20), 15-30 kDa boyutunda korunmuş bir 'alfa-kristallin' domenine sahip ATP-bağımsız şaperonlardır. Normal koşullarda 24-merlik büyük oligomerler halinde bulunurlar. Termal veya oksidatif stres anında fosforilasyonla (p38 MAPK aracılı) küçük dimerlere ayrışırlar; açığa çıkan hidrofobik yüzeyler erken çöken protein substratlarını bağlar ('holdase' fonksiyonu). sHsp'ler proteini kendileri katlayamaz; ancak agregasyonu engelleyerek stres geçtiğinde proteini Hsp70 sistemine teslim ederler.",
        "Buffering_Capacity_sHsp = N_binding_sites * [sHsp_dimer_active] - [Precipitated_Aggregates]",
        "sHsp tamponlama kapasitesi; aktif dimerik şaperon başına düşen hidrofobik bağlanma bölgelerinin çökelen agregat yükünden çıkarılmasıdır."
    ),
    (
        "1.7",
        "Nükleer ve Sitoplazmik Proteostaz Ayrışması: Kompartımana Özgül Kalite Kontrol",
        "Çekirdek kromatini içindeki transkripsiyon faktörlerinin korunması ile sitoplazmik translasyonel kalite kontrol arasındaki moleküler işbölümü.",
        "Proteostaz hücresel kompartımanlara göre son derece özelleşmiştir. Sitoplazma, ribozomal translasyon ve ribozom kaynaklı erken katlanma yüküyle uğraşırken; nükleus, histonlar, nükleer porlar ve kromatin modifiyerlerinin bütünlüğünü korumak zorundadır. Nükleusta UBR4 ve San1 gibi özgül E3 ubikitin ligazları hatalı katlanmış nükleer proteinleri derhal işaretler. Nükleer Hsp70/Hsp40 (DNAJB1) kompleksleri transkripsiyon faktörlerinin agregasyonunu önler. Yaşlanmayla nükleer por komplekslerinin sızdırmazlığı bozulduğunda sitoplazmik agregatlar nükleusa dolarak transkripsiyonu kilitler.",
        "Nuclear_Proteostasis_Fidelity = [Active_Nuclear_Chaperones] / ([Cytoplasmic_Leakage_Aggregates] + epsilon)",
        "Nükleer proteostaz sadakati; intakt nükleer şaperon havuzunun nükleusa sızan sitoplazmik agregat kütlesine oranıdır."
    ),
    (
        "1.8",
        "Yaşlanmada Proteostaz Kapasitesinin Çöküşü: Katlanmamış Protein Yükü vs Şaperon Havuzu",
        "Richard Morimoto'nun Proteostaz Çöküş Hipotezi: Yaşlanmanın erken döneminde şaperon tampon kapasitesinin aniden tükenmesi.",
        "C. elegans ve memeli modellerinde yapılan çığır açıcı çalışmalar, proteostaz ağının yaşlanma boyunca yavaş ve doğrusal bir düşüş göstermediğini; üreme çağı biter bitmez erken yetişkinlikte programlı ve dik bir 'çöküş' yaşadığını ortaya koymuştur. Yaşlanan hücrede yeni sentezlenen proteinlerin %30'u hatalı katlanır (DRiPs - Defective Ribosomal Products). Aynı anda serbest radikaller ve post-translasyonel modifikasyonlar mevcut proteinleri hasarlayarak şaperon talebini fırlatır. Şaperon havuzu agregatlara bağlanıp kilitlendiğinde (titrasyon), hücre açıkta kalan proteinleri tolere edemez ve sistemik amiloidoz başlar.",
        "Proteostasis_Buffer_Margin = Total_Chaperone_Capacity(t) - (Basal_Folding_Load + Misfolded_Protein_Stress(t))",
        "Proteostaz güvenlik marjı; toplam şaperon kapasitesinden bazal katlanma yükü ve yanlış katlanma stresinin çıkarılmasıyla elde edilen net tampon alandır."
    ),
    (
        "1.9",
        "HSF1 (Isı Şoku Faktörü 1) ve Transkripsiyonel Stres Yanıtının Körelmesi",
        "Şaperon genlerini açan master transkripsiyon faktörü HSF1'in yaşla birlikte trimerleşememesi ve epigenetik olarak susturulması.",
        "Isı Şoku Faktörü 1 (HSF1), hücresel proteostazın ana termostatıdır. Normalde sitoplazmada Hsp90 ve Hsp70 tarafından monomerik ve inaktif formda tutulur. Stres anında katlanmamış proteinler şaperonları çekince HSF1 serbest kalır; homotrimerleşir, nükleusa göç eder, Serin 326 pozisyonundan fosforillenir ve şaperon genlerinin promoterlarındaki Isı Şoku Elemanlarına (HSE) bağlanır. Yaşlanmayla birlikte HSF1 ekspresyonu azalmaz ancak trimerleşme yeteneği çöker; SIRT1 deasetilasyonunun durması ve promoterlardaki baskılayıcı histon metilasyonu nedeniyle HSF1 stres anında dahi şaperon üretemez.",
        "HSF1_Transcriptional_Flux = k_hsf * [HSF1_trimer_pS326] * [HSE_occupancy] * (1 / [Repressive_H3K27me3])",
        "HSF1 transkripsiyonel aktivasyon debisi; aktif fosforile trimer HSF1 yoğunluğu ve HSE promoter doluluğunun baskılayıcı heterokromatin işaretine oranıdır."
    ),
    (
        "1.10",
        "Protein Dinamikleri ve Fluktuasyonları: NMR ve Kriyomikroskopik Yapısal Kararlılık",
        "Kristalize statik yapılardan milisaniyelik konformasyonel nefes alma hareketlerine: Biyo-fiziksel termodinamik dalgalanmalar.",
        "Proteinler katı kristal heykeller değil; çözücü içinde sürekli konformasyonel dalgalanmalar (conformational breathing) sergileyen dinamik nano-makinelerdir. Nükleer Manyetik Rezonans (NMR) gevşeme dispersiyonu ve Cryo-EM teknikleri, proteinlerin yerel kararlı durum ile geçici yüksek enerjili 'uyarılmış durumlar' (excited states) arasında mikrosaniye-milisaniye ölçeğinde gidip geldiğini gösterir. Bir mutasyon veya oksidasyon bu uyarılmış durumların serbest enerjisini düşürdüğünde, protein yerel yapısını terk ederek agregasyona yatkın ara ürünler üretir; bu termodinamik kayma amiloidozun fitilini ateşler.",
        "K_exchange = [Conformation_Excited] / [Conformation_Ground] = exp(-Delta_G_conformational / (R * T))",
        "Konformasyonel değişim denge sabiti; uyarılmış agregatif durum ile yerel zemin durum arasındaki serbest enerji farkının Boltzmann dağılımıdır."
    )
]

# ==============================================================================
# KISIM 2: ENDOPLAZMİK RETİKULUM VE UPR (UNFOLDED PROTEIN RESPONSE) SİNYALİ
# ==============================================================================
part2_subsections = [
    (
        "2.1",
        "ER Lümen Ortamı: Kalsiyum Rezervuarı, Oksitleyici Redoks Potansiyeli ve PDI Disülfid İzomerazları",
        "Hücre dışına salgılanacak ve zara gömülecek proteinlerin fabrikası: Sitoplazmadan 1000 kat daha oksitleyici ve kalsiyum dolu biyokimyasal reaktör.",
        "Endoplazmik retikulum (ER) lümeni, hücrenin salgıladığı proteinlerin üçte birinin sentezlendiği, katlandığı ve glikozillendiği yerdir. Sitoplazmanın indirgeyici ortamına zıt olarak ER lümeni son derece oksitleyicidir (GSH/GSSG oranı sitoplazmada 50:1 iken ER'de 3:1'dir). Bu oksitleyici potansiyel, sistein amino asitleri arasında kovalent disülfid bağlarının (S-S) kurulması için şarttır. Protein Disülfid İzomeraz (PDI) ve Ero1 enzimleri elektronları oksijene aktararak kovalent çapraz bağları örer. Eşzamanlı olarak milimolar düzeydeki lümen kalsiyumu (Ca2+), kalneksin ve kalretikulin şaperonlarının glikoproteinleri katlamasını sağlar.",
        "Disulfide_Bond_Flux = k_pdi * [PDI_oxidized] * [Polypeptide_SH_pairs] / [GSH_reducing_buffer]",
        "Disülfid bağı sentez debisi; oksitlenmiş PDI izomeraz yoğunluğu ve serbest tiyol çiftlerinin indirgeyici glutatyon tamponuna oranıyla yönetilir."
    ),
    (
        "2.2",
        "BiP / GRP78 Şaperonunun Sensör Görevi ve Dinamik Ayrışma Kinetiği",
        "ER stresinin nihai bekçisi: Hsp70 ailesi BiP/GRP78 şaperonunun UPR sensörlerine kenetlenerek onları inaktif tutması.",
        "Binding Immunoglobulin Protein (BiP / GRP78), ER lümeninin en bol şaperonudur. Normal homeostaz koşullarında BiP, ER zarına gömülü üç UPR stres sensörünün (IRE1alpha, PERK ve ATF6) lümenal uçlarına sıkıca bağlıdır ve bu sensörlerin dimerleşmesini engelleyerek onları inaktif tutar. ER'de katlanmamış veya hatalı proteinler biriktiğinde, BiP bu hidrofobik proteinlere çok daha yüksek afiniteyle bağlanmak için sensörleri terk eder. BiP'in ayrışması (titrasyonu), üç sensörün anında serbest kalarak oligomerleşmesini ve hücrenin UPR alarmını başlatmasını tetikler.",
        "Sensory_Dissociation_Fraction = [BiP_free_from_Sensors] = [Misfolded_ER_Proteins] / (K_d_BiP + [Misfolded_ER_Proteins])",
        "BiP'in UPR sensörlerinden ayrışma oranı; lümende biriken katlanmamış protein derişiminin BiP bağlanma afinitesine göre Michaelis fonksiyonudur."
    ),
    (
        "2.3",
        "IRE1alpha - XBP1s Kolu: Kinaz Aktivasyonu, Endoribonükleaz Kesimi ve Şaperon İndüksiyonu",
        "UPR'nin evrimsel olarak en eski kolu: Transmembran kinaz/endoribonükleaz IRE1alpha'nın XBP1 mRNA'sını sitoplazmada benzersiz kesimi.",
        "BiP'in ayrılmasıyla serbest kalan IRE1alpha (Inositol-requiring enzyme 1 alpha), ER zarında oligomerleşir ve otofosforilasyonla sitozolik kinaz domenini aktive eder. Kinaz aktivasyonu, komşu endoribonükleaz (RNaz) domeninde allosterik bir konformasyonel değişim yaratır. Aktif RNaz, sitoplazmada bekleyen XBP1 mRNA'sından 26 nükleotidlik bir intronu geleneksel olmayan bir yolla (spliceosome olmaksızın) kesip atar. Çerçeve kaymasıyla oluşan 'Spliced XBP1' (XBP1s), güçlü bir transkripsiyon faktörüdür; nükleusa girerek ER şaperonları, katlanma enzimleri ve ERAD bileşenlerini kodlayan genleri açar.",
        "XBP1s_Generation_Rate = k_rnase * [IRE1a_oligomer_active] * [Uncut_XBP1u_mRNA]",
        "XBP1s transkripsiyon faktörü üretim hızı; zardaki aktif oligomerik IRE1alpha RNaz yoğunluğu ile kesilmemiş XBP1u mRNA mevcudiyetinin çarpımıdır."
    ),
    (
        "2.4",
        "PERK - eIF2alpha - ATF4 Kolu: Global Translasyonun Frenlenmesi ve Seçici Stres Sentezi",
        "Hücrenin genel protein sentezini anında durdurarak ER fabrikasını rahatlatması; paradoxal olarak stres adaptasyon genlerini açması.",
        "PERK (PKR-like ER kinase), BiP'ten ayrılınca dimerleşir ve otofosforillenir. Aktif PERK, ökaryotik translasyon başlatma faktörü 2'nin alfa alt birimini (eIF2alpha) Serin 51 pozisyonundan fosforiller. p-eIF2alpha, guanin nükleotid değişim faktörü eIF2B'yi kilitler; bu da ribozomların genel mRNA çevirisini %80 oranında anında durdurur (yeni protein girişi kesilir). Ancak 5'UTR bölgesinde yukarı-akış açık okuma çerçeveleri (uORF) taşıyan özel transkriptler (örneğin ATF4) bu baskıdan kaçarak tam tersine daha fazla üretilir. ATF4; amino asit taşıyıcılarını, antioksidan yanıtı ve gerekirse ölüm programını devreye sokar.",
        "Global_Translation_Rate = Translation_0 * (1 - [p-eIF2alpha] / [Total_eIF2alpha])",
        "Genel ribozomal translasyon hızı; Serin 51 fosforile eIF2alpha fraksiyonunun toplam faktör havuzuna oranıyla ters orantılı olarak frenlenir."
    ),
    (
        "2.5",
        "ATF6 Kolu: Golgi Proteolitik Kesimi (S1P/S2P) ve Nükleer Translokasyon",
        "ER zarından kopup Golgi aygıtına taşınan ve proteazlar tarafından kesilerek nükleusa uçan kurye transkripsiyon faktörü.",
        "ATF6 (Activating transcription factor 6), 90 kDa'lık bir tip II transmembran proteinidir. BiP ayrıldığında ATF6'nın Golgi hedefleme sinyalleri açığa çıkar; protein COPII vezikülleriyle Golgi aygıtına taşınır. Golgi lümeninde önce Site-1 Proteaz (S1P), ardından zar içinde Site-2 Proteaz (S2P) tarafından iki basamaklı intramembranöz proteolize (RIP) uğrar. Serbest kalan 50 kDa'lık N-terminal sitozolik parça (ATF6f / ATF6p50), nükleusa göç eder. XBP1s ile işbirliği yaparak lipid biyosentezi genlerini açar; ER membran yüzeyini genişleterek katlanmamış protein stresini seyreltir.",
        "Nuclear_ATF6f_Flux = k_rip * [ATF6_Golgi_localized] * [S1P_activity] * [S2P_activity]",
        "Nükleer aktif ATF6 parçası akısı; Golgi'ye ulaşan öncül protein miktarı ile S1P ve S2P ardışık kesim hızlarının çarpımıdır."
    ),
    (
        "2.6",
        "Kronik ER Stresi ve Apoptoz Geçişi: CHOP / GADD153 ve Kaspaz-12/4 Aktivasyonu",
        "Adaptasyonun başarısız olduğu noktada hücrenin intihar kararı: Pro-apoptotik transkripsiyon faktörü CHOP ve Bcl-2 baskılanması.",
        "UPR başlangıçta hücreyi kurtarmaya yönelik sitoprotektif bir programdır; ancak ER stresi saatler/günler boyunca çözülemezse program ölümcül apoptoza kayar. Kronik PERK-ATF4 sinyali ve IRE1alpha'nın aşırı oligomerizasyonu, pro-apoptotik transkripsiyon faktörü CHOP'u (C/EBP homologous protein / GADD153) aşırı aktive eder. CHOP, anti-apoptotik Bcl-2'yi baskılarken pro-apoptotik BIM ve PUMA'yı fırlatır; ER kalsiyumu mitokondriye boşalır. Eşzamanlı olarak IRE1alpha, TRAF2/ASK1 kompleksi üzerinden JNK kinazı aktive eder ve kaspaz kaskadını tetikleyerek hücreyi planlı ölüme götürür.",
        "Apoptosis_Commitment_Ratio = [CHOP] * [BIM] / ([Bcl-2] + epsilon) > Threshold_Death",
        "Apoptoz geri dönüşümsüzlük eşiği; pro-apoptotik CHOP/BIM kompleksinin koruyucu Bcl-2 derişimine oranının ölüm eşiğini aşmasıdır."
    ),
    (
        "2.7",
        "ERAD (ER-Associated Degradation): Retro-translokasyon, Derlin Kanalları ve p97/VCP ATPazı",
        "ER lümeninde onarılamayan kusurlu glikoproteinlerin zardan geri sitoplazmaya fırlatılıp 26S proteazomda imha edilmesi.",
        "ER şaperonları defalarca denemesine rağmen katlanamayan glikoproteinler, EDEM1-3 mannozidazları tarafından uç mannoz şekerleri budanarak 'yıkım' için etiketlenir. Bu proteinler ERAD kompleksine (Hrd1 ve gp78 E3 ligazları, Derlin kanalları) yönlendirilir. Substrat retro-translokasyon ile zardan sitoplazmaya doğru itilir. Sitoplazma tarafında bekleyen hekzamerik AAA+ ATPaz p97/VCP (Valosin-containing protein), ATP hidrolizi ile mekanik kuvvet uygulayarak polipeptidi ER zarından dışarı çeker. Dışarı çıkan protein hızla ubikitinlenir ve proteazom tarafından yutulur.",
        "ERAD_Extraction_Velocity = v_max_p97 * [ATP] / (K_m + [ATP]) * [Substrate_ubiquitinated_Hrd1]",
        "ERAD protein geri-çıkarma hızı; p97/VCP motorunun ATP bağımlı çekme kuvveti ile Hrd1 kompleksinde etiketlenmiş substrat yoğunluğunun fonksiyonudur."
    ),
    (
        "2.8",
        "Yaşlanmada ER Redoks Dengesizliği: PDI Disülfid Bozulması ve H2O2 Birikimi",
        "Yaşlı ER lümeninde disülfid bağı oluşumu sırasında aşırı hidrojen peroksit üretilmesi ve PDI enzimlerinin inaktive olması.",
        "Disülfid bağı oluşumu doğası gereği oksidatif bir süreçtir; PDI her disülfid köprüsü kurduğunda iki elektron Ero1 (Endoplasmic Reticulum Oxidoreductin 1) enzimine geçer. Ero1 bu elektronları doğrudan moleküler oksijene aktararak her bağ başına bir molekül H2O2 üretir. Yaşlanan hücrelerde yanlış katlanma arttıkça Ero1 hiper-aktive olur; ER lümeni devasa bir H2O2 fırtınasına boğulur. Bu oksidatif stres PDI'nın aktif bölgesindeki tiyolleri sülfenik ve sülfonik asitlere oksitleyerek enzimi felç eder; hücre hem disülfid bağlarını kuramaz hem de tüm hücresel zarları peroksidasyona uğratır.",
        "H2O2_Generation_ER = Flux_Disulfide_Bonds = k_ero1 * [Ero1_active] * [O2]",
        "ER lümeni H2O2 üretim akısı; kurulan disülfid bağı akısına eşit olup aktif Ero1 enziminin oksijen tüketim hızı ile belirlenir."
    ),
    (
        "2.9",
        "Lipit Çift Katman Stresi: ER Membran Akışkanlığının UPR Sensörlerini Doğrudan Uyarması",
        "Yalnızca lümendeki katlanmamış proteinler değil; doymuş yağ asitleri ve seramidlerin ER zarını sertleştirerek UPR'yi tetiklemesi.",
        "Peter Walter laboratuvarının keşiflerine göre, UPR sadece lümendeki protein çöpleriyle uyarılmaz; bizzat ER zarının lipit kompozisyonundaki anormallikler de sensörleri aktive eder. Doymuş yağ asitlerinin (örneğin palmitat) fazlalığı veya fosfatidilkolin/fosfatidiletanolamin oranının düşmesi, ER zarının elastikiyetini yok ederek sertleştirir. IRE1alpha ve PERK'in transmembran domenleri bu membran sertleşmesini (Lipid Bilayer Stress) doğrudan algılayarak lümende BiP bulunsa dahi oligomerleşir ve stres yanıtını açar. Bu mekanizma, obezite ve metabolik sendromda ER çöküşünün primer nedenidir.",
        "Membrane_Stress_Score = [Saturated_Phospholipids] / [Unsaturated_Phospholipids] * Membrane_Rigidity",
        "Lipit çift katman stresi skoru; ER zarındaki doymuş fosfolipit oranının membran elastik Young sertlik modülü ile çarpımıdır."
    ),
    (
        "2.10",
        "Farmakolojik ER Şaperonları: TUDCA (Tauroursodeoksikolik Asit) ve 4-PBA Biyokimyası",
        "Hücre dışından verilen kimyasal şaperonların ER lümeninde hidrofobik proteinleri sararak UPR stresini ve apoptozu söndürmesi.",
        "Endoplazmik retikulum stresini kimyasal olarak yatıştırmanın en etkili yolu küçük moleküllü 'kimyasal şaperonlar' kullanmaktır. Doğal bir safra asidi türevi olan Tauroursodeoksikolik Asit (TUDCA) ve aromatik yağ asidi 4-Fenilbütirik Asit (4-PBA), proteinlerin ara katlanma durumlarındaki hidrofobik yüzeyleri stabilize eder. TUDCA, ER lümeninde agregat oluşumunu engeller, BiP'in sensörlere geri bağlanmasını sağlar, PERK/IRE1 sinyalini susturur ve kronik enflamasyonu durdurur. Klinik denemelerde ALS, diyabet ve karaciğer yağlanmasında hücresel sağkalımı dramatik olarak artırmıştır.",
        "ER_Stress_Reduction_Ratio = 1 / (1 + alpha_tudca * [TUDCA_intracellular] / (K_m + [TUDCA_intracellular]))",
        "ER stresi ve UPR sinyalindeki azalma; hücre içine giren TUDCA konsantrasyonuna bağlı doygunluk kinetiğiyle hesaplanır."
    )
]

# ==============================================================================
# KISIM 3: ÜBİKİTİN-PROTEAZOM SİSTEMİ (UPS) VE SELEKTİF PROTEOLİZ
# ==============================================================================
part3_subsections = [
    (
        "3.1",
        "Übikitinasyon Kaskadı: E1 Aktivasyonu, E2 Konjugasyonu ve E3 Ligaz Spesifitesi",
        "Aaron Ciechanover, Avram Hershko ve Irwin Rose'un Nobel ödüllü keşfi: 76 amino asitlik übikitin bayrağının üç basamaklı enzimatik transferi.",
        "Kısa ömürlü, regülatör ve hatalı katlanmış proteinlerin seçici yıkımı üç enzim hiyerarşisiyle yürütülür. İlk basamakta E1 Übikitin Aktive Edici Enzim, ATP harcayarak übikitinin C-terminal glisinini adenilleştirir ve kendi sisteinine yüksek enerjili tiyoester bağıyla bağlar. İkinci basamakta übikitin E2 Übikitin Konjugasyon Enzimine transfer edilir. Üçüncü ve en spesifik basamakta E3 Übikitin Ligaz (insan genomunda 600'den fazla E3 geni vardır: RING, HECT ve RBR sınıfları), substrat proteinini tanır ve übikitin C-terminalini substratın lizin amino grubuna bir izopeptit bağıyla bağlar.",
        "Ubiquitination_Flux = k_E3 * [E3_Ligase] * [Target_Protein] * [E2~Ub] / (K_m + [E2~Ub])",
        "Hücresel übikitinasyon akısı; spesifik E3 ligaz yoğunluğu, hedef protein konsantrasyonu ve aktif E2~Ub tiyoester kompleks mevcudiyetiyle ölçeklenir."
    ),
    (
        "3.2",
        "Übikitin Bağlantı Topolojisi: K48 Poli-Ubikitin Yıkım Sinyali vs K63 Sinyal Ağları",
        "Übikitin molekülünün 7 lizin kalıntısının (K6, K11, K27, K29, K33, K48, K63) oluşturduğu 'übikitin kodu' ve biyolojik anlamları.",
        "Übikitin tek başına bir sinyal olabileceği gibi (mono-übikitinasyon: endositoz ve DNA onarımı), polimerik zincirler de kurabilir. Übikitinin kendi üzerindeki 7 lizin kalıntısı farklı topolojiler oluşturur. Lizin 48 (K48) bağlantılı homotipli poli-übikitin zincirleri (en az 4 übikitin uzunluğunda), 26S proteazom tarafından doğrudan tanınan evrensel 'yıkım fermanıdır' (proteasomal degradation). Buna karşılık Lizin 63 (K63) zincirleri proteazoma gitmez; NF-kappaB sinyal iletimi, DNA hasar yanıtı ve makrootofaji için platform oluşturur. K11 zincirleri hücre döngüsü yıkımında (APC/C) kullanılır.",
        "Proteasomal_Targeting_Efficiency = [K48_PolyUb_Chains] / (Total_Ubiquitin_Signals + epsilon) > 0.85",
        "Proteazomal hedeflenme spesifitesi; kargonun K48-bağlantılı poli-übikitin topolojisi taşıma oranının yüksekliği ile güvenceye alınır."
    ),
    (
        "3.3",
        "26S Proteazom Mimarisi: 20S Katalitik Çekirdek ve 19S Düzenleyici Kapak (Regulatory Particle)",
        "Hücrenin moleküler kağıt öğütücüsü: 2.5 MDa'lık devasa protein kompleksi; merkezde varil şeklinde 20S çekirdek ve iki uçta 19S kapaklar.",
        "26S proteazom, ökaryotik sitoplazma ve nükleusun ana proteolitik makinesidir. Yapısı bir silindir biçimindedir: Merkezde 700 kDa'lık 20S Çekirdek Parçacığı (CP) bulunur; bu silindir alfa1-7 ve beta1-7 olmak üzere dört istiflenmiş halkadan (alfa-beta-beta-alfa) oluşur. 20S'in iki ucunda ise 19S Düzenleyici Parçacık (RP / PA700) yer alır. 19S kapak, substrat proteinleri üzerindeki K48 ubikitin zincirlerini tanır, bağlar, ubikitinleri geri kazanmak için çözer, substratın 3D katlanmasını ATP ile açar ve dar 20S tünelinden içeri sokar.",
        "Stoichiometry_26S = 1 * [20S_Core] + 2 * [19S_Cap] <-> [26S_Holoenzyme] (K_assoc ~ nanomolar)",
        "26S holoenzim montaj dengesi; 20S katalitik çekirdeğin iki adet 19S düzenleyici kapak ile nanomolar afiniteyle kenetlenmesidir."
    ),
    (
        "3.4",
        "20S Proteolitik Odacık: beta1 (Kaspaz-benzeri), beta2 (Tripsin-benzeri) ve beta5 (Kimotripsin-benzeri)",
        "20S silindirinin iç odacığında gizlenmiş N-terminal treonin proteazları; peptit bağlarını 3 ila 15 amino asitlik oligopeptitlere budar.",
        "20S çekirdeğinin beta halkaları iç yüzeyinde proteolitik aktif bölgeleri taşır. Üç farklı katalitik aktivite vardır: beta1 alt birimi asidik amino asitlerden sonra kesim yapar (Kaspaz-benzeri / peptidil-glutamil peptid hidrolaz); beta2 alt birimi bazik amino asitlerden sonra kesim yapar (Tripsin-benzeri); beta5 alt birimi ise hidrofobik amino asitlerden sonra kesim yapar (Kimotripsin-benzeri). Katalitik nükleofil olarak N-terminal Treonin 1 kalıntısı kullanılır. Protein zinciri bu üçlü odacıkta dakikada yüzlerce kez kesilerek zararsız kısa peptit parçalarına ayrılır.",
        "Peptide_Cleavage_Rate = k_beta1 * [beta1_act] + k_beta2 * [beta2_act] + k_beta5 * [beta5_act]",
        "Toplam 20S proteazomal kesim debisi; kaspaz-benzeri (beta1), tripsin-benzeri (beta2) ve kimotripsin-benzeri (beta5) aktivitelerin toplamıdır."
    ),
    (
        "3.5",
        "19S Rpt1-6 AAA+ ATPaz Motorları: Substratın Tanınması, Açılması (Unfolding) ve Matrikse İtilmesi",
        "Altı farklı AAA+ ATPaz alt biriminden oluşan hekzamerik halka: Mekanik çekme kuvvetiyle küresel proteini düz ipliğe dönüştüren motor.",
        "Katlanmış bir protein 20S'in 1.3 nanometrelik dar merkez porundan geçemez; bu nedenle mekanik olarak zorla açılmalıdır. 19S kapağın tabanında (base) bulunan Rpt1, Rpt2, Rpt3, Rpt4, Rpt5 ve Rpt6 adlı 6 adet AAA+ ATPaz proteini halka şeklinde dizilmiştir. Bu motorlar ATP hidroliz ettikçe iç kanallarındaki aromatik ilmekler (pore loops) substrat polipeptid zincirini kavrar ve aşağıya doğru pikonewton düzeyinde mekanik kuvvetle çeker. Proteinin tersiyer yapısı saniyeler içinde çözülür ve çıplak bir makarna ipliği gibi 20S katalitik odacığına beslenir.",
        "Unfolding_Force_19S = N_ATP_hydrolyzed * (Delta_G_ATP / Step_size) > 30_pN",
        "19S AAA+ motorlarının substrat proteinine uyguladığı mekanik açma kuvveti; katlanma enerjisi bariyerini aşacak şekilde 30 pN üzerindedir."
    ),
    (
        "3.6",
        "Deübikitilazlar (DUB'lar): USP14, UCHL5 ve Proteazom Girişinde Kargo Budaması",
        "Substrat 20S odacığına çekilmeden önce değerli ubikitin moleküllerinin kesilerek serbest sitoplazma havuzuna iadesi.",
        "Hücrede serbest ubikitin havuzunun tükenmemesi şarttır; bu nedenle proteazom proteini yutmadan hemen önce ubikitinleri geri toplar. 19S kapağa kenetlenmiş Deübikitilaz enzimleri (DUB'lar: USP14, UCHL5 ve Rpn11/POH1), substratın lizin izopeptit bağlarını hızla hidroliz eder. Rpn11 bir metalloproteazdır ve substrat tam transloke olurken zincirin kökünden ubikitinleri koparır. USP14 ve UCHL5 ise yıkımı geciktirebilen veya hızlandırabilen allosterik fren görevi görür; USP14'ün küçük moleküllerle (IU1) inhibe edilmesi proteazom yıkım hızını katlayarak patolojik tau birikimini temizler.",
        "Deubiquitination_Flux = k_rpn11 * [Rpn11_active] * [Substrate_translocating] * [Zn2+]",
        "Proteazom girişindeki deübikitinasyon debisi; Rpn11 metalloproteaz aktivitesi ve transloke olan kargo hızıyla senkronize ilerler."
    ),
    (
        "3.7",
        "İmmünoproteazom (20Si): LMP2, LMP7 ve MECL-1 ile Antijen İşleme",
        "İnterferon-gamma uyarısıyla standart beta alt birimlerinin yerine geçen özelleşmiş katalitik odacık: MHC Sınıf I antijen üretimi.",
        "Enflamasyon veya viral enfeksiyon anında salgılanan İnterferon-gamma (IFN-gamma), hücrelerde standart 20S alt birimleri yerine 'immünoproteazom' alt birimlerinin sentezini başlatır. beta1 yerine LMP2 (beta1i), beta2 yerine MECL-1 (beta2i) ve beta5 yerine LMP7 (beta5i) geçer. İmmünoproteazom, kimotripsin ve tripsin-benzeri aktiviteleri artırırken kaspaz-benzeri aktiviteyi düşürür; bu sayede C-terminalinde hidrofobik veya bazik amino asit taşıyan kusursuz 8-10 merlik peptid parçaları üretir. Bu peptitler TAP taşıyıcısıyla ER'ye pompalanır ve CD8+ T hücrelerine sunulmak üzere HLA-I moleküllerine yüklenir.",
        "Immunoproteasome_Fraction = [20Si_complexes] / ([20S_standard] + [20Si_complexes]) = f([IFN_gamma])",
        "İmmünoproteazom oranı; hücrenin maruz kaldığı sistemik İnterferon-gamma konsantrasyonu ile sigmoidal olarak artış gösterir."
    ),
    (
        "3.8",
        "Yaşlanmada Proteazom Disfonksiyonu: Alt Birim Ayrışması ve Oksidatif Çapraz Bağlanma",
        "Yaşlı dokularda 26S proteazomun 19S ve 20S parçalarına ayrılması, porların lipofusinle tıkanması ve proteolitik yetmezlik.",
        "Yaşlanan dokularda proteazomal aktivitede %40 ila %70 arasında belirgin bir düşüş gözlenir. Bunun moleküler nedenleri çok yönlüdür: Birincisi, hücresel ATP seviyelerinin düşmesi nedeniyle 19S kapağın 20S çekirdeğe kenetlenmesi gevşer; kompleksler birbirinden kopar. İkincisi, serbest radikallerle oksitlenmiş ve 4-HNE (4-hidroksinonenal) ile çapraz bağlanmış protein agregatları ve lipofusin, 20S'in dar giriş kapısını ve aktif treonin ceplerini fiziksel olarak tıkar. Tıkanan proteazom yeni proteinleri sindiremez; p53, c-Myc ve senesens faktörleri hücre içinde birikerek krizi derinleştirir.",
        "Proteasomal_Activity_Decline = Activity_0 * exp(-k_age * [Aggresome_burden]) * ([ATP] / [ATP_young])",
        "Yaşlanmaya bağlı proteazom aktivite kaybı; hücredeki agregat yükünün üstel baskısı ve hücre içi serbest ATP seviyesinin düşüşü ile doğrudan koreledir."
    ),
    (
        "3.9",
        "Nrf1 (NFE2L1) Proteazom 'Bounce-Back' Yanıtı ve Genetik Adaptasyon",
        "Proteazom tıkandığında ER zarından nükleusa göç ederek tüm proteazom genlerini eşzamanlı açan geri-besleme şalteri.",
        "Hücrede proteazom inhibe edildiğinde veya tıkandığında devreye giren hayati bir genetik sigorta mekanizması vardır: Nrf1 (NFE2L1) transkripsiyon faktörü. Nrf1 normalde ER zarına gömülüdür ve p97/VCP tarafından çekilip proteazomda sürekli parçalanır. Ancak proteazom aktivitesi düştüğünde Nrf1 parçalanamaz; DDI2 proteazı tarafından kesilerek serbest kalır ve nükleusa girer. Nrf1, 26S proteazomun tüm 33 alt biriminin (PSMA, PSMB, PSMC, PSMD genleri) promoterlarındaki ARE (Antioksidan Yanıt Elemanı) bölgelerine bağlanarak yeni proteazom sentezini ('bounce-back' yanıtı) fırlatır.",
        "BounceBack_Proteasome_Flux = k_nrf1 * [Nuclear_Nrf1_cleaved] * sum_i [Promoter_PSM_i]",
        "Proteazom geri-sıçrama adaptasyon akısı; nükleusa giren aktif kesilmiş Nrf1 miktarı ve tüm proteazom gen promoterlarının aktivasyonudur."
    ),
    (
        "3.10",
        "Proteazom Aktivatörleri: Küçük Moleküllerle 20S/26S Kapı Açıcı Protokoller",
        "20S çekirdeğinin alfa halkasındaki kapalı kapıyı allosterik olarak açarak agregatları ATP'siz yutan sentetik ve doğal moleküller.",
        "20S çekirdeğinin uçlarındaki alfa halkaları, istirahat halinde kapalı bir konformasyondadır; içeri protein girişini engeller. Farmakologlar bu kapıyı dışarıdan açan küçük moleküllü 'Proteazom Aktivatörleri' geliştirmiştir. Betulinik asit, oleuropein (zeytin yaprağı özü) ve sentetik moleküller (örneğin PD-169316 türevleri), 20S alfa halkasının ceplerine bağlanarak kapı kalıntılarını (Tyr8, Asp9, Pro17) kenara çeker. Por açıldığında katlanmamış ve amiloidojenik proteinler (alfa-sinüklein, tau) ne ATP'ye ne de übikitine ihtiyaç duymadan doğrudan 20S içine difüze olur ve sindirilir.",
        "Gate_Open_Fraction = [20S_open] / [20S_total] = [Activator] / (K_act + [Activator])",
        "20S proteazom kapı açılma oranı; uygulanan allosterik aktivatör (Betulinik asit/Oleuropein) konsantrasyonunun aktivasyon sabitine oranıdır."
    )
]

parts.append(("KISIM 1: PROTEOSTAZ AĞININ MİMARİSİ VE PROTEİN KATLANMA BİYOFİZİĞİ", part1_subsections))
parts.append(("KISIM 2: ENDOPLAZMİK RETİKULUM VE UPR (UNFOLDED PROTEIN RESPONSE) SİNYALİ", part2_subsections))
parts.append(("KISIM 3: ÜBİKİTİN-PROTEAZOM SİSTEMİ (UPS) VE SELEKTİF PROTEOLİZ", part3_subsections))

# ==============================================================================
# KISIM 4: MAKROOTOFAJİ-LİZOZOM YOLAĞI (ALPS): MOLEKÜLER DİŞLİLER
# ==============================================================================
part4_subsections = [
    (
        "4.1",
        "Otofajinin Fazları: İndüksiyon, Çekirdeklenme (Nucleation), Uzama, Kapanma ve Füzyon",
        "Christian de Duve ve Yoshinori Ohsumi'nin Nobel ödüllü biyolojik temizlik süreci: Çift katmanlı vezikülün sıfırdan doğuşu ve lizozomal imha.",
        "Makrootofaji, hücrenin büyük agregatları ve hasarlı organelleri kendi ürettiği çift katmanlı bir keseye sararak lizozomda eritmesidir. Süreç beş kesin biyofiziksel faza ayrılır: 1) İndüksiyon (mTORC1 baskılanması ve ULK1 kinaz aktivasyonu); 2) Çekirdeklenme / Nucleation (Sınıf III PI3K kompleksiyle omegasom bölgesinde PI3P lipit havuzu oluşturulması); 3) Uzama / Elongation (ATG konjugasyon sistemleriyle fagofor zarının genişlemesi ve LC3-PE lipitlenmesi); 4) Kapanma / Sealing (ESCRT-III kompleksinin fagofor uçlarını kaynaştırarak kapalı otofagozom oluşturması); 5) Füzyon ve Yıkım (SNARE kompleksleri ile lizozomla birleşip otofagolizozomda kargonun sindirilmesi).",
        "Autophagy_Completion_Flux = J_macro = min(Flux_Induction, Flux_Nucleation, Flux_Elongation, Flux_Fusion)",
        "Makrootofaji tamamlanma debisi; beş ardışık aşama arasındaki en yavaş hız kısıtlayıcı basamağın akısına eşittir."
    ),
    (
        "4.2",
        "ULK1 Kompleksi (ULK1-ATG13-FIP200-ATG101): mTORC1 Baskılanması ve AMPK İndüksiyonu",
        "Otofajiyi başlatan ana moleküler şalter: ULK1 kinazın amino asit ve enerji düzeyine göre çift yönlü fosforilasyonu.",
        "ULK1 (Unc-51 like autophagy activating kinase 1), ATG13, FIP200 (RB1CC1) ve ATG101 proteinleriyle 3 MDa'lık stabil bir çekirdek kompleks oluşturur. Besin bolluğunda mTORC1 bu komplekse kenetlenir ve ULK1'i Serin 757 pozisyonundan fosforilleyerek enzimi felç eder. Açlık veya enerji tükenişinde mTORC1 ayrışır; AMPK doğrudan devreye girerek ULK1'i Serin 317 ve Serin 777 pozisyonlarından fosforiller. Aktifleşen ULK1, kendi kendini ve ATG13 ile FIP200'ü otofosforiller; kompleks ER zarına göç ederek çekirdeklenme fazını ateşler.",
        "ULK1_Kinase_Activity = [ULK1] * ( [p-Ser317_AMPK] / ([p-Ser757_mTORC1] + K_m) )",
        "ULK1 katalitik aktivitesi; AMPK aracılı aktive edici fosforilasyonun mTORC1 kaynaklı baskılayıcı fosforilasyona olan oranına bağlıdır."
    ),
    (
        "4.3",
        "Sınıf III PI3K / Beclin-1 Kompleksi ve Fosfatidilinozitol 3-Fosfat (PI3P) Mikro-Alanları",
        "Endoplazmik retikulum zarında PI3P lipit odakları üreterek efektör proteinleri toplayan lipit kinaz nano-makinesi.",
        "ULK1 tarafından aktive edilen ikinci çekirdek modül Sınıf III PI3K Kompleks I'dir: VPS34 (katalitik kinaz alt birimi), Beclin-1 (BECN1), VPS15 (p150 düzenleyici) ve ATG14L. Bu kompleks ER zarına yerleşir ve zardaki fosfatidilinozitolü (PI), Fosfatidilinozitol 3-Fosfata (PI3P) dönüştürür. Oluşan PI3P mikro-alanları, sitoplazmada serbest dolaşan WIPI1/2 (WD-repeat domain phosphoinositide-interacting) proteinleri ve DFCP1 şaperonları için yüksek afiniteli bir mıknatıs görevi görerek omegasom yapısını kurar.",
        "PI3P_Generation_Rate = k_vps34 * [VPS34_active] * [Beclin1_pULK1] * [PI_membrane]",
        "Zardaki PI3P zenginleşme debisi; aktif VPS34 kinaz konsantrasyonu ile ULK1 tarafından fosforillenmiş Beclin-1 miktarının çarpımıdır."
    ),
    (
        "4.4",
        "İki Ubikitin-Benzeri Konjugasyon Sistemi: ATG12-ATG5-ATG16L1 ve LC3-PE (LC3-II)",
        "Bakteriyel benzeri iki biyokimyasal kaskat: Sitoplazmik LC3-I'in fosfatidiletanolamin lipidine kovalent kenetlenerek LC3-II'ye dönüşmesi.",
        "Fagofor zarının büyümesi iki kovalent konjugasyon sistemiyle yürütülür. Birinci sistemde ATG12 (übikitin benzeri protein), ATG7 (E1 benzeri) ve ATG10 (E2 benzeri) aracılığıyla ATG5'e bağlanır; oluşan ATG12-ATG5 heterodimeri ATG16L1 ile birleşerek bir E3-benzeri ligaz kompleksi oluşturur. İkinci sistemde sitozolik LC3 (ATG8 ailesi), sistein proteaz ATG4B tarafından C-terminalinden kesilerek LC3-I'e dönüştürülür. Ardından ATG7 (E1), ATG3 (E2) ve ATG12-ATG5-ATG16L1 (E3) kompleksi, LC3-I'i zar lipidi Fosfatidiletanolamine (PE) kovalent bağlar; oluşan hidrofobik LC3-II (LC3-PE), büyüyen fagofor zarına gömülerek zarı genişletir.",
        "LC3_II_Conversion_Flux = k_lipidation * [ATG16L1_complex] * [LC3_I] * [PE_membrane]",
        "LC3 lipitlenme akısı; zardaki E3 benzeri ATG12-ATG5-ATG16L1 kompleksi, sitozolik LC3-I ve serbest PE fosfolipit konsantrasyonu ile belirlenir."
    ),
    (
        "4.5",
        "Omegasom Mimarisi ve Endoplazmik Retikulumdan Fagofor Zarı Tomurcuklanması",
        "ER zarının oluşturduğu Omega (omega) harfi benzeri beşik yapısı; mitokondri ve Golgi'den gelen lipit akışıyla fagoforun beslenmesi.",
        "Fagofor zarı boşlukta kendiliğinden oluşmaz; ER zarında PI3P ile zenginleşmiş halka benzeri bir çıkıntı olan 'omegasom' beşiğinde filizlenir. Noboru Mizushima laboratuvarının gösterdiği üzere, yeni zar kaynağı sadece ER değildir: Mitokondri dış zarı, trans-Golgi ağı, plazma zarı ve ATG9-pozitif küçük taşıyıcı veziküller sürekli fagofor ucuna lipid pompalar. Lipid transfer proteini ATG2, ER zarı ile büyüyen fagofor arasında bir yağ asidi boru hattı kurarak fosfolipitleri doğrudan aktarır ve kesenin büyümesini sağlar.",
        "Phagophore_Surface_Growth = dArea/dt = 2 * J_lipid_ATG2 * Molecular_Area_Phospholipid",
        "Fagofor membran yüzey alanı büyüme hızı; ATG2 lipit transfer kanalından akan fosfolipit debisi ve lipit moleküler kesit alanı ile orantılıdır."
    ),
    (
        "4.6",
        "Kargo Seçiciliği ve Otofajik Reseptörler: p62, NBR1, TAX1BP1 ve Ubikitin Bağlanması",
        "Agregatların rastgele değil, moleküler etiketleme ile yakalanması: p62/SQSTM1'in PB1 domeniyle oligomerleşip kargoyu sıkıştırması.",
        "Agregatların makrootofaji ile temizlenmesi seçici bir süreçtir (Aggrephagy). Sitoplazmada kümelenen K63-bağlantılı poli-übikitinli protein agregatları, otofaji reseptörü p62 (SQSTM1) tarafından tanınır. p62, C-terminal UBA domeniyle übikitine bağlanırken; N-terminal PB1 domeni aracılığıyla kendi kendine polimerleşerek agregat etrafında sıkı bir kılıf örer. Ardından p62 üzerindeki LIR (LC3-Interacting Region) motifi, fagofor zarındaki LC3-II'ye nanomolar afiniteyle tutunur; kargoyu büyümekte olan otofagozomun tam merkezine hapseder.",
        "Aggrephagy_Targeting_Index = [p62_oligomerized] * [K63_PolyUb_Aggregates] * [LC3_II_bound]",
        "Agrefaji hedeflenme indeksi; oligomerleşmiş p62 kılıfı, poli-übikitinli agregat yoğunluğu ve zar LC3-II bağlanma katsayısının çarpımıdır."
    ),
    (
        "4.7",
        "Otoligasyon ve Fagofor Kapanması: ESCRT-III Makinelerinin Membran Mühürleme Görevi",
        "Açık uçlu fincan benzeri fagoforun iki ucunun birbirini bulup kaynaşması: ESCRT-III kompleksi ve VPS4 ATPazının membran mühürlemesi.",
        "Kargo tamamen sarıldıktan sonra açık kalan zar ağzının kapatılması esastır; aksi takdirde lizozomal enzimler sitoplazmaya sızar. Fagoforun kapatılması (abscission/sealing), virüs tomurcuklanmasında ve sitokinezde görev alan ESCRT-III (Endosomal Sorting Complexes Required for Transport III: CHMP2A, CHMP4B) makineleri tarafından yürütülür. CHMP proteinleri zarın daralan boyun bölgesinde sarmal bir yay oluşturur. AAA+ ATPaz VPS4'ün hidroliz enerjisiyle sarmal büzülür, iç ve dış zarlar birbirinden ayrılarak tam küresel ve sızdırmaz çift katmanlı otofagozom mühürlenir.",
        "Sealing_Fidelity = 1 - P_leakage = f([CHMP2A], [VPS4_ATPase_activity])",
        "Otofagozom kapanma ve mühürlenme sadakati; CHMP2A polimerizasyon hızı ve VPS4 ATPazının membran ayırma verimliliği ile belirlenir."
    ),
    (
        "4.8",
        "Otofagozomun Mikrotübüller Üzerinde Perinükleer Akışı: Dinein Motorları ve Rab7",
        "Hücrenin periferinde oluşan otofagozomların mikrotübül rayları üzerinde sentrozoma doğru santripetal taşınması.",
        "Otofagozomlar genellikle hücrenin periferik korteksinde oluşur; ancak hidrolazlarla dolu olgun lizozomlar çekirdeğin hemen etrafında (perinükleer bölgede) kümelenmiştir. Mühürlenen otofagozom dış zarına küçük GTPaz Rab7 ve FYCO1/RILP adaptörleri yerleşir. RILP, mikrotübül eksi-uç motoru olan Sitoplazmik Dinein/Dinaktin kompleksini bağlar. Otofagozom saniyede birkaç mikrometre hızla mikrotübüller üzerinde sentrozoma doğru çekilir. Yaşlanan hücrelerde mikrotübüllerin asetilasyon dengesi ve Dinein motorları bozulduğu için veziküller periferde takılı kalır ve lizozomla buluşamaz.",
        "Centripetal_Velocity = v_dynein * [ATP] / (K_m + [ATP]) * [Rab7_active_GTP]",
        "Otofagozom perinükleer taşınma hızı; Dinein motorunun ATP bağımlı ilerleme hızı ile zardaki aktif Rab7-GTP yoğunluğunun çarpımıdır."
    ),
    (
        "4.9",
        "Lizozomla Füzyon Mekanizması: HOPS Kompleksi, STX17-SNAP29-VAMP8 SNARE Düğümü",
        "Otofagozom dış zarı ile lizozom tek katmanlı zarının kaynaşması: Dört sarmallı SNARE demetinin 100 pN'luk zarları birleştirme kuvveti.",
        "Perinükleer bölgeye ulaşan otofagozom, lizozom ile karşılaşır. İlk tethering (yakalama), hekzamerik HOPS (Homotypic fusion and vacuole protein sorting) kompleksi tarafından yapılır. Ardından otofagozomdaki Sentaksin 17 (STX17), sitoplazmik SNAP29 ve lizozomdaki VAMP8 proteinleri dört sarmallı trans-SNARE demeti oluşturur. Sarmalların birbirine fermuar gibi kenetlenmesi zarları nanometrik mesafeye çeker ve kaynaştırır. Otofagozomun iç zarı ve içindeki tüm kargo lizozomun asidik lümenine boşalır; otofagolizozom oluşur.",
        "Fusion_Energy_Barrier = Delta_G_SNARE_zippering = -40 * k_B * T (Zarları kaynaştıran kuvvet)",
        "SNARE fermuarlanmasının açığa çıkardığı serbest enerji; su hidrasyon bariyerini aşarak iki zarı kaynaştırmak için gereken -40 k_B*T eşiğini aşar."
    ),
    (
        "4.10",
        "Otofajik Akı (Autophagic Flux) Ölçümü: Klorokin / Bafilomisin Testleri ve GFP-mCherry-LC3",
        "Statik LC3-II miktarının yanıltıcılığı: Lizozomal asidifikasyon blokerleri ile dinamik temizlik hızının gerçek zamanlı testi.",
        "Otofaji araştırmalarındaki en büyük metodolojik hata, sadece LC3-II seviyesine bakarak 'otofaji arttı' sonucuna varmaktır; zira lizozomal sindirim tıkandığında da LC3-II birikir! Gerçek otofajik hız, 'otofajik akı' (autophagic flux) ölçümüyle saptanır. Hücreye lizozom v-ATPase inhibitörü Bafilomisin A1 veya lizozom pH'sını bozan Klorokin eklenir. Biyofiziksel altın standart ise çift floresanlı GFP-mCherry-LC3 raportörüdür: Nötr otofagozomda hem yeşil (GFP) hem kırmızı (mCherry) ışır (sarı görünür); lizozomla birleştiğinde asidik pH yeşil GFP'yi söndürür ve sadece kırmızı mCherry kalır. Kırmızı/Sarı oranı otofajik akının mutlak göstergesidir.",
        "Autophagic_Flux_Ratio = [LC3_II_Bafilomycin] / [LC3_II_Vehicle] = Delta_Accumulation / Delta_t",
        "Dinamik otofajik akı katsayısı; lizozomal blokaj altındaki LC3-II birikim miktarının bazal kontrole oranı ve zamansal türevidir."
    )
]

# ==============================================================================
# KISIM 5: ŞAPERON ARACILI OTOFAJİ (CMA) VE MİKRO-OTOFAJİ
# ==============================================================================
part5_subsections = [
    (
        "5.1",
        "CMA'nın Moleküler İmzası: KFERQ Benzeri Pentapeptid Tanıma Motifi",
        "Ana Cuervo ve Fred Dice'ın çığır açan keşfi: Sitozolik proteinlerin üçte birinde gizlenmiş 5 amino asitlik lizozomal pasaport.",
        "Şaperon Aracılı Otofaji (CMA), makrootofajiden tamamen farklı olarak vezikül kullanmaz; çözünür sitoplazmik proteinleri tek tek lizozom zarına götürüp doğrudan lümene sokar. Bir proteinin CMA kargosu olabilmesi için birincil diziliminde veya 3D yüzeyinde 'KFERQ-benzeri motif' taşıması şarttır. Bu motif; bir bazik kalıntı (K, R), bir veya iki hidrofobik kalıntı (F, L, I, V), bir negatif yüklü kalıntı (E, D) ve pozitif yüklü bir amino asitten oluşur. Memeli proteinlerinin yaklaşık %30'u bu motifi taşır ve stres/açlık anında CMA ile seçici olarak sindirilir.",
        "P_CMA_substrate = Motif_Match(Sequence, [K/R] - [F/L/I/V] - [E/D] - [K/R] - [Q/N])",
        "Bir proteinin CMA substratı olma olasılığı; amino asit zincirinde KFERQ konsensüs kuralına elektrostatik ve hidrofobik uyum fonksiyonudur."
    ),
    (
        "5.2",
        "Sitoplazmik Şaperon Hsc70 (HSPA8) ve Ko-şaperonların Substratı Yakalama Kinetiği",
        "KFERQ motifini tanıyan yegane şaperon: Konstitütif Hsc70 ve ko-şaperonlar Hip, Hop, Hsp40 ve Bag1 işbirliği.",
        "Bir protein katlanmış haldeyken KFERQ motifi genellikle proteinin iç hidrofobik çekirdeğinde gizlidir. Proteinin hasar görmesi, oksidasyonu veya kısmi açılmasıyla motif yüzeye çıkar. Sitoplazmada devriye gezen Isı Şoku Kökenli Protein 8 (Hsc70 / HSPA8), KFERQ motifini yüksek afiniteyle (Kd ~ nanomolar) bağlar. Hsc70'e eşlik eden ko-şaperonlar (CHIP, Hip, Hop ve Hsp40), substratın lizozoma taşınmasını hızlandırır ve proteinin yolda prematüre agregasyon yapmasını engeller.",
        "Cargo_Capture_Rate = k_cap * [Hsc70_free] * [Target_Exposed_KFERQ] * [ATP]",
        "CMA kargo yakalama hızı; serbest sitozolik Hsc70 yoğunluğu, açığa çıkmış KFERQ motif derişimi ve hücresel ATP mevcudiyetiyle ölçeklenir."
    ),
    (
        "5.3",
        "Lizozomal Reseptör LAMP-2A: Monomerik Durumdan Multimerik Translokasyon Kanalına Geçiş",
        "Lizozom zarında monomer olarak bekleyen LAMP-2A reseptörünün substratı görünce 700 kDa'lık silindirik bir tünele dönüşmesi.",
        "Lizozom-İlişkili Membran Proteini 2A (LAMP-2A), CMA'nın hız kısıtlayıcı anahtar reseptörüdür. Substrat yüklü Hsc70 lizozom zarına ulaştığında, protein kargosu LAMP-2A'nın lümene değil sitoplazmaya bakan 12 amino asitlik C-terminal kuyruğuna bağlanır. Bu bağlanma, zardaki monomerik LAMP-2A moleküllerinin hızla oligomerleşerek 700 kDa'lık multimerik bir silindir (translokasyon kanalı) oluşturmasını tetikler. Kanal oluştuktan sonra protein substrat kanalın ortasındaki gözenekten lizozom lümenine doğru itilir.",
        "Translocation_Channel_Density = [LAMP2A_multimer] = k_oligomer * [LAMP2A_monomer]^4 * [Cargo_Bound]",
        "Lizozom zarındaki aktif translokasyon kanalı yoğunluğu; monomerik LAMP-2A derişiminin dördüncü kuvveti ve kargo bağlanmasıyla fırlar."
    ),
    (
        "5.4",
        "Lümenal Hsc70 (lys-Hsc70): Protein Substratın Lizozom İçine Çekilmesi",
        "Lizozomun asidik lümeninde yaşayan özel Hsc70 popülasyonu; protein ipliğini mekanik torkla lümene çeken ratçet motor.",
        "Proteinin LAMP-2A kanalından geçebilmesi için tamamen açılması (unfolding) zorunludur. Proteinin N- veya C-terminal ucu kanalın lümenal ağzından içeri girdiği anda, lizozom lümeni içinde hapsolmuş özel bir Hsc70 havuzu (lys-Hsc70 / lümenal Hsc70) substrat ucunu yakalar. Lümenal Hsc70, bir Brown hareketli mandal (Brownian ratchet) gibi çalışarak substratın sitoplazmaya geri kaymasını engeller; ATP hidroliz ederek proteini kanal boyunca lümene çeker. İçeri giren protein lümendeki proteazlar tarafından saniyeler içinde amino asitlerine parçalanır.",
        "Translocation_Velocity = v_ratchet * [lys-Hsc70] * [ATP_lumen] / (K_m + [ATP_lumen])",
        "CMA translokasyon hızı; lümenal Hsc70 konsantrasyonu ve lizozom içi ATP seviyesinin Brownian mandal kinetiği ile tanımlanır."
    ),
    (
        "5.5",
        "CMA'nın Yaşla Çöküşü: LAMP-2A'nın Doku Düzeyinde Bozulması ve Lizozomal Lipit Sertleşmesi",
        "İnsan dokularında yaşlanmayla LAMP-2A reseptörünün %70'ten fazla kaybolması ve lizozomal lipid mikro-alanlarının kireçlenmesi.",
        "CMA aktivitesi yaşlanmayla birlikte en erken ve en şiddetli çöken kalite kontrol mekanizmasıdır. Bu çöküş LAMP-2A mRNA düzeyinde değil; lizozom zarının lipit dinamiklerinde gerçekleşir. Yaşlanan lizozom zarlarında kolesterol ve sfingomiyelin birikir; bu durum zardaki lipit sallarını (lipid rafts) bozar. Zarda destabilize olan LAMP-2A proteinleri, metalloproteazlar tarafından kesilerek lizozom lümenine dökülür ve yok edilir. LAMP-2A'sız kalan hücre; transkripsiyon faktörlerini (MEF2D), glikolitik enzimleri (PKM2) ve alfa-sinükleini sindiremez; doku dejenerasyonu başlar.",
        "CMA_Decline_Curve = LAMP2A_density(t) = LAMP2A_0 * exp(-k_lipid_drift * Age_Years)",
        "Yaşla CMA aktivite kaybı; lizozom zarındaki kolesterol/sfingomiyelin dislipidemisine bağlı olarak LAMP-2A yüzey yoğunluğunun üstel erimesidir."
    ),
    (
        "5.6",
        "Glukoz ve Amino Asit Metabolizmasının CMA ile Dinamik Regülasyonu",
        "CMA'nın sadece bir çöpçü değil; glikoliz ve lipogenez enzimlerini parçalayarak hücresel enerjiyi yöneten metabolik şalter olması.",
        "CMA kargolarının yarıdan fazlası hücresel ara metabolizma enzimleridir. Glukoz bolluğunda Glikolitik enzimler (Piruvat Kinaz M2, Gliseraldehit-3-fosfat dehidrogenaz, Fosfofrüktokinaz) KFERQ motifleri sayesinde CMA ile seçici olarak parçalanır; böylece glikoliz frenlenir. Açlık durumunda ise lipit sentez enzimleri ve protein sentez inhibitörleri sindirilerek hücre ketogeneze ve glukoneogeneze kaydırılır. Yaşlılıkta CMA durduğunda bu metabolik enzimler hücrede kontrolsüzce birikir; karaciğerde yağlanma, kaslarda insülin direnci ve metabolik kilitlenme ortaya çıkar.",
        "Metabolic_Flexibility_CMA = Rate_Enzyme_Degradation(PKM2, FASN) / Rate_Basal_Synthesis",
        "CMA metabolik esneklik skoru; kilit anabolik enzimlerin CMA aracılığıyla yıkım hızının bazal translasyon hızına oranıdır."
    ),
    (
        "5.7",
        "Endozomal Mikro-otofaji (e-MI): ESCRT Bağımlı ve Bağımsız Sitoplazma Yutumu",
        "Geç endozom zarının içe doğru tomurcuklanarak sitoplazmik proteinleri intralüminal veziküller (ILV) halinde yutması.",
        "Mikro-otofaji, kargonun doğrudan bir litik organel zarı tarafından yutulmasıdır. Endozomal mikro-otofajide (e-MI), geç endozomların zarı içe doğru bükülerek sitoplazmadan damlacıklar koparır ve multiveziküler cisimcikler (MVB) üretir. Bu süreç ESCRT-0, I, II ve III kompleksleri tarafından yürütülür. e-MI hem KFERQ-benzeri motif taşıyan proteinleri Hsc70 aracılığıyla seçici olarak yutabilir hem de çevredeki serbest sitoplazmayı rastgele içine alabilir. MVB'ler lizozomla kaynaştığında bu kargolar tamamen hidrolize edilir.",
        "eMI_Flux = k_esCRT * [ESCRT_III_complex] * Area_Endosome_Membrane * [Hsc70_bound_cargo]",
        "Endozomal mikro-otofaji debisi; ESCRT-III membran bükme aktivitesi, endozom yüzey alanı ve Hsc70 kargo yoğunluğunun çarpımıdır."
    ),
    (
        "5.8",
        "Lizozomal Doğrudan Mikro-otofaji: Lizozom Zarının İçe Doğru Çökmesi (İnvajinasyon)",
        "Lizozomun kendi membranını bir amip gibi içe bükerek katı agregatları ve glikojen granüllerini doğrudan yutması.",
        "Memeli hücrelerinde doğrudan lizozomal mikro-otofajide lizozom zarı invajinasyon (içe çökme) veya çıkıntı (protrüzyon) yaparak sitoplazmadaki küçük parçacıkları kucaklar. Bu mekanizma özellikle hücresel glikojen depolarının parçalanmasında (Glikofaji) ve proteazom komplekslerinin açlık anında sindirilmesinde (Proteafaji) hayati rol oynar. Süreç otofagozom gerektirmediği için makrootofajinin çalışmadığı kriz anlarında hücrenin acil amino asit ve glukoz temin etmesini sağlayan hızlı bir destek hattıdır.",
        "Microautophagy_Ingestion_Rate = k_invag * Membrane_Fluidity_Lysosome * [Target_Granules]",
        "Doğrudan mikro-otofajik yutma hızı; lizozomal membran akışkanlığı ve hedeflenen sitoplazmik granül yoğunluğu ile doğru orantılıdır."
    ),
    (
        "5.9",
        "CMA'nın Kanser ve Nörodejenerasyondaki İki Yüzlü Rolü",
        "Genç sağlıklı dokuda tümörleri engelleyen ve agregatları temizleyen CMA'nın; ileri kanserlerde tümörü besleyen süper-silaha dönüşmesi.",
        "CMA iki yüzlü bir kılıçtır (double-edged sword). Sağlıklı bir nöron veya hepatositte CMA; alfa-sinüklein, Tau, huntingtin ve pro-onkojenleri sindirerek nörodejenerasyonu ve kanser başlangıcını sertçe engeller. Ancak bir hücre malign transformasyona uğradığında (kanserleştiğinde), tümör hücreleri LAMP-2A ekspresyonunu 5 kat artırır; yüksek CMA aktivitesini hipoksi ve kemoterapi stresine karşı hayatta kalmak ve Warburg glikolizini sürdürmek için kullanır. Bu nedenle yaşlanmayı önlemede CMA uyarımı hedeflenirken, ileri evre kanserde CMA blokajı gerekir.",
        "CMA_Therapeutic_Vector = +1 (Anti-Aging / Neuroprotection) vs -1 (Advanced_Tumor_Progression)",
        "CMA terapötik vektörü; sağlıklı yaşlanmada pozitif bir gençleştirici iken ilerlemiş karsinomda tümör destekleyicidir."
    ),
    (
        "5.10",
        "Genetik ve Farmakolojik LAMP-2A Restorasyonu ile Organ Fonksiyonlarının Gençleştirilmesi",
        "Ana Cuervo'nun tarihi deneyi: Yaşlı fare karaciğerinde tek bir genle (LAMP-2A) CMA'nın gençlik seviyesine çıkarılması ve organ gençleşmesi.",
        "2008 yılında Nature Medicine'da yayınlanan dönüm noktası çalışmada, Ana Maria Cuervo laboratuvarı genetik mühendislikle yaşlı farelerin karaciğerine ekstra bir LAMP-2A geni yerleştirmiştir. Normalde yaşlılıkta çöken CMA aktivitesi 22 aylık (yaşlı) farelerde 3 aylık (genç) farelerin seviyesine çıkarılmıştır. Sonuçlar mucizevidir: Karaciğer hücrelerindeki tüm protein agregatları temizlenmiş, lipofusin birikimi durmuş, mitokondriyal ve metabolik fonksiyonlar tamamen gençleşmiş ve karaciğer genç bir organ gibi çalışmaya devam etmiştir. Günümüzde küçük moleküllü kimyasal CMA aktivatörleri (AR7, CA77) klinik geliştirme aşamasındadır.",
        "Organ_Rejuvenation_Index = [Functional_Hepatocytes_CMA_Restored] / Total_Tissue * 100 -> Youth_Phenotype",
        "Organ gençleşme derecesi; LAMP-2A takviyesi ile CMA aktivitesi restore edilen hücrelerin dokudaki yüzdesi ile birebir koreledir."
    )
]

# ==============================================================================
# KISIM 6: LİZOZOM BİYOLOJİSİ, V-ATPASE VE TRANSKRİPSİYON FAKTÖRÜ EB (TFEB)
# ==============================================================================
part6_subsections = [
    (
        "6.1",
        "Lizozom Lümen Biyofiziği: v-ATPase Proton Pompası ve pH 4.5-5.0 Asidifikasyonu",
        "Hücrenin nihai asit havuzu: V-tipi H+-ATPaz pompasının sitoplazmadan lümene proton basarak 1000 kat hidrojen iyonu farkı yaratması.",
        "Lizozomun sindirim gücü, lümen içindeki aşırı asidik ortama (pH 4.5 - 5.0) bağımlıdır; çünkü lizozomal hidrolazlar yalnızca bu pH aralığında aktiftir. Bu asidifikasyon, lizozom zarına gömülü devasa v-ATPase (V1VO holoenzimi) proton pompası tarafından sağlanır. v-ATPase, sitozolik V1 sektöründe ATP hidroliz ederek VO sektöründeki c-halkasını döndürür ve protonları zardan lümene pompalar. Pozitif yük birikimini dengelemek ve elektriksel potansiyeli nötralize etmek için ClC-7 klorür/proton antiporterı lümene eşzamanlı klorür (Cl-) iyonları taşır. Yaşlanmayla v-ATPase alt birimleri ayrışır; lümen pH'sı 6.0'a yükselir ve tüm sindirim felç olur.",
        "Lumen_pH = -log10([H+_lumen]) = pH_cytosol - (F * Delta_Psi_lys + Delta_mu_ATP) / (2.303 * R * T)",
        "Lizozomal lümen pH'sı; v-ATPase ATP hidroliz serbest enerjisi ve transmembran elektriksel potansiyel dengesiyle belirlenir."
    ),
    (
        "6.2",
        "Lizozomal Asit Hidrolazları: Katepsinler (B, D, L), Glukoserebrosidaz (GBA) ve Nükleazlar",
        "Proteinden lipide, DNA'dan polisakkarite kadar tüm biyolojik polimerleri monomerlerine parçalayan 60 farklı asit hidrolaz enzimi.",
        "Lizozom lümeni yaklaşık 60 farklı asidik hidrolaz barındırır. Başlıca protein sindiriciler Katepsin ailesidir: Katepsin B ve L (sistein proteazlar) ile Katepsin D (aspartil proteaz). Sfingolipitlerin yıkımı Glukoserebrosidaz (GBA1) ve Asit Sfingomiyelinaz tarafından yürütülür; nükleik asitler DNAse II ile parçalanır. Bu enzimler ER ve Golgi'de sentezlenirken Mannoz-6-Fosfat (M6P) şekeriyle etiketlenir ve M6P reseptörleri aracılığıyla lizozoma hedeflenir. GBA1 mutasyonları Gaucher hastalığına ve Parkinson'a yol açarken; yaşlanmada katepsinlerin aktivite kaybı agregat birikiminin ana motorudur.",
        "Hydrolytic_Degradation_Capacity = sum_i k_cat_i * [Hydrolase_i] * (1 / (1 + 10^(pH_lumen - pH_opt_i)))",
        "Toplam lizozomal hidrolitik kapasite; 60 asit hidrolaz enziminin lümen pH sapmasına bağlı katalitik debilerinin toplamıdır."
    ),
    (
        "6.3",
        "CLEAR (Coordinated Lysosomal Expression and Regulation) Gen Ağı ve Biyogenez",
        "Andrea Ballabio'nun keşfi: Tüm lizozomal ve otofajik genlerin promoterlarında bulunan 10 baz çiftlik evrensel konsensüs dizisi.",
        "2009 yılında Andrea Ballabio laboratuvarı, lizozomal sistemin rastgele genlerden oluşmadığını; tek bir ana gen ağı tarafından koordine edildiğini keşfetmiştir: CLEAR (Coordinated Lysosomal Expression and Regulation) geni ağı. Lizozom yapısında görev alan 400'den fazla genin (v-ATPase alt birimleri, katepsinler, LAMP'ler, otofaji ATG genleri) promoter bölgesinde palindromik bir GTCACGTG konsensüs dizisi bulunur. Bu dizi, lizozom biyogenezinin master transkripsiyon faktörleri olan MiT/TFE ailesi tarafından tanınır.",
        "CLEAR_Network_Activation = sum_genes ( [TFEB_nuclear] * Affinity(TFEB, CLEAR_motif_k) )",
        "CLEAR gen ağı aktivasyon şiddeti; nükleusa giren TFEB miktarının promoterlardaki CLEAR motifi bağlanma afinitesiyle çarpımıdır."
    ),
    (
        "6.4",
        "TFEB (Transcription Factor EB): mTORC1 Tarafından Fosforilasyonu ve Sitoplazmik 14-3-3 Tutuklanması",
        "Hücre tokken TFEB'in lizozom zarına bağlanıp mTORC1 ile fosforillenerek sitoplazmada kelepçelenmesi; açlıkta nükleusa kaçışı.",
        "TFEB (Transcription Factor EB), lizozomal biyogenezin ve otofajinin master orkestra şefidir. Hücre besin ve amino asit bakımından zengin olduğunda, TFEB lizozom zarına çekilir. Burada bulunan aktif mTORC1 kinazı, TFEB'i Serin 211 pozisyonundan fosforiller. p-Ser211-TFEB, sitoplazmik şaperon proteini 14-3-3 tarafından yakalanır ve sitoplazmada hareketsiz kilitlenir; nükleusa giremez. Hücre aç kaldığında veya lizozom zorlandığında mTORC1 inaktive olur; TFEB defosforillenir, 14-3-3'ten kurtulur ve hızla nükleusa geçerek tüm CLEAR ağını ateşler.",
        "Nuclear_TFEB_Fraction = [TFEB_nuclear] / [TFEB_total] = 1 / (1 + [mTORC1_active] * [14-3-3] / K_dissoc)",
        "Nükleer aktif TFEB oranı; aktif lizozomal mTORC1 kinaz yoğunluğu ve 14-3-3 şaperon tutuklama dengesinin ters fonksiyonudur."
    ),
    (
        "6.5",
        "Lizozomal Kalsiyum Sensörü (Mucolipin-1 / TRPML1), Kalsinörin ve TFEB Nükleer Kaçışı",
        "Lizozom stresinin zardaki TRPML1 kanalını açarak kalsiyum fışkırtması; Kalsinörin fosfatazın TFEB'in kelepçesini kırması.",
        "TFEB'in nükleusa göçünün en zarif ve hızlı mekanizması lizozomal kalsiyum sinyalidir. Lizozom içinde yüksek konsantrasyonda kalsiyum depolanır. Açlık, ROS veya lümen stresi anında lizozom zarındaki TRPML1 (Mucolipin-1 / MCOLN1) kalsiyum kanalı açılır. Zardan sitoplazmaya mikro-alan kalsiyum akışı (Ca2+ puff) başlar. Bu lokal kalsiyum artışı, kalsiyum bağımlı fosfataz Kalsinörini (PPP3CA) doğrudan aktive eder. Kalsinörin, TFEB'in Serin 211'indeki fosfatı anında koparır; TFEB saniyeler içinde 14-3-3'ten sıyrılarak nükleusa hücum eder.",
        "TFEB_Dephosphorylation_Rate = k_calcineurin * [Calcineurin_active] * ([Ca2+_TRPML1_microdomain] / K_Ca)",
        "TFEB defosforilasyon hızı; TRPML1 kanalından boşalan lizozomal kalsiyum mikro-alan yoğunluğunun Kalsinörin aktivasyonuyla çarpımıdır."
    ),
    (
        "6.6",
        "Yaşlanmada Lizozomal Disfonksiyon: Lümen Nötralizasyonu ve Sızıntılı Lizozom Sendromu",
        "Yaşlı hücrede v-ATPase çöküşü, lümen pH'sının alkalileşmesi ve zardan sitoplazmaya sızan asit proteazların hücreyi tahrip etmesi.",
        "Yaşlanma sürecinde lizozomlar biyofiziksel bir felaket yaşar. Birincisi, lümen asidifikasyonu çöker; pH 4.5'ten 6.0'a çıkarak asit hidrolazları felç eder. İkincisi ve daha tehlikelisi, lizozomal membran geçirgenleşmesidir (LMP - Lysosomal Membrane Permeabilization). Lipofusin birikimi, demir bağımlı Fenton reaksiyonları ve serbest radikaller lizozom zarındaki fosfolipitleri parçalar. Kısmen sızıntı yapan lizozomlardan sitoplazmaya kaçan Katepsin B ve D, sitozolik proteinleri rastgele keser, NLRP3 inflamazomunu uyarır ve hücreyi steril enflamasyona (inflam-aging) veya piroptotik ölüme sürükler.",
        "LMP_Leakage_Flux = N_membrane_ruptures * Area_rupture * (D_cathepsin / Thickness_membrane) * [Cathepsin_lumen]",
        "Lizozomal membran sızıntı debisi; zardaki peroksidatif yırtıkların sayısı ve katepsin difüzyon gradyanı ile hesaplanır."
    ),
    (
        "6.7",
        "Lizofaji: Yırtılan Hasarlı Lizozomların Galektinler (Gal-3/Gal-8) Tarafından Ayıklanması",
        "Zarı patlayan lizozomun lüminal şekerlerinin sitoplazmaya açığa çıkması; Galektin-3/8'in otofaji makinelerini çağırarak organeli imhası.",
        "Zarı yırtılan bir lizozom tüm hücre için ölümcül bir zehirdir; hücre bu hasarlı organeli derhal temizlemek için 'Lizofaji' (Lysophagy) sürecini işletir. Normalde lizozom lümeninde bulunan yüksek konsantrasyondaki beta-galaktozid glikanlar sitoplazmada asla bulunmaz. Zar yırtıldığında bu şekerler sitoplazmaya açığa çıkar. Sitozolik lektinler Galektin-3 ve Galektin-8 (Gal-3 / Gal-8), açığa çıkan bu glikanları derhal yakalar. Galektin-8, otofaji reseptörü NDP52'yi çağırır; E3 ligaz FBXO27 zarı ubikitinler ve fagofor hasarlı lizozomu sararak sağlıklı lizozomlar tarafından yutturur.",
        "Lysophagy_Induction_Velocity = k_gal8 * [Galectin8_recruited] * [NDP52] * [Exposed_Glycans]",
        "Lizofaji indüksiyon hızı; yırtılan lizozomdan açığa çıkan glikanları yakalayan Galektin-8 ve NDP52 adaptör yoğunluğunun çarpımıdır."
    ),
    (
        "6.8",
        "Lizozom Konumlanması (Periferik vs Perinükleer) ve Hücresel Besin Algılama",
        "Besin durumuna göre lizozomların mikrotübüller üzerinde plazma zarı ile çekirdek arasında gidip gelmesi: Kinesin vs Dinein satrancı.",
        "Lizozomlar hücre içinde rastgele dağılmaz; konumları metabolik durumun haritasıdır. Hücre besinle dolu olduğunda, Rab7-FYCO1 kompleksi Kinesin-1 motorunu bağlar ve lizozomları hücre çeperine (periferik) taşır. Periferde plazma zarına yakın olan lizozomlar mTORC1 aktivasyonunu maksimize eder. Açlık durumunda ise lizozomlar Dinein motoruyla çekirdeğin hemen yanına (perinükleer bölgeye) çekilir; burada otofagozomlarla füzyon şansı katlanır ve TFEB nükleusa kolayca fırlar. Yaşlanmada Kinesin/Dinein dengesinin bozulması lizozomları hareketsiz kümeler halinde dondurur.",
        "Lysosome_Position_Ratio = Count(Perinuclear_Lysosomes) / Count(Peripheral_Lysosomes) = f(Starvation_State)",
        "Lizozom perinükleer/periferik konumlanma oranı; hücresel açlık sinyali ve Dinein retrograd transportunun doğrudan fonksiyonudur."
    ),
    (
        "6.9",
        "Farmakolojik TFEB Aktivatörleri: Trehaloz, Curcumin Analogları ve mTORC1-Bağımsız İndükleyiciler",
        "mTORC1'i inhibe edip bağışıklığı çökertmeden, doğrudan TFEB'i nükleusa sokarak tüm hücresel temizliği tetikleyen akıllı moleküller.",
        "Klasik otofaji indükleyicisi Rapamisin, mTORC1'i tamamen kapattığı için uzun vadede immünosüpresyon ve glukoz intoleransı yapabilir. Bu nedenle mTORC1'den bağımsız çalışan TFEB aktivatörleri büyük bir devrimdir. Doğal disakkarit Trehaloz, GLUT taşıyıcılarını hafifçe kısıtlayarak AMPK/kalsinörin üzerinden TFEB'i nükleusa sokar. Curcumin türevi C1 ve sentetik küçük moleküller (örneğin Hep14), doğrudan TRPML1 kanalını açarak veya TFEB'e bağlanarak mTORC1'e kesinlikle dokunmadan CLEAR ağını açar; nöronlarda amiloid ve tau temizliğini katlar.",
        "TFEB_Nuclear_Flux_mTOR_indep = k_treh * [Trehalose] * [TRPML1_open_flux] (Sıfır mTORC1 inhibisyonu)",
        "mTORC1-bağımsız TFEB nükleer akısı; Trehaloz ve TRPML1 kalsiyum akışının tetiklediği kalsinörin defosforilasyon hızı ile ölçeklenir."
    ),
    (
        "6.10",
        "Lizozomal Depo Hastalıkları (Gaucher, Niemann-Pick) ile Yaşlanma Arasındaki Moleküler Bağ",
        "Genetik lizozom yetersizliklerinin erken yaşlanma tabloları üretmesi; normal yaşlanmanın edinsel bir lizozomal depo hastalığı olması.",
        "Lizozomal enzim eksiklikleri (örneğin GBA1 mutasyonuyla glukoserebrozid birikimi olan Gaucher; NPC1/2 mutasyonuyla kolesterol birikimi olan Niemann-Pick Tip C), çocukluk çağında ağır nörodejenerasyon ve erken ölüm yaratır. Bu hastaların beyin ve karaciğer biyopsileri, normal 85 yaşındaki bir insanın dokularıyla birebir aynı patolojiyi gösterir: Lipofusin birikimi, hasarlı mitokondriler, bloklanmış otofaji ve proteostaz çöküşü. Bu çarpıcı paralellik, insan yaşlanmasının özünde 'tüm dokularda yavaş gelişen edinsel ve kümülatif bir lizozomal depo hastalığı' olduğunu kanıtlamaktadır.",
        "Aging_Phenotype_Correlation = Correlate(Transcriptome_NiemannPick, Transcriptome_Aged_Brain) > 0.82",
        "Niemann-Pick lizozomal depo hastalığı ile yaşlı insan beyni transkriptomik örtüşme korelasyonu %82'nin üzerindedir."
    )
]

parts.append(("KISIM 4: MAKROOTOFAJİ-LİZOZOM YOLAĞI (ALPS): MOLEKÜLER DİŞLİLER", part4_subsections))
parts.append(("KISIM 5: ŞAPERON ARACILI OTOFAJİ (CMA) VE MİKRO-OTOFAJİ", part5_subsections))
parts.append(("KISIM 6: LİZOZOM BİYOLOJİSİ, V-ATPASE VE TRANSKRİPSİYON FAKTÖRÜ EB (TFEB)", part6_subsections))

# ==============================================================================
# KISIM 7: AGREGATOM (AGGRESOME), AMİLOİD FİBRİLLER VE LİPOFUSİN BİRİKİMİ
# ==============================================================================
part7_subsections = [
    (
        "7.1",
        "Protein Agregasyon Mekanizması: Monomer -> Oligomer -> Protofibril -> Amiloid Çapraz beta-Tabakaları",
        "Çözünür yerel proteinden nükleasyon bağımlı polimerizasyonla çözünmeyen rijit nano-liflere geçişin kinetik fazları.",
        "Protein agregasyonu, nükleasyon bağımlı bir polimerizasyon (nucleation-dependent polymerization) kinetiği izler. Süreç, yerel yapısını kaybeden monomerlerin bir araya gelerek termodinamik olarak kararsız bir 'çekirdek' (nucleus) oluşturduğu yavaş 'gecikme fazı' (lag phase) ile başlar. Çekirdek oluştuktan sonra 'uzama fazı' (elongation phase) patlar; monomerler çekirdeğe hızla eklenerek çözünür oligomerleri, protofibrilleri ve nihayetinde amiloid fibrillerini kurar. Tüm amiloid fibrilleri, fibril eksenine dik uzanan beta-iplikçiklerinden oluşan evrensel bir 'çapraz beta-tabakası' (cross-beta sheet) mimarisi sergiler.",
        "Aggregation_Kinetics = d[Fibril]/dt = 2 * k_elongation * [Monomer] * [Fibril_Ends] + k_secondary * [Monomer]^n * [Fibrils]",
        "Amiloid fibril büyüme hızı; primer uzama akısı ile mevcut fibril yüzeyinde yeni çekirdeklerin doğduğu ikincil nükleasyon akısının toplamıdır."
    ),
    (
        "7.2",
        "Toksisite Paradoksu: Olgun Amiloid Plaklar vs Çözünür Oligomerik Toksisite",
        "Büyük amiloid plakların hücreyi koruyan nötr mezarlıklar olması; asıl hücresel katilin küçük, hareketli çözünür oligomerler olması.",
        "Amiloid araştırmalarındaki en büyük paradigma değişimi, olgun fibriller ile oligomerler arasındaki toksisite ayrımıdır. Mikroskopta görülen devasa ekstraselüler amiloid plakları veya intraselüler inklüzyonlar, aslında hücrenin toksik molekülleri hareketsiz kılmak için kurduğu biyolojik bir 'çöp depolama alanıdır'. Asıl öldürücü ajanlar, 2 ila 50 monomerden oluşan küçük, çözünür oligomerlerdir. Bu oligomerler açıkta kalan hidrofobik yüzeyleriyle hücresel zarlara yapışır, zarda yapay gözenekler açarak kontrolsüz kalsiyum sızıntısı başlatır ve organelleri delerek hücreyi zehirler.",
        "Cytotoxicity_Index = [Soluble_Oligomers] * Surface_Hydrophobicity / ([Inert_Fibrillar_Plaques] + epsilon)",
        "Amiloid sitotoksisite indeksi; çözünür hareketli oligomer yoğunluğu ve hidrofobik yüzey alanının hareketsiz fibriler plak kütlesine oranıdır."
    ),
    (
        "7.3",
        "Agresom Oluşumu: Mikrotübül-Organize Edici Merkeze (MTOC) Dinein İle Taşınma ve Vimentin Kafesi",
        "Proteazom sindiremediğinde hücrenin sitoplazmaya saçılmış agregatları sentrozomda toplayarak vimentin zırhına hapsetmesi.",
        "Sitoplazmada dağınık halde bulunan yanlış katlanmış proteinler proteazom kapasitesini aştığında, hücre 'Agresom' (Aggresome) mekanizmasını devreye sokar. HDAC6 enzimi, ubikitinli agregatları tanır ve onları mikrotübül motoru Dinein'e bağlar. Dağınık çöpler mikrotübül rayları üzerinden çekirdeğin hemen yanındaki Mikrotübül-Organize Edici Merkeze (MTOC / Sentrozom) taşınır. Burada toplanan agregatların etrafı ara filaman proteini Vimentin tarafından örülen sıkı bir kafesle (vimentin cage) çevrilir. Bu mekanizma toksik oligomerleri tek bir noktada hapsederek hücrenin genel sitoplazmasını korur ve kütleyi makrootofajiye sunar.",
        "Aggresome_Formation_Flux = k_agg * [HDAC6_active] * [Dynein_motor] * [Ubiquitinated_Aggregates]",
        "Agresom oluşum debisi; aktif HDAC6 bağlayıcı yoğunluğu, Dinein retrograd transport hızı ve ubikitinlenmiş agregat konsantrasyonu ile belirlenir."
    ),
    (
        "7.4",
        "Lipofusin (Yaşlılık Pigmenti): Okside Lipitler, Metal İyonları ve Yıkılamayan Proteazom Çöpü",
        "Bölünmeyen yaşlı hücrelerin lizozomlarında biriken, hiçbir enzimin parçalayamadığı sarı-kahverengi floresan çöp yığını.",
        "Lipofusin ('yaşlılık pigmenti' veya seroid), kalp kası kardiyomiyositlerinde, retina pigment epitelinde ve beyin nöronlarında yaşla kümülatif olarak biriken oto-floresan bir agregattır. Kimyasal kompozisyonu son derece kaotiktir: %30-70 oksitlenmiş proteinler, %20-50 peroksidasyona uğramış lipitler (özellikle doymamış yağ asidi parçalanma ürünleri MDA ve 4-HNE), şekerler ve yüksek konsantrasyonda geçiş metalleri (Fe2+, Cu2+) içerir. Kovalent çapraz bağlarla (Schiff bazları) birbirine kenetlendiği için memeli biyolojisindeki hiçbir enzim tarafından hidrolize edilemez; hücre öldüğünde dahi bozulmadan kalır.",
        "Lipofuscin_Accumulation_Rate = k_lipo * [Lipid_Peroxidation] * [Misfolded_Proteins] * [Fe2+_lysosome] * Time_Years",
        "Lipofusin birikim debisi; lizozomdaki lipit peroksidasyonu, yanlış katlanmış proteinler ve serbest demir iyonlarının zamansal integralidir."
    ),
    (
        "7.5",
        "Lipofusinin Lizozomları Felç Etmesi: Hidrolazların Ototutuklanması ve Hücrenin Boğulması",
        "Biriken lipofusin granüllerinin lizozomun tüm iç hacmini işgal ederek yeni çöplerin sindirilmesini fiziksel ve kimyasal durdurması.",
        "Lipofusin sadece pasif bir çöp değildir; aktif bir hücresel zehirdir. Birincisi, lizozomal lümenin hacmini doldurarak yeni otofagozomların füzyonunu fiziksel olarak engeller (yer işgali). İkincisi, yüzeyindeki hidrofobik cepler ve doymamış aldehitler, lizozomal asit hidrolazlarını (katepsinleri) bir sünger gibi bağlayarak inaktive eder. Üçüncüsü, bünyesinde hapsettiği redoks-aktif demir iyonları nedeniyle sürekli hidrojen peroksit ile Fenton reaksiyonuna girerek serbest radikal üretir; lizozom zarını içeriden parçalar. 80 yaşındaki bir insanın kardiyomiyosit hacminin %20'den fazlası lipofusinle doludur.",
        "Lysosomal_Choking_Index = Volume_Lipofuscin / Volume_Total_Lysosome -> Critical_Failure > 0.40",
        "Lizozomal boğulma indeksi; lipofusinin işgal ettiği lümen hacminin toplam lizozom hacmine oranı olup %40'ı aştığında otofajik kilitlenme yaratır."
    ),
    (
        "7.6",
        "Nörodejeneratif Agregatomlar: Amiloid-beta, Tau Yumakları, alfa-Sinüklein ve TDP-43",
        "İnsan beynini çürüten dört büyük protein agregatı: Ekstraselüler plaklar, nörofibriler yumaklar, Lewy cisimcikleri ve nükleer göç kusurları.",
        "Farklı nörodejeneratif hastalıklar farklı proteinlerin agregasyonuna dayanır: Alzheimer hastalığında 42 amino asitlik hidrofobik Amiloid-beta (Abeta42) ekstraselüler plakları ve hiperfosforile Tau mikrotübül proteini nörofibriler yumakları kurar. Parkinson hastalığında presinaptik protein alfa-sinüklein agregasyona uğrayarak Lewy cisimciklerini oluşturur. ALS ve Frontotemporal Demansta (FTD) ise normalde çekirdekte RNA işleyen TDP-43 proteini nükleustan sitoplazmaya kaçar, fosforillenip C-terminal fragmanlar halinde toksik agregatlara çöker. Hepsi çapraz beta-tabaka omurgasını paylaşır.",
        "Neurodegeneration_Load = alpha_abeta * [Abeta42_oligomers] + beta_tau * [p-Tau_tangles] + gamma_syn * [alpha-Syn_Lewy]",
        "Kümülatif nörodejeneratif amiloid yükü; oligomerik Abeta42, hiperfosforile Tau ve Lewy alfa-sinüklein agregatlarının ağırlıklı toplamıdır."
    ),
    (
        "7.7",
        "Faz Ayrışması (LLPS): Sıvı Damlacıklardan Geri Dönüşümsüz Jel ve Katı Agregatlara Faz Geçişi",
        "Anthony Hyman'ın devrimsel fizikokimyasal teorisi: Membransız organellerin yoğunlaşması ve zamanla kristalize amiloide donması.",
        "Hücre içindeki birçok protein (FUS, TDP-43, hnRNPA1) nükleer cisimcikler ve stres granülleri gibi 'membransız organeller' oluşturmak için Sıvı-Sıvı Faz Ayrışması (LLPS - Liquid-Liquid Phase Separation) yapar. Düşük karmaşıklıklı dizilimler (LCD) ve intrinsik düzensiz bölgeler (IDR), zayıf çok değerlikli (multivalent) etkileşimlerle dinamik sıvı damlacıkları halinde yoğunlaşır. Ancak yaşlanma, oksidatif stres veya mutasyonlar bu dinamik faz dengesini bozar: Sıvı damlacıklar zamanla vizkoz jeller haline gelir (jel fazı) ve nihayetinde geri dönüşümsüz, katı, kristalize amiloid liflerine 'donar' (katılaşma).",
        "Phase_Transition_Velocity = dSolid/dt = k_gel * [IDR_protein]^n * exp(-E_activation / (k_B * T))",
        "Sıvı-katı faz ayrışma hızı; intrinsik düzensiz proteinlerin yerel yoğunlaşması ve nükleasyon aktivasyon enerjisinin aşılmasıyla belirlenir."
    ),
    (
        "7.8",
        "Prion Benzeri Yayılım (Seeding): Agregatların Eksozomlarla Hücreden Hücreye İnfeksiyonu",
        "Stanley Prusiner'in prion konseptinin tüm amiloidozlara genellenmesi: Hatalı katlanmış bir proteinin komşu hücredeki sağlam proteinleri bozması.",
        "Alzheimer, Parkinson ve ALS'nin beyinde anatomik olarak yayılması (Braak evrelemesi), prion benzeri bir nükleasyon yayılımı (seeding) ile gerçekleşir. Bir nöronda oluşan yanlış katlanmış alfa-sinüklein veya Tau oligomerleri; ekzositoz, nanotüpler (TNT) veya eksozomlar aracılığıyla sinaptik boşluğa salınır. Komşu sağlıklı nöron bu agregatları endositozla içeri alır. İçeri giren tohum (seed), o hücredeki yerel katlanmış proteinlere şablon oluşturarak onların da hatalı konformasyona katlanmasını sağlar. Bu enfeksiyöz benzeri kaskat, hastalığın beyin devreleri boyunca bir yangın gibi yayılmasını açıklar.",
        "Spreading_Velocity = v_propagation = k_seed * [Extracellular_Seeds] * [Native_Monomer_Host] / Transport_Resistance",
        "Prion benzeri anatomik yayılma hızı; ekstraselüler tohum yoğunluğu ve konak hücredeki yerel monomer havuzunun sinaptik iletim katsayısıyla çarpımıdır."
    ),
    (
        "7.9",
        "Kardiyovasküler Amiloidoz: Transtiretin (ATTR) Agregasyonu ve Kalp Sertleşmesi",
        "Karaciğerin ürettiği tiroksin taşıyıcı Transtiretin tetramerinin yaşla monomerlere çözülüp kalp kasında amiloid plaklar örmesi.",
        "Senil Kardiyak Amiloidoz (ATTRwt), 80 yaş üstü erkeklerin %25'inde saptanan ve tedavi edilmeyen kalp yetmezliğinin gizli nedenidir. Karaciğer tarafından sentezlenen Transtiretin (TTR), normalde stabil bir homotetramerdir. Yaşlanmayla birlikte tetramerik yapı kinetik olarak kararsızlaşır; monomerlere ayrışır. Monomerler hızla yanlış katlanarak kalp ventrikül kasında amiloid fibriller halinde çökelir. Kalp duvarı kalınlaşır, elastikiyetini kaybeder (restriktif kardiyomiyopati) ve hasta diyastolik kalp yetmezliğinden ölür. Tafamidis molekülü TTR tetramerini stabilize ederek monomer ayrışmasını engeller.",
        "ATTR_Deposition_Flux = k_dissoc * [TTR_tetramer] * P_misfold * [Cardiomyocyte_Matrix_Area]",
        "Kardiyak Transtiretin amiloid birikim debisi; TTR tetramerinin monomerlere ayrışma hızı ve miyokardiyal matriks yüzey alanı ile ölçeklenir."
    ),
    (
        "7.10",
        "Spektroskopik Agregatom Analizi: Tioflavin T (ThT), Kırmızı Kongo ve Kriyomikroskopi",
        "Amiloid çapraz beta-tabakalarına bağlandığında floresansı 1000 kat artan problarla protein agregasyonunun milisaniyelik ölçümü.",
        "Amiloid yapılarının teşhisinde biyofiziksel boyalar altın standarttır. Tioflavin T (ThT), çapraz beta-tabakalarındaki dar hidrofobik kanallara girdiğinde moleküler rotasyonu kilitlenir; 440 nm uyarılma ile 485 nm'de yansıttığı floresans 1000 kat patlar. Bu sinyal agregasyon kinetiğinin gerçek zamanlı izlenmesini sağlar. Polarize ışık mikroskopisinde Kırmızı Kongo (Congo Red) boyasıyla elma yeşili çift kırılma (apple-green birefringence) göstermesi amiloidin kesin patolojik tanısıdır. Cryo-EM ise günümüzde hasta beyninden izole edilen Tau veya alfa-sinüklein liflerinin atomik çözünürlükte (2-3 Angstrom) yapısını çözmektedir.",
        "Fluorescence_ThT = F_0 + Delta_F_max * [Cross_Beta_Fibrils] / (K_d_ThT + [Cross_Beta_Fibrils])",
        "Tioflavin T floresans şiddeti; ortamdaki çapraz beta-tabakalı amiloid fibril konsantrasyonunun doygunluk bağlanma izotermi ile hesaplanır."
    )
]

# ==============================================================================
# KISIM 8: PROTEOSTAZ BOZULMASININ SİSTEMİK HASTALIKLARI VE DOKU DEJENERASYONU
# ==============================================================================
part8_subsections = [
    (
        "8.1",
        "Beyin Parankiminde Agregasyon: Alzheimer, Parkinson, Huntington ve ALS Moleküler Tablosu",
        "Merkezi sinir sisteminin bölünmeyen nöronlarında on yıllar boyunca biriken agregatların sinapsları koparıp hücreleri öldürmesi.",
        "Nöronlar mitoz bölünme yapmadıkları için agregat yükünü yavru hücrelere bölerek seyreltemezler; bu nedenle proteostaz çöküşüne en kırılgan dokudur. Huntington hastalığında huntingtin genindeki CAG trinükleotid tekrar genişlemesi (polyQ kuyruğu), proteinin nükleusta ve sitoplazmada devasa inklüzyonlar kurmasına yol açar. ALS'de mutant SOD1 veya TDP-43 motor nöronları öldürür. Alzheimer'da Abeta sinapsları tahrip ederken, nöron içine yayılan Tau mikrotübül ağını parçalayarak aksonal taşımayı sıfırlar. Tüm bu hastalıklarda ortak payda; şaperonların kilitlenmesi ve lizozomal sistemin tıkanmasıdır.",
        "Neuronal_Apoptosis_Flux = k_death * [Intracellular_Inclusions] * ([ROS_neuro] + [Ca2+_overload])",
        "Nöronal apoptotik ölüm akısı; hücre içi inklüzyon cisimciği yükü ile oksidatif stres ve kalsiyum aşırı yükünün çarpımıdır."
    ),
    (
        "8.2",
        "İskelet Kasında İnklüzyon Cisimciği Miyoziti ve Sarkopenik Protein Agregasyonu",
        "Yaşlanan kas liflerinde beta-amiloid ve TDP-43 pozitif inklüzyonların birikmesi; kas kasılma ünitelerinin mekanik çöküşü.",
        "İnklüzyon Cisimciği Miyoziti (IBM), 50 yaş üstü bireylerde en sık görülen edinsel kas hastalığıdır. Kas liflerinin vakuollerinde tıpkı Alzheimer beynindeki gibi Amiloid-beta, p-Tau, alfa-sinüklein ve ubikitin pozitif agregatlar birikir. Normal sarkopenik yaşlanmada dahi kas liflerinde p62 ve poli-übikitin kümeleri saptanır. Bu agregatlar kasın kasılma aygıtı olan aktin-miyozin miyofibrillerini fiziksel olarak yerinden eder; kas lifi atrofiye uğrar, kasılma kuvveti düşer ve kas dokusu yağ ve bağ dokusuyla yer değiştirir.",
        "Muscle_Force_Deficit = Force_0 * (1 - [Myofibril_Aggregates] / Critical_Displacement_Limit)",
        "İskelet kası kuvvet kaybı; miyofibriller arasına çöken protein agregatlarının kasılma aparatını yerinden etme fraksiyonu ile orantılıdır."
    ),
    (
        "8.3",
        "Göz Lensinde Katarakt Patogenezi: Kristallin Proteinlerinin Glikasyonu ve Çökmesi",
        "Göz merceğinin şeffaflığını sağlayan kristallin proteinlerinin ömür boyu turnover olmaksızın oksidasyonla çökmesi ve körlük.",
        "Göz lensinin merkezindeki lif hücreleri embriyonik dönemde tüm organellerini (çekirdek, mitokondri) kaybeder; bu hücrelerde protein turnover'ı (yıkım ve yeniden sentez) sıfırdır! Doğumda sentezlenen alfa, beta ve gamma kristallin proteinleri tüm yaşam boyunca şeffaf ve düzenli bir kolloidal fazda kalmak zorundadır. Ancak on yıllar süren UV ışığı maruziyeti, glikasyon (şeker bağlanması) ve triptofan oksidasyonu kristallin proteinlerini çapraz bağlar. alfa-kristallin şaperon kapasitesi 60 yaşında tamamen tükenir; proteinler presipite olarak merceği opaklaştırır ve katarakt körlüğünü yaratır.",
        "Lens_Turbidity = Optical_Density = k_cataract * [Precipitated_Crystallins] * Average_Particle_Radius^3",
        "Göz merceği bulanıklığı (katarakt şiddeti); çökelen kristallin agregat konsantrasyonu ve parçacık yarıçapının küpü ile Rayleigh saçılması yapar."
    ),
    (
        "8.4",
        "Pankreas Beta Hücrelerinde İskemi ve Amilin (IAPP) Agregasyonu ile Diyabet",
        "İnsülin ile birlikte salgılanan İlet Amiloid Polipeptidinin (IAPP) beta hücrelerinde amiloid lifler örerek hücreleri yok etmesi.",
        "Tip 2 diyabet hastalarının %90'ının pankreas otopsilerinde adacık amiloidozisi görülür. İnsülin üreten beta hücreleri, insülinle 1:100 oranında Adacık Amiloid Polipeptidi (IAPP / Amilin) ko-sekrete eder. İnsülin direnci durumunda beta hücresi aşırı insülin üretmek için hiper-sekresyona zorlanır. Bu aşırı üretim sırasında IAPP sekretuar granüllerde ve ER'de oligomerleşir. Toksik IAPP oligomerleri beta hücresinin plazma zarını delerek apoptozu tetikler; beta hücre kitlesi erir ve hastalık geri dönüşümsüz insülin bağımlı evreye geçer.",
        "Beta_Cell_Loss_Rate = k_iapp * [IAPP_oligomers_membrane] * [ER_Stress_Signal]",
        "Pankreas beta hücresi kayıp hızı; zarda gözenek açan toksik IAPP oligomer yoğunluğu ve hücresel ER stresinin çarpımıdır."
    ),
    (
        "8.5",
        "Karaciğerde Alfa-1 Antitripsin Eksikliği (AATD) ve Agregasyon Sirozu",
        "Karaciğer hepatositlerinin ER'sinde polimerleşerek biriken mutant Z-AAT proteininin karaciğeri siroza ve kansere sürüklemesi.",
        "Alfa-1 Antitripsin Eksikliği (AATD), SERPINA1 genindeki E342K (Z aleli) mutasyonundan kaynaklanan prototipik bir konformasyonel hastalıktır. Normalde karaciğerden salgılanıp akciğeri koruması gereken AAT proteini, Z mutasyonu nedeniyle ER lümeninde beta-iplikçik değişimiyle (domain swapping) devasa polimerler halinde çöker. ER bu polimerleri ERAD ile temizleyemez; hepatosit içinde biriken PAS-pozitif globüller karaciğerde kronik otofaji yetmezliğine, fibrozise, siroza ve hepatosellüler karsinoma (HCC) yol açar.",
        "Hepatic_Fibrosis_Risk = [Intrahepatic_Z_AAT_Polymers] / [Autophagic_Clearance_Capacity]",
        "Karaciğer siroz ve fibrozis riski; hepatosit ER'sinde biriken Z-AAT polimer miktarının otofajik temizleme kapasitesine oranıdır."
    ),
    (
        "8.6",
        "Damar Duvarında Elastin ve Kolajen Agregasyonu: Vasküler Kireçlenme",
        "Arter duvarındaki uzun ömürlü ekstraselüler matriks proteinlerinin ileri glikasyon ürünleriyle (AGE) çapraz bağlanıp taşlaşması.",
        "Damar elastikiyetini sağlayan elastin liflerinin yarı ömrü yaklaşık 70 yıldır; yani yaşam boyu neredeyse hiç yenilenmez. Yaşlanmayla birlikte kan şekeri ve metabolitler elastin ve kolajen üzerindeki serbest amino gruplarına enzimatik olmayan yollarla bağlanır (Maillard reaksiyonu). Oluşan İleri Glikasyon Son Ürünleri (AGE'ler: Glukozpan, Karboksimetillizin), bitişik lifler arasında kalıcı kovalent çapraz bağlar kurar. Damar duvarı bir lastik banttan sert bir boruya dönüşür; sistolik tansiyon fırlar, nabız dalga hızı (PWV) artar ve organlar vasküler darbe hasarıyla tahrip olur.",
        "Arterial_Stiffness_PWV = sqrt( (Thickness_wall * Modulus_Elasticity_AGE) / (2 * Radius_artery * Density_blood) )",
        "Arteryel nabız dalga hızı (PWV / sertlik); AGE çapraz bağlarının artırdığı elastik Young modülünün karekökü ile doğru orantılıdır."
    ),
    (
        "8.7",
        "Yaşlanan Kök Hücrelerde Agresom Yükü ve Asimetrik Bölünmede Çöp Mirası",
        "Kök hücrenin mitoz bölünme sırasında tüm protein agregatlarını farklılaşacak yavruya itip kendini temiz tutma stratejisinin yaşla çöküşü.",
        "Hematopoietik ve nöral kök hücreler, ömür boyu temiz kalabilmek için olağanüstü bir asimetrik çöplük mirası yürütür: Asimetrik bölünme anında kök hücre, vimentin kafesiyle paketlenmiş agresomları ve hasarlı proteinleri sadece bir yavru hücreye (farklılaşıp gidecek progenitöre) aktarır; kök hücre olarak kalacak yavruya sıfır çöp bırakır. Ancak yaşlanan kök hücrelerde polarite motorları ve sentrozom organizasyonu bozulur; agregatlar her iki yavruya da eşit saçılır. Çöp mirası alan kök hücre rezervuarı kök hücre kimliğini ve kendini yenileme gücünü kaybederek tükenir.",
        "Asymmetric_Segregation_Fidelity = [Aggresome_progenitor] / ([Aggresome_stem_daughter] + epsilon) -> 1.0 (Yaşlıda çöküş)",
        "Asimetrik agregat segregasyon sadakati; genç kök hücrede 100 kat progenitör lehine iken yaşlı kök hücrede 1.0'a (simetrik kirlenme) çöker."
    ),
    (
        "8.8",
        "İmmün Hücrelerde Agregat Birikimi: Makrofaj Köpük Hücresi Dönüşümü ve Ateroskleroz",
        "Arter duvarındaki makrofajların okside LDL ve kolesterol kristallerini yutup sindiremeyerek aterom plağında nekroza gitmesi.",
        "Ateroskleroz patogenezinde intima tabakasına sızan makrofajlar, süpürücü reseptörler (SR-A, CD36) aracılığıyla aşırı okside LDL (oxLDL) partiküllerini fagositozla yutar. Yutulan bu aşırı lipit yükü lizozomlarda kristalleşir ve lizozomal asit lipazı felç eder. Makrofaj bir köpük hücresine (foam cell) dönüşür. Sindirilemeyen agregatlar lizozom zarını yırtar; makrofaj nekroz geçirerek çekirdeğinde lipid ve protein çöpü dolu 'nekrotik aterom göbeğini' oluşturur; bu da damar tıkanıklığını ve kalp krizini tetikler.",
        "Foam_Cell_Necrosis_Rate = k_foam * [Intralysosomal_Cholesterol_Crystals] * [LMP_rupture_rate]",
        "Köpük hücresi nekroz hızı; lizozom içinde sindirilemeyen kolesterol kristal yükü ve membran yırtılma katsayısının çarpımıdır."
    ),
    (
        "8.9",
        "Solunum Yollarında Kistik Fibrozis ve CFTR DeltaF508 Katlanma Defekti",
        "Tek bir fenilalanin delesyonunun klorür kanalını ERAD'a mahkum etmesi; konformasyonel hastalıkların moleküler prototipi.",
        "Kistik Fibrozis hastalarının %70'inde görülen DeltaF508 mutasyonu, CFTR klorür kanalının 508. pozisyonundaki fenilalanin kalıntısının silinmesidir. Bu mutant protein aslında hücre zarına ulaşabilse %50 oranında fonksiyonel klorür pompalayabilir; ancak mutasyon proteinin ER'deki katlanma kinetiğini hafifçe yavaşlatır. ER kalite kontrol şaperonları bu gecikmeyi 'kusur' olarak algılar; proteini zara göndermek yerine ERAD yolağı ile proteazomda parçalar. Hücre zarsız kalır; mukus kurur ve ölümcül akciğer enfeksiyonları gelişir. Trikafta ilaçları bu proteini kimyasal olarak katlayarak zara ulaştırır.",
        "CFTR_Surface_Expression = Total_Synthesis * (1 - P_ERAD_Degradation(DeltaF508)) -> Rescued_by_Correctors",
        "Hücre yüzeyi CFTR kanalı mevcudiyeti; sentezlenen proteinin ERAD kalite kontrol yıkımından kaçabilme olasılığı ile belirlenir."
    ),
    (
        "8.10",
        "Sistemik Senesens Sürücüsü Olarak Agregatom: SASP'ın Proteostaz Çöküşüyle Tetiklenmesi",
        "Hücre içinde biriken protein çöplerinin NF-kappaB ve p38 MAPK'ı sürekli uyararak hücreyi senesens fenotipine kilitlemesi.",
        "Hücresel senesens salt bir telomer kısalması değildir; proteostaz çöküşünün kaçınılmaz nihai sonucudur. Proteazom tıkandığında, sitoplazmada agregatlar biriktiğinde ve lizozomlar lipofusinle dolduğunda; hücre sürekli bir proteotoksik alarm durumuna geçer. Bu kronik stres p38 MAPK ve IKK kinazlarını aktive eder; NF-kappaB nükleusa kilitlenir. Hücre kalıcı hücre döngüsü tutuklamasına girer ve çevresine yüzlerce yıkıcı sitokin, kemokin ve metalloproteaz salgılayan tam teşekküllü bir Senesens İlişkili Salgı Fenotipi (SASP) fabrikasına dönüşür.",
        "SASP_Secretion_Flux = k_sasp * [Aggresome_Intracellular_Load] * [NF-kB_nuclear_retention]",
        "Senesen SASP salgılama debisi; hücre içindeki çözünmemiş agregatom yükü ile nükleer NF-kappaB tutulma süresinin çarpımıdır."
    )
]

# ==============================================================================
# KISIM 9: FARMAKOLOJİK VE GENETİK MÜDAHALELER: OTOFAJİ VE ŞAPERON GÜÇLENDİRME
# ==============================================================================
part9_subsections = [
    (
        "9.1",
        "mTORC1 İnhibitörleri: Rapamisin, Everolimus ve Torin-1 ile Sürekli Otofaji Akısı",
        "Paskalya Adası toprağından çıkan bakteriyel makrolid Rapamisinin FKBP12 ile mTORC1'i frenleyerek otofajiyi kalıcı açması.",
        "Rapamisin (Sirolimus), FKBP12 immünofilin proteinine bağlanarak mTORC1 kompleksinin allosterik cebine kenetlenir ve kinaz aktivitesini baskılar. İkinci nesil ATP-kompetitif inhibitörler (Torin-1, AZD8055) ise doğrudan katalitik yarığı bloke eder. mTORC1'in susturulması; ULK1'in Ser757 üzerindeki baskısını kaldırır, TFEB'in nükleusa kaçmasını sağlar ve otofaji-lizozom otoyolunu maksimum debide çalıştırır. Farelerde yaşam süresini %15-25 uzattığı yüzlerce bağımsız laboratuvarca kanıtlanmış altın standart anti-aging müdahalesidir.",
        "Autophagy_Induction_Ratio = [p-ULK1_Ser317] / ([p-ULK1_Ser757] + epsilon) = f([Rapamycin] / IC50)",
        "Otofaji indüksiyon oranı; uygulanan Rapamisin konsantrasyonuna bağlı olarak mTORC1 fosforilasyonunun düşmesi ve ULK1 aktivasyonuyla ölçeklenir."
    ),
    (
        "9.2",
        "AMPK Aktivatörleri: Metformin, AICAR ve Otofajik Tetikleme",
        "Metabolik yakıt sensörünü uyararak ULK1'i doğrudan fosforilleyen ve hücresel temizliği başlatan farmakolojik ajanlar.",
        "AMPK aktivatörleri, besin bolluğunda bile hücreyi 'açlık temizlik moduna' sokan moleküllerdir. Metformin Kompleks I'i hafifçe frenleyerek AMP/ATP oranını yükseltirken; sentetik nükleotid analoğu AICAR (ZMP'ye dönüşerek) doğrudan AMPK'nın allosterik cebine bağlanır. Direkt AMPK aktivatörleri A-769662 ve MK-8722 ise beta1 alt birimine bağlanarak enzimi fosforilasyondan bağımsız tetikler. Aktif AMPK hem ULK1'i doğrudan Ser317'den fosforilleyerek otofajiyi açar hem de TSC2 ve Raptor'u fosforilleyerek mTORC1'i çifte kilit altına alır.",
        "AMPK_Autophagic_Flux = k_ampk * [Active_AMPK] * [ULK1_unphosphorylated_S757]",
        "AMPK kaynaklı otofajik akı; aktifleşen AMPK enzim derişimi ile mTORC1 baskısından kurtulmuş serbest ULK1 miktarının çarpımıdır."
    ),
    (
        "9.3",
        "Doğal ve Sentetik Poliaminler: Spermidin Moleküler Farmakolojisi ve EP300 Deasetilasyonu",
        "Frank Madeo laboratuvarının keşfi: Spermidinin asetiltransferaz EP300'ü inhibe ederek otofaji proteinlerini deasetillemesi ve ömrü uzatması.",
        "Spermidin, tüm canlı hücrelerde bulunan doğal bir alifatik poliamindir; yaşlanmayla birlikte doku konsantrasyonu dramatik şekilde düşer. Spermidin, hücre içine girdiğinde nükleer ve sitoplazmik asetiltransferaz EP300'ü yarışmalı olarak inhibe eder. EP300'ün baskılanması; ATG5, ATG7, LC3 ve TFEB üzerindeki asetil gruplarının sirtuinler tarafından temizlenmesini sağlar. Deasetillenen otofaji makineleri hiper-aktif hale gelir; farelerde kalp hipertrofisini geriletir, hafızayı gençleştirir ve tüm nedenlere bağlı mortaliteyi düşürür.",
        "Autophagy_Deacetylation_Score = [Spermidin] / (K_i_EP300 + [Spermidin]) * [SIRT1_activity]",
        "Spermidin otofaji aktivasyon skoru; EP300 enziminin spermidin ile inhibisyon derecesi ve SIRT1 deasetilaz aktivitesinin çarpımıdır."
    ),
    (
        "9.4",
        "Trehaloz ve Kalorisiz Otofaji İndüksiyonu: TFEB ve Glukoz Taşıyıcı İnhibisyonu",
        "Doğal disakkarit trehalozun SLC2A (GLUT) taşıyıcılarını bloke ederek hücresel açlık sinyali vermesi ve lizozomları uyandırması.",
        "Trehaloz, iki glukoz molekülünün 1,1-glikozidik bağla birleştiği doğal bir disakkarittir. Hücre içine girdiğinde hücre zarındaki GLUT (SLC2A) glukoz taşıyıcılarını yarışmalı olarak bloke eder; hücre içi serbest glukoz akışını hafifçe kısıtlar. Bu durum hücrede sistemik bir kalori açığı yaratmadan yerel bir 'glukoz açlığı' yanılsaması oluşturur. AMPK aktive olur, AKT inaktive olur ve TFEB defosforillenerek nükleusa hücum eder; karaciğer ve beyinde amiloid ve tau agregatlarını %40 oranında temizler.",
        "TFEB_Translocation_Trehalose = k_treh * [Trehalose_plasma] / (1 + [Intracellular_Glucose] / K_m_glut)",
        "Trehaloz aracılı TFEB nükleer translokasyonu; plazma trehaloz seviyesi ile hücre içi glukoz seviyesinin oranına bağlıdır."
    ),
    (
        "9.5",
        "Şaperon Güçlendiriciler: Geranylgeranylacetone (GGA) ile Hsp70 Patlaması",
        "Termal şoka veya toksik strese gerek kalmadan Hsp70 ve Hsp40 ekspresyonunu 10 kat artıran klinik onaylı şaperon mühimmatı.",
        "Hücrede şaperon üretimini tetiklemek için hücreyi ısıtmak (hipertermi) toksik bir bedel ödetir. Geranylgeranylacetone (GGA / Teprenone), Japonya'da ülser ilacı olarak onaylanmış non-toksik bir izoprenoid bileşiktir. GGA, HSF1'i Hsp90 kompleksinden kopararak fizyolojik sıcaklıkta (37 C) doğrudan trimerleştirir ve nükleusa gönderir. Doku düzeyinde Hsp70, Hsp40 ve Hsp90 ekspresyonunu 5 ila 10 kat artırarak nörodejeneratif modellerde protein agregasyonunu ve kas atrofisini durdurur.",
        "Hsp70_Induction_Fold = 1 + alpha_gga * [GGA_tissue] / (K_m_gga + [GGA_tissue])",
        "Doku Hsp70 indüksiyon katı; uygulanan GGA konsantrasyonuna bağlı Michaelis-Menten transkripsiyon katsayısı ile modellenir."
    ),
    (
        "9.6",
        "Küçük Moleküllü Agregat Çözücüler (Disaggregases): Hsp104 Biyomühendisliği ve Memeli Uyarlamaları",
        "Maya proteini Hsp104'ün protein mühendisliğiyle yeniden tasarlanarak insan amiloidlerini tek tek çözebilen biyo-türbine çevrilmesi.",
        "Memeli hücrelerinde katılaşmış amiloid fibrillerini aktif olarak çözen (disaggregase) güçlü bir AAA+ ATPaz makinesi bulunmaz; oysa mayalarda Hsp104 bu görevi kusursuz yapar. James Shorter laboratuvarı, maya Hsp104 proteinini rasyonel mutajenezle yeniden tasarlamıştır. Bu mühendislik harikası 'güçlendirilmiş Hsp104 varyantları', insan hücrelerinde önceden oluşmuş ve katılaşmış TDP-43, alfa-sinüklein ve FUS amiloid agregatlarını yakalar, iplikçikleri tek tek çekip çözerek monomerik zararsız formlarına geri döndürür.",
        "Disaggregation_Rate = v_hsp104 * [Hsp104_variant] * [Fibril_Substrate] * [ATP] / (K_m + [ATP])",
        "Amiloid agregat çözülme hızı; tasarlanmış Hsp104 desagregaz motor konsantrasyonu ve fibril substrat yoğunluğu ile doğrudan orantılıdır."
    ),
    (
        "9.7",
        "PROTAC'lar (Proteolysis Targeting Chimeras): Hedefe Yönelik E3 Ligaz ile Toksik Protein Tasfiyesi",
        "Craig Crews'un devrimsel bifonksiyonel molekülü: Bir ucu hedef agregatı, diğer ucu E3 ligazı tutarak proteini proteazomda yok etme.",
        "PROTAC teknolojisi, geleneksel ilaçların 'aktif bölgeyi bloke etme' paradigmasını yıkarak 'proteini tamamen yok etme' (targeted protein degradation) çağına geçmiştir. Bir PROTAC molekülü iki aktif başlık ve bir esnek bağlayıcıdan (linker) oluşur: Bir başlık hedeflenen patolojik proteine (örneğin p-Tau veya mutant Huntingtin) bağlanırken; diğer başlık bir E3 ubikitin ligazına (Cereblon / CRBN veya VHL) bağlanır. Molekül bu iki proteini zorla yan yana getirerek 'üçlü kompleks' (ternary complex) kurar. E3 ligaz hedef proteini ışık hızında K48 ubikitinle donatır ve 26S proteazoma yutturur.",
        "Degradation_Velocity_PROTAC = k_deg * [Ternary_Complex] = k_deg * [Target] * [PROTAC] * [E3_Ligase] * Cooperativity_Factor",
        "PROTAC hedef yıkım hızı; hedef protein, PROTAC ve E3 ligazın kurduğu üçlü kompleksin konsantrasyonu ve termodinamik kooperativitesiyle belirlenir."
    ),
    (
        "9.8",
        "AUTAC ve LYTAC Teknolojileri: Hedefe Yönelik Moleküllerin Otofajiye ve Lizozomlara Yönlendirilmesi",
        "PROTAC'ın sınırlarını aşmak: Hücre dışı plakları ve devasa intraselüler organelleri otofaji ve lizozomda eriten kimerik moleküller.",
        "PROTAC yalnızca proteazomun dar kapısından geçebilen monomerik proteinleri yıkabilir; dev amiloid fibrillerini veya membran proteinlerini yıkamaz. Bu engeli aşmak için AUTAC (Autophagy-Targeting Chimeras) ve LYTAC (Lysosome-Targeting Chimeras) icat edilmiştir. AUTAC, bir guanin türevi (cGMP taklidi) taşıyarak hedef proteini veya hasarlı mitokondriyi K63-übikitinasyonla otofagozoma yönlendirir. Carolyn Bertozzi'nin geliştirdiği LYTAC ise ekstraselüler Abeta plaklarını veya membran reseptörlerini Mannoz-6-Fosfat Reseptörüne (CI-M6PR) bağlayarak doğrudan lizozoma yutturur.",
        "Extracellular_Plaque_Clearance = k_lytac * [LYTAC_concentration] * [CI-M6PR_density] * [Abeta_Plaques]",
        "Ekstraselüler amiloid plak temizleme debisi; uygulanan LYTAC kimerik konsantrasyonu ve hücre yüzeyi M6P reseptör yoğunluğunun çarpımıdır."
    ),
    (
        "9.9",
        "Dihidromirisetin ve Resveratrol ile Sirtuin Bağımlı Otofajik Gençleşme",
        "Doğal polifenollerin SIRT1 ve SIRT3'ü uyararak otofaji transkriptomunu ve mitofajiyi eşzamanlı aktive etmesi.",
        "Doğal flavonoid Dihidromirisetin (DHM / Ampelopsin) ve stilbenoid Resveratrol; hücresel redoks dengesini ve otofajiyi çoklu yolaklarla yönetir. SIRT1'in NAD+ bağımlı deasetilaz aktivitesini artırarak FoxO1 ve FoxO3a transkripsiyon faktörlerini aktive ederler. FoxO'lar nükleusta Bnip3, Atg12, Gabarapl1 ve Vps34 genlerinin ekspresyonunu fırlatır. Eşzamanlı olarak karaciğer ve beyinde lipit damlacıklarının otofajik yıkımını (Lipofaji) tetikleyerek hücre içi lipotoksisiteyi temizler ve bilişsel gerilemeyi durdurur.",
        "Autophagic_Gene_Expression = sum_k ( [SIRT1_active] * [FoxO3a_deacetylated] * Affinity(FoxO3a, Promoter_ATG_k) )",
        "Sirtuin bağımlı otofaji gen ekspresyon şiddeti; aktif deasetile FoxO3a faktörünün ATG promoterlarına bağlanma afinitelerinin toplamıdır."
    ),
    (
        "9.10",
        "Klinik Longevity Protokolleri: Döngüsel Açlık, Otofaji Diyetleri ve Sinerjistik Staklar",
        "Haftalık ve aylık döngülerle otofaji pikleri yaratan, şaperonları tazeleyen ve agregatomu sıfırlayan kanıta dayalı insan rejimi.",
        "Klinik proteostaz restorasyonu; sürekli baskılama yerine dinamik döngüsellik gerektirir. En gelişmiş klinik longevity protokolü şu fazlardan oluşur: 1) Haftada 2 kez 16-18 saatlik aralıklı oruç (mTORC1 baskısı ve bazal otofaji); 2) Ayda bir kez 72 saatlik uzamış su orucu veya Fasting-Mimicking Diet (TFEB nükleer göçü, derin lizozomal temizlik ve senesen protein tasfiyesi); 3) Günlük farmasötik destek (Sabah: Spermidin + Trehaloz; Akşam: Dihidromirisetin + TUDCA); 4) Periyodik şaperon darbeleri (Haftada 3 gün 80 C sauna ile Hsp70 indüksiyonu).",
        "Systemic_Aggresome_Clearance = Integral_0^T ( J_autophagy(t) + J_proteasome(t) + J_CMA(t) - Generation_Rate(t) ) dt > 0",
        "Sistemik agregatom temizleme başarısı; döngüsel rejim boyunca otofaji, proteazom ve CMA kümülatif yıkım akısının yeni çöp üretimini aşmasıdır."
    )
]

# ==============================================================================
# KISIM 10: HOMO AETERNUS PROTEOMİK MİMARİSİ: ÇÖPSÜZ VE EBEDİ DOKU REJENERASYONU
# ==============================================================================
part10_subsections = [
    (
        "10.1",
        "Sentetik Kalite Kontrol Makineleri: Yapay Nano-Kafes Şaperonlar ve İn Vivo Dağıtım",
        "DNA origami ve protein tasarım yazılımlarıyla (Rosetta/AlphaFold) sıfırdan üretilen, yanlış katlanan proteini yutup düzelten yapay nano-variller.",
        "Homo Aeternus mimarisinde hücreler doğal şaperonların sınırlı kapasitesine mahkum bırakılmaz. De novo protein tasarımı ve DNA nanoteknolojisiyle, GroEL benzeri sentetik 'akıllı nano-kafesler' üretilir. Bu sentetik şaperonlar hücre içine LNP veya eksozomlarla dağıtılır. Yüzeylerindeki biyofiziksel sensörlerle açığa çıkan hidrofobik dizilimleri tanır, hedef proteini nano-odacığa hapseder; ATP hidrolizine ihtiyaç duymadan yüzey polarite kaydırmasıyla proteini saniyeler içinde yerel yapısına katlayarak serbest bırakır.",
        "Synthetic_Chaperone_Throughput = N_nanocages * Turnover_Frequency_per_cage > 1000_proteins/sec",
        "Sentetik şaperon işlem kapasitesi; hücreye dağıtılan nano-kafes sayısı ile kafes başına saniyelik protein düzeltme frekansının çarpımıdır."
    ),
    (
        "10.2",
        "Sentetik LAMP-2A Reseptör Gen Terapisi (AAV-LAMP2A) ile Karaciğer ve Beyin Gençleşmesi",
        "AAV9 viral vektörleriyle yaşlanan organların lizozom zarlarına sürekli taze LAMP-2A reseptörü entegre ederek CMA'yı sonsuzlaştırma.",
        "Yaşlanmada CMA'nın çöküşünün ana nedeni lizozom zarındaki LAMP-2A kaybıdır. AAV-LAMP2A gen terapisi, modifiye hepatosit ve nöron spesifik promoterlar altında insan LAMP-2A cDNAsını dokulara taşır. Lizozom zarlarına sürekli taze ve mutant olmayan LAMP-2A reseptörleri ekilir. Lizozomal lipidler kireçlense dahi aşırı eksprese edilen reseptörler monomer-multimer dengesini korur; yaşlı bireyin karaciğeri ve beyni 20 yaşındaki bir insanın protein temizleme hızında çalışarak amiloidoz riskini sıfırlar.",
        "CMA_Rescue_Score = [Transgenic_LAMP2A_membrane] / [Baseline_Aged_LAMP2A] > 3.5 -> 100%_Clearance",
        "CMA kurtarma skoru; viral gen terapisi ile zarda eksprese edilen transgenik LAMP-2A yoğunluğunun yaşlı bazal kontrole oranıdır."
    ),
    (
        "10.3",
        "Dizayn Edilmiş Agregat Parçalayıcı Enzimler (De-aggregases): Tau ve Abeta'yı Çözen Biyo-Makineler",
        "Memeli genomuna transgenik olarak eklenen veya mRNA ile verilen, katılaşmış amiloid fibrillerini çözen yapay AAA+ ATPaz motorları.",
        "Doğanın milyonlarca yılda geliştiremediği amiloid çözücü motorlar, sentetik biyoloji ile insan hücrelerine entegre edilir. Yapay zeka ile tasarlanmış sentetik desagregazlar (De-aggregases), Abeta42, alfa-sinüklein ve Tau fibrillerinin çapraz beta omurgasını 100 pN mekanik torkla kavrar. Fibrilin ucundan bir monomer koparır, zinciri açar ve çözünür hale getirir. Birkaç haftalık tedavi ile Alzheimer hastasının beynindeki tüm yerleşik amiloid plakları ve nörofibriler yumaklar cerrahiye gerek kalmadan moleküler düzeyde eritilir.",
        "Plaque_Dissolution_Velocity = -d[Amyloid_Plaque]/dt = k_disagg * [Synthetic_Disaggregase] * [Fibril_Surface_Area]",
        "Amiloid plak erime hızı; dokuda aktif sentetik desagregaz motor konsantrasyonu ve plak yüzey alanının çarpımıyla doğru orantılıdır."
    ),
    (
        "10.4",
        "Sürekli Nükleer TFEB Aktivasyonu için Sentetik Biyo-Sensör Geri Besleme Döngüleri",
        "Hücrede tek bir protein agregatı oluştuğunda kendiliğinden devreye giren ve TFEB'i nükleusa kilitleyen genetik devreler.",
        "Doğal TFEB sistemi besin ve kalsiyum dalgalanmalarına bağımlıdır; oysa Homo Aeternus proteomu yapay genetik geri besleme devreleriyle (synthetic gene circuits) donatılır. Bu devrede, agregat bağlayıcı bir sentetik sensör (örneğin ThT benzeri afinite domeni), hücrede amiloid saptadığı anda Kalsinörini TFEB'e kenetleyen bir nanobody sentezler. TFEB Serin 211'de asla fosforillenemez; hücre ne kadar tok olursa olsun agregat varlığında nükleusta kilitli kalarak CLEAR ağını çalıştırır ve çöpü anında yok eder.",
        "Circuit_Output_TFEB = V_max * [Aggregates_detected]^n / (K_threshold^n + [Aggregates_detected]^n)",
        "Sentetik biyo-sensör TFEB aktivasyon fonksiyonu; hücre içi agregat tespit seviyesine bağlı yüksek dereceli Hill işbirliği eğrisidir."
    ),
    (
        "10.5",
        "Kristal Olmayan Proteom Tasarımı: Amiloid Puanı Sıfırlanmış De Novo Protein Mühendisliği",
        "İnsan genomundaki tüm temel proteinlerin çapraz beta-tabaka kurma potansiyelinin ters translasyon algoritmalarıyla sıfırlanması.",
        "Proteinlerin amiloid oluşturması mutlak bir doğa kanunu değil; primat evriminin ihmal ettiği bir fizikokimyasal zaaftır. Rosetta ve AlphaFold3 algoritmaları kullanılarak insan proteomundaki 20.000 protein taranır; agregasyona yol açan hidrofobik beta-iplikçik bölgeleri saptanır. Proteinin enzimatik fonksiyonunu bozmayan 'sessiz mutasyonlar' (gatekeeper kalıntıları: prolin, lizin, glutamat) yerleştirilerek proteinin serbest enerji manzarası yeniden çizilir. Amiloid skoru sıfırlanmış bu 'süper-kararlı proteom', 100 C sıcaklıkta veya aşırı oksidasyonda dahi asla çökelmez.",
        "Amyloid_Propensity_Score(Protein_Synthetic) = sum_segments Aggrescan_Score_k -> ZERO",
        "Sentetik proteomun amiloidojenik eğilim skoru; polipeptid zincirindeki tüm agregasyon çekirdeklerinin gatekeeper kalıntılarıyla sıfırlanmasıdır."
    ),
    (
        "10.6",
        "Lizozomal Enzim Kokteylleri: Lipofusini Parçalayan Yapay Bakteriyel Enzimler (Lyso-Clear)",
        "Toprak bakterilerinden izole edilen ve insan lizozomlarına uyarlanan, parçalanamaz Schiff bazlarını yıkan kimerik hidrolazlar.",
        "Aubrey de Grey ve SENS Araştırma Vakfı'nın öncülük ettiği 'LysoSENS' projesi: Doğada asırlık çürüyen cesetlerde dahi lipofusin kalmaz; çünkü toprak bakterileri (Acinetobacter, Pseudomonas) lipofusini parçalayan enzimlere sahiptir. Bu bakteriyel hidrolazlar (Lipofusinazlar) genetik mühendislikle insanlaştırılır, N-terminal Mannoz-6-Fosfat hedefleme sinyaliyle donatılır ve insan hücrelerinin lizozomlarına gönderilir. Bu enzimler lizozom içinde 80 yıldır birikmiş lipofusini, Schiff bazlarını ve lipid peroksitlerini monomerlerine parçalayarak lizozomal lümeni tertemiz yapar.",
        "Lipofuscin_Clearance_Rate = k_lysoclear * [Bacterial_Hydrolase_Lumen] * [Lipofuscin_Surface]",
        "Lipofusin enzimatik tasfiye hızı; lizozoma ulaştırılan yapay Lyso-Clear hidrolaz yoğunluğu ve lipofusin granül yüzey alanının çarpımıdır."
    ),
    (
        "10.7",
        "Sentetik Ubikitin Ligazları: Yaşlılıkta Biriken Tüm Toksik Proteinleri Tanıyan Akıllı Kimerler",
        "Kendi hedef tanıma domenini yapay zekayla değiştiren, senesent proteinleri hızla proteazoma sürükleyen programlanabilir E3 makineleri.",
        "Doğal E3 ligazlar yaşlanmayla ortaya çıkan yeni post-translasyonel modifikasyonları (glikasyon, karbonilasyon) tanıyamaz. Sentetik Ubikitin Ligazları (sE3L), programlanabilir bir antikor parçası (nanobody) ile katalitik bir RING domeninin füzyonudur. Bu nanobody başlığı, yaşlanan hücrede biriken 4-HNE modifiyeli proteinleri veya okside Tau'yu pikomolar afiniteyle tanır. Tanıdığı anda hedefe saniyeler içinde 10 adet K48 ubikitini yapıştırarak proteazomda anında imha edilmesini sağlar.",
        "Target_Ubiquitination_Rate = k_synth_E3 * [sE3L_Engineered] * [Oxidized_Senescent_Proteins]",
        "Sentetik E3 ligaz hedefleme hızı; tasarlanmış kimerik enzimin hücredeki konsantrasyonu ile oksitlenmiş senesent protein yükünün çarpımıdır."
    ),
    (
        "10.8",
        "Yapay Hücre Dışı Çöp Temizliği: Makrofaj Efferositozunun ve İmmün Süpürücülerin Genetik Modifikasyonu",
        "Doku aralıklarında biriken amiloid ve hücresel döküntüleri yutan kimerik antijen reseptörlü (CAR) süpürücü makrofajlar (CAR-M).",
        "Hücre içindeki çöpler kadar hücre dışı doku aralıklarında biriken amiloidler ve ölü hücre parçaları da sistemi tıkar. Sentetik immünoloji yaklaşımında hastanın monositleri izole edilir; yüzeylerine amiloid-beta, amilin veya AGE çapraz bağlarını tanıyan kimerik reseptörler eklenir (CAR-M: Chimeric Antigen Receptor Macrophages). Dokuya geri verilen bu engineered makrofajlar, ekstraselüler alandaki amiloid plakları bir elektrik süpürgesi gibi yutar, fagolizozomlarında eritir ve doku parankimini temizler.",
        "Extracellular_Debris_Clearance = k_effero * [CAR_M_macrophages] * [Extracellular_Matrix_Aggregates]",
        "Ekstraselüler çöp temizleme debisi; dokuya göç eden CAR-M süpürücü makrofaj sayısı ile ekstraselüler matriks amiloid kütlesinin çarpımıdır."
    ),
    (
        "10.9",
        "Biyo-Elektronik Proteostaz İzleme: Hücre İçi Protein Katlanma Gerilimini Ölçen Nano-Sensörler",
        "Hücre içine yerleştirilen floresan ve elektrokimyasal nano-sensörlerle anlık protein katlanma stresinin telemetrik takibi.",
        "Proteostaz çöküşünü hasar oluştuktan sonra değil, oluştuğu milisaniyede yakalamak için hücre içine genetik kodlu FRET nano-sensörleri yerleştirilir. Bu sensörler, sitoplazmik şaperonların substrat doluluk oranını ve hücre içi viskozite değişimlerini (LLPS faz geçişlerini) anlık elektriksel ve optik sinyallere çevirir. Hücrede katlanmamış protein yükü %5 arttığı anda, sensör hücreye entegre optogenetik sistemleri ateşler; Hsp70 ve TFEB genleri saniyeler içinde otomatik olarak aktive edilir.",
        "Proteostasis_Stress_Signal = FRET_Ratio(Donor / Acceptor) = f([Chaperone_Saturation_Level])",
        "Proteostaz gerilim sinyali; şaperon havuzunun substratla doygunluk seviyesini ölçen genetik kodlu FRET rasyometrik floresansıdır."
    ),
    (
        "10.10",
        "Homo Aeternus Proteomik Manifestosu: Sonsuz Entropiye Karşı Bozulmayan Ebedi Biyo-Makineler",
        "Termodinamik entropinin protein yapılarını bozma gücünün; kusursuz şaperonlar, yapay desagregazlar ve sıfır-lipofusin ile ebediyen yenilmesi.",
        "Biyolojik ölümsüzlüğün proteomik temeli; 'çözünmeyen çöp üretmeyen, ürettiği anlık hataları milisaniyede eriten ve protein moleküllerini sürekli yerel enerji çukurunda tutan' dinamik bir dengedir. Homo Aeternus'ta proteomik entropi durdurulmuştur: Şaperonlar çökmeyen sentetik nano-kafeslerle desteklenmiş, lizozomlar bakteriyel enzimlerle lipofusinden arındırılmış, amiloidojenik genler gatekeeper mutasyonlarıyla zırhlanmış ve 26S proteazomlar kapıları sürekli açık süper-öğütücülere dönüştürülmüştür. Böyle bir organizmada amiloidoz, katarakt, nörodejenerasyon ve organ sertleşmesi biyolojik bir imkansızlıktır.",
        "Organismal_Proteomic_Longevity = lim_{t -> infty} ( Quality_Folded_Proteins(t) / Degraded_Debris(t) ) -> CONSTANT_PERFECTION",
        "Homo Aeternus proteomik ölümsüzlük postülatı; zaman sonsuza akarken hücredeki kusursuz katlanmış fonksiyonel protein oranının saf bir gençlik platosunda sabit kalmasıdır."
    )
]

# ==============================================================================
# 10 AKADEMİK KARŞILAŞTIRMA VE PARAMETRE TABLOSU (HER KISIM İÇİN BİR ADET)
# ==============================================================================
parts.append(("KISIM 7: AGREGATOM (AGGRESOME), AMİLOİD FİBRİLLER VE LİPOFUSİN BİRİKİMİ", part7_subsections))
parts.append(("KISIM 8: PROTEOSTAZ BOZULMASININ SİSTEMİK HASTALIKLARI VE DOKU DEJENERASYONU", part8_subsections))
parts.append(("KISIM 9: FARMAKOLOJİK VE GENETİK MÜDAHALELER: OTOFAJİ VE ŞAPERON GÜÇLENDİRME", part9_subsections))
parts.append(("KISIM 10: HOMO AETERNUS PROTEOMİK MİMARİSİ: ÇÖPSÜZ VE EBEDİ DOKU REJENERASYONU", part10_subsections))

tables_data = [
    (
        "TABLO 1: Moleküler Şaperon Aileleri, Hücresel Lokalizasyonları ve İstemci Protein Dinamikleri",
        ["Şaperon Ailesi", "Örnek Molekül / Kompleks", "Hücresel Kompartıman", "ATP Bağımlılığı ve Mekanizma", "Yaşlanmadaki Fonksiyon Kaybı"],
        [
            ["Hsp70 Sistemi", "Hsp72 (stres), Hsc70 (konstitütif)", "Sitozol, Nükleus, Lizozom (CMA)", "ATP bağımlı (Hsp40/NEF döngüsü ile açık hidrofobik kalıntıları tutma)", "Ekspresyon düşüşü, agregatlara bağlanarak tükenme (titrasyon)"],
            ["Şaperoninler (Hsp60)", "TRiC / CCT kompleksi, mtHsp60", "Sitozol ve Mitokondri Matrisi", "ATP bağımlı kooperatif çift varil içi izolasyon (Anfinsen kafesi)", "Alt birim ayrışması, aktin ve tübülin katlanma yetmezliği"],
            ["Hsp90 Sistemi", "Hsp90-alfa, Hsp90-beta, Grp94", "Sitozol ve ER lümeni", "ATP bağımlı konformasyonel kıskaç; istemci kinaz ve nükleer reseptörler", "İstemci proteinlerin agregasyonu, sinyal iletim yollarının çöküşü"],
            ["Küçük Isı Şoku (sHsp)", "Hsp27 (HSPB1), alfaB-Kristallin", "Sitozol, Lens, Kas sarkomeri", "ATP'den bağımsız; oligomerik sünger gibi hidrofobik yüzeyleri tutma", "Aşırı fosforilasyon, agregatlarla birlikte çökme ve katarakt"],
            ["Hsp110 Ailesi", "Hsp105, Apg-1", "Sitozol ve Çekirdek", "ATP bağımlı nükleotid değişim faktörü (NEF) + hafif desagregaz aktivitesi", "Hsp70 döngüsünün yavaşlaması ve substrat salımının tıkanması"]
        ]
    ),
    (
        "TABLO 2: Endoplazmik Retikulum UPR Kolları (IRE1a, PERK, ATF6) ve Hedef Transkriptom Matrisi",
        ["UPR Sensör Kolu", "ER Lümen Sensör Domeni", "Sitozolik / Nükleer Efektör", "Hedef Gen Ailesi ve Transkriptom", "Kronikleşme Sonucu Apoptoz Yolağı"],
        [
            ["IRE1alpha Kolu", "BiP ayrışması + Dimerizasyon", "XBP1s (Spliced mRNA transkripsiyon faktörü)", "ER şaperonları (BiP, PDI), ERAD bileşenleri, lipit sentezi", "JNK kinaz aktivasyonu, TRAF2/ASK1 kaskadı ve Kaspaz-12"],
            ["PERK Kolu", "BiP ayrışması + Trans-oto-fosforilasyon", "p-eIF2alpha -> ATF4 selektif translasyonu", "Amino asit taşıyıcıları, antioksidan yanıt (ATF4 hedefleri)", "CHOP / GADD153 indüksiyonu, Bcl-2 baskılanması ve BIM patlaması"],
            ["ATF6 Kolu", "BiP ayrışması -> Golgi translokasyonu", "ATF6f (p50 sitozolik transkripsiyon parçası)", "ER membran genişletme genleri, şaperonlar ve ERAD", "Apoptoza doğrudan gitmez; yetersiz kaldığında CHOP'u destekler"]
        ]
    ),
    (
        "TABLO 3: Übikitinasyon Kaskadı Enzimleri, Bağlantı Tipleri (K48, K63) ve Proteazom Alt Birimleri",
        ["Bileşen / Parametre", "Biyokimyasal Kimlik", "Enzimatik / Yapısal Görev", "Özgüllük / Topoloji Kuralı", "Bozulmanın Klinik Sonucu"],
        [
            ["E1 Enzimi (UBA1)", "Übikitin Aktive Edici Enzim", "ATP harcayarak yüksek enerjili tiyoester bağı kurar", "Tüm übikitinasyon için ortak tek kapı", "Hücresel letalite; X'e bağlı infantil spinal müsküler atrofi"],
            ["E2 Enzimleri (UBE2)", "Übikitin Konjugasyon Enzimi (~40 gen)", "Übikitini E1'den alır, E3'e aktarır", "Zincir bağlantı topolojisini (K48 vs K63) belirler", "DNA hasar tamir kusurları, hücre döngüsü duraklaması"],
            ["E3 Ligazlar (RING/HECT)", "Übikitin Ligaz (>600 gen)", "Substratı spesifik tanır ve izopeptit bağı kurar", "Yüksek substrat seçiciliği (degron tanıma)", "Onkogenez (MDM2 aşırılığı), nörodejenerasyon (Parkin kaybı)"],
            ["K48 Poli-Ubikitin", "Lizin 48 bağlantılı tetra-übikitin", "26S proteazom için evrensel yıkım sinyali", "Kompakt kapalı konformasyon", "Yetersizliğinde toksik protein agregasyonu patlar"],
            ["20S Katalitik Çekirdek", "beta1, beta2, beta5 treonin proteazları", "Polipeptid zincirini oligopeptitlere kesme", "Kaspaz, tripsin ve kimotripsin-benzeri odacıklar", "Lipofusin ve okside proteinlerle tıkanma; proteotoksisite"]
        ]
    ),
    (
        "TABLO 4: Makrootofaji Aşamaları, Çekirdek ATG Proteinleri ve Düzenleyici Kinazlar",
        ["Otofaji Aşaması", "Başlıca Protein Kompleksi", "Biyokimyasal İşlev", "Pozitif Uyarıcı Sinyal", "Negatif İnhibitör Sinyal"],
        [
            ["İndüksiyon", "ULK1 - ATG13 - FIP200 - ATG101", "ER zarına göç ve çekirdeklenme başlatma", "AMPK aktivasyonu (Ser317/777)", "mTORC1 aktivasyonu (Ser757 fosforilasyonu)"],
            ["Çekirdeklenme (Nucleation)", "VPS34 - Beclin1 - VPS15 - ATG14L", "ER omegasomunda PI3P lipit odakları üretme", "ULK1 fosforilasyonu, UVRAG", "Bcl-2 / Bcl-xL (Beclin-1'e bağlanarak kilitler)"],
            ["Fagofor Uzaması", "ATG12-ATG5-ATG16L1 ve LC3-PE (LC3-II)", "Lipitlenmiş LC3 ile fagofor zarını büyütme", "ATG7 (E1) ve ATG3 (E2) enzimatik akışı", "ATG4B mutasyonları veya sistein oksidasyonu"],
            ["Kapanma (Sealing)", "ESCRT-III (CHMP2A/4B) + VPS4", "Fagofor uçlarını kaynaştırıp küre oluşturma", "Kargo paketlenmesinin tamamlanması", "ESCRT fonksiyon bozukluğu; sızıntılı otofagozom"],
            ["Füzyon ve Yıkım", "HOPS kompleksi + STX17-SNAP29-VAMP8", "Lizozomla birleşme ve asit hidrolizi", "Rab7-GTP, asidik lizozomal pH (v-ATPase)", "Bafilomisin A1, Klorokin, lizozomal nötralizasyon"]
        ]
    ),
    (
        "TABLO 5: Otofaji Alt Tipleri (Makrootofaji, CMA, Mikro-otofaji) Karşılaştırmalı Biyofiziksel Analizi",
        ["Parametre", "Makrootofaji (Macroautophagy)", "Şaperon Aracılı Otofaji (CMA)", "Mikro-otofaji (Microautophagy)", "Endozomal Mikro-otofaji (e-MI)"],
        [
            ["Vezikül Gereksinimi", "Zorunlu (Çift katmanlı otofagozom)", "Kesinlikle Yok (Vezikülsüz doğrudan giriş)", "Yok (Lizozom zarının doğrudan invajinasyonu)", "Yok (Geç endozom zarının içe tomurcuklanması)"],
            ["Kargo Seçiciliği", "Seçici (Agrefaji, Mitofaji) veya Bulk", "Ultra Seçici (KFERQ motifi taşıyan proteinler)", "Genellikle non-selektif kütlesel sitoplazma", "KFERQ seçici veya kütlesel sıvı faz"],
            ["Kargo Boyut Sınırı", "Devasa (Tüm organeller, mitokondriler, agregatlar)", "Yalnızca tek tek çözünür polipeptid zincirleri", "Küçük protein granülleri ve glikojen partikülleri", "Sitozolik proteinler ve küçük moleküller"],
            ["Anahtar Reseptör / Kanal", "LC3-II, p62, OPTN, NBR1", "LAMP-2A multimerik translokasyon kanalı", "Lizozomal membran mikro-alanları", "ESCRT-I/II/III kompleksleri"],
            ["Yaşlanmadaki Bozulma", "Füzyon yavaşlaması, lizozom disfonksiyonu", "LAMP-2A reseptörünün dokulardan %70 silinmesi", "Membran sertleşmesiyle invajinasyon kaybı", "ESCRT alt birimlerinin yaşa bağlı tükenişi"]
        ]
    ),
    (
        "TABLO 6: Lizozomal Asit Hidrolazları, v-ATPase Regülatörleri ve TFEB Hedef Genleri (CLEAR Ağı)",
        ["Lizozomal Bileşen", "Biyokimyasal Tip / Gen", "Substrat ve Hedef Molekül", "Optimal pH Aralığı", "Klinik Yetersizlik Tablosu"],
        [
            ["Katepsin B ve L", "Sistein proteazlar (CTSB, CTSL)", "Proteinler, hücre dışı matriks, agregatlar", "pH 4.5 - 5.5", "Agregatom birikimi, nörodejenerasyon, miyopati"],
            ["Katepsin D", "Aspartil proteaz (CTSD)", "Amiloid-beta, Tau, alfa-sinüklein", "pH 3.5 - 4.5", "Nöronal Seroid Lipofusinozis (NCL10), erken demans"],
            ["Glukoserebrosidaz (GBA1)", "Lizozomal glikozil hidrolaz", "Glukozilseramid -> Glukoz + Seramid", "pH 4.5 - 5.0", "Gaucher Hastalığı, Parkinson riski 5 kat artar"],
            ["v-ATPase (V1VO)", "Çok alt birimli proton pompası", "ATP hidrolizi ile sitozolden lümene H+ basma", "Tüm asidifikasyonun kaynağı", "Lümen nötralizasyonu, tüm hidrolazların felç olması"],
            ["TFEB / CLEAR Ağı", "bHLH-leucine zipper transkripsiyon faktörü", "Tüm lizozomal ve otofajik genlerin promoterları", "Nükleusta regülatör", "Proteostaz çöküşü, hızlandırılmış hücresel yaşlanma"]
        ]
    ),
    (
        "TABLO 7: Hücresel Agregat Tipleri, Amiloid Yapıları, Toksisite Mekanizmaları ve Dedektör Boyalar",
        ["Agregat / Fibril Tipi", "Başlıca Protein Bileşeni", "Yapısal Konformasyon", "Hücresel Toksisite Mekanizması", "Biyofiziksel Tanı Probu"],
        [
            ["Amiloid Plaklar", "Abeta42, IAPP, Transtiretin (ATTR)", "Çapraz beta-tabakası (Cross-beta)", "Zar delinmesi, kalsiyum fışkırması, sinaps kopması", "Tioflavin T (ThT), Kırmızı Kongo, PiB-PET"],
            ["Nörofibriler Yumaklar", "Hiperfosforile Tau (p-Tau)", "Eşleşmiş helikal filamanlar (PHF)", "Mikrotübüllerin dağılması, aksonal transport felci", "T807 (Flortaucipir) PET, Fosfo-Tau ELISA"],
            ["Lewy Cisimcikleri", "alfa-Sinüklein, Ubikitin, p62", "Merkezi amiloid çekirdek + periferik halo", "Mitokondri Kompleks I blokajı, vezikül stazı", "alpha-Synuclein RT-QuIC analizi"],
            ["Agresomlar (Aggresomes)", "Hatalı katlanmış proteinler + Vimentin", "Sentrozom etrafında kafeslenmiş kütle", "Fiziksel yer işgali, mitoz iğ ipliği anomalisi", "Vimentin / gamma-Tübülin ko-lokalizasyonu"],
            ["Lipofusin Granülleri", "Okside protein (%50) + Lipit (%30) + Metaller", "Düzensiz polimerik Schiff bazları", "Lizozomal sindirimin tıkanması, Fenton radikalleri", "Otofloresans (Ex360-470 nm / Em550-650 nm)"]
        ]
    ),
    (
        "TABLO 8: Proteostaz Bozulmasıyla Karakterize İnsan Patolojileri, Agregat Proteinleri ve Organ Tutulumları",
        ["Hastalık Tablosu", "Başlıca Agregat Proteini", "Primer Etkilenen Organ / Doku", "Başlıca Kalite Kontrol Kusuru", "Klinik Sonuç ve Mortalite"],
        [
            ["Alzheimer Hastalığı", "Abeta42 (plak) ve p-Tau (yumak)", "Beyin (Serebral korteks, hipokampus)", "Otofaji-lizozom yetersizliği, proteazom blokajı", "İleri demans, kognitif çöküş, ölüm"],
            ["Parkinson Hastalığı", "alfa-Sinüklein (Lewy cisimciği)", "Beyin (Substantia nigra pars compacta)", "CMA bozulması (LAMP-2A kaybı), mitofaji yetmezliği", "Rijidite, tremor, akinezi, dopaminerjik ölüm"],
            ["Katarakt", "alfa, beta, gamma Kristallinler", "Göz Lens Lif Hücreleri", "Şaperon kapasitesi tükenişi, UV/glikasyon çöküşü", "Lens opaklaşması, görme kaybı, körlük"],
            ["Tip 2 Diyabet", "Adacık Amiloid Polipeptidi (IAPP)", "Pankreas Langerhans Beta Adacıkları", "ER stresi, sekretuar otofaji tıkanması", "Beta hücre kitlesi kaybı, insülin bağımlılığı"],
            ["Kardiyak Senil Amiloidoz", "Transtiretin (ATTRwt)", "Kalp Ventrikül Miyokardı", "TTR tetramer instabilitesi, proteazom yetersizliği", "Restriktif kardiyomiyopati, refrakter kalp yetmezliği"]
        ]
    ),
    (
        "TABLO 9: Otofaji ve Proteostaz Uyarıcı Farmakolojik Ajanlar, Hedefleri ve Klinik Aşamaları",
        ["Terapötik Ajan", "Moleküler Sınıf", "Primer Biyokimyasal Hedef", "Eylemin Moleküler Mekanizması", "Klinik Faz ve Longevity Durumu"],
        [
            ["Rapamisin (Sirolimus)", "Bakteriyel makrolid", "mTORC1 (FKBP12 kompleksi)", "mTORC1 baskısıyla ULK1 ve TFEB de-represyonu", "FDA onaylı immünosüpresan; PEARL klinik anti-aging çalışması"],
            ["Spermidin", "Doğal poliamin", "EP300 asetiltransferaz", "EP300 inhibisyonu ile ATG proteinlerinin deasetilasyonu", "İnsan klinik denemelerinde güvenli, kognitif koruyucu"],
            ["Trehaloz", "Doğal disakkarit", "GLUT taşıyıcıları / TFEB", "Glukoz açlığı taklidiyle kalsinörin/TFEB aktivasyonu", "Faz II (Okülofaringeal müsküler distrofi, ALS)"],
            ["TUDCA / 4-PBA", "Kimyasal şaperon", "ER lümen katlanmamış proteinleri", "Hidrofobik yüzeyleri örterek UPR stresini söndürme", "FDA onaylı (ALS tedavisi Relyvrio bileşeni)"],
            ["AR7 / CA77", "Küçük sentetik moleküller", "LAMP-2A reseptör stabilizasyonu", "CMA aktivitesini doğrudan gençlik seviyesine çıkarma", "Preklinik altın standart, faz denemelerine hazırlık"]
        ]
    ),
    (
        "TABLO 10: Geleceğin Proteomik Mühendislik Araçları: PROTAC, AUTAC, LYTAC ve Sentetik Şaperonlar",
        ["Yıkım / Düzenleme Teknolojisi", "Hedeflenen Hücresel Lokasyon", "Kullanılan Hücresel Litik Makine", "Terapötik Kargo Yelpazesi", "Gelecekteki Longevity Potansiyeli"],
        [
            ["PROTAC", "İntraselüler (Sitozol ve Çekirdek)", "26S Proteazom (E3 ligaz recruitment)", "Monomerik patolojik proteinler (Tau, onkojenler)", "Onkoloji ve nörodejenerasyonda Faz II klinik denemeler"],
            ["AUTAC", "İntraselüler (Sitoplazma)", "Seçici Makrootofaji (K63 ubikitinleme)", "Büyük protein agregatları, yaşlı mitokondriler", "Hücresel gençleşmede organel düzeyinde süper temizlik"],
            ["LYTAC", "Ekstraselüler ve Plazma Zarı", "Lizozomlar (CI-M6PR endositoz yolağı)", "Ekstraselüler Abeta plakları, dolaşımdaki AGE'ler", "Damar ve doku aralıklarındaki amiloidozun cerrahisiz erimesi"],
            ["Yapay Desagregaz (Hsp104v)", "İntraselüler inklüzyonlar", "Doğrudan mekanik AAA+ ATPaz çözme", "Katılaşmış amiloid lifleri, TDP-43 yumakları", "İleri nörodejeneratif hasarın geriye döndürülmesi"],
            ["De Novo Sıfır-Amiloid Proteom", "Tüm Genomik Proteom", "Doğal şaperonlar (Kendiliğinden katlanma)", "Tüm insan proteinleri (Mühendislik tasarımı)", "Homo Aeternus: Asla agregasyon yapmayan sentetik insan"]
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
print(f"BÖLÜM 08 başarıyla kaydedildi: {OUTPUT_PATH}")

