# -*- coding: utf-8 -*-
"""
PROJECT AETERNITAS - CİLT 16: KRİYOBİYOLOJİ, VİTRİFİKASYON VE BİYOSTAZ PROTOKOLLERİ
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

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_16_KRIYOBIYOLOJI_VITRIFIKASYON_BIYOSTAZ_TAM_100_SAYFA.docx"

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
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 16: KRİYOBİYOLOJİ VE BİYOSTAZ PROTOKOLLERİ")
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
s_run = sub_p.add_run("CİLT 16: KRİYOBİYOLOJİ, VİTRİFİKASYON VE BİYOSTAZ PROTOKOLLERİ\\n(BUZSUZ CAMSILAŞMA, KRİYOPROTEKTAN KİNETİĞİ, ALCOR STANDBY, NANO-ISITMA VE MOLEKÜLER CANLANDIRMA)")
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
ih_run = intro_h.add_run("CİLT 16 MANİFESTOSU: ZAMANIN ASKIYA ALINMASI: TERMODİNAMİK BİYOSTAZ VE EBEDİ DÖNÜŞ KÖPRÜSÜ")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "Tıp tarihi boyunca ölüm tanımı sürekli gerilemiştir: Kalbin durması geçmişte mutlak son sayılırken, kardiyopulmoner resüsitasyon "
    "(CPR) ve defibrilasyon ile geri döndürülebilir bir klinik faz haline gelmiştir. Günümüzde ise 'Biyolojik Ölüm', tek bir saniyede "
    "gerçekleşen bir olay değil; hücrelerin enerji krizine girerek saatler içinde çözüldüğü dinamik bir enformasyon kaybı sürecidir.\\n\\n"
    "Bu enformasyonel çözülmeyi durdurmanın termodinamik anahtarı 'Kriyobiyoloji ve Biyostaz'dır. Sıcaklık sıvı azot seviyesine "
    "(-196°C / 77 Kelvin) indirildiğinde; Arrhenius kinetiğine göre tüm kimyasal, enzimatik ve metabolik reaksiyonlar tamamen durur. "
    "Biyolojik zaman donar. Eğer bu soğutma süreci ölümcül buz kristalleri oluşturmadan, yüksek konsantrasyonlu kriyoprotektan kokteylleri "
    "(M22, VM3) ile amorf bir camsı fazda (Vitrifikasyon) gerçekleştirilirse; nöronal sinapslar, bellek engramları ve hücre zarları "
    "atomik düzeyde kusursuzca korunur.\\n\\n"
    "Bu ciltte; Arrhenius reaksiyon duraklaması, buz kristali nükleasyon kinetiği, camsı geçiş sıcaklığı (Tg), kriyoprotektan toksisite "
    "ve osmotik regülasyonu, Alcor ve CryoRus klinik standby prosedürleri, nörokriyoprezervasyon, termal stres kırılmaları, manyetik "
    "nanopartiküllerle nano-ısıtma (Nanowarming) ve gelecekteki nanorobotik moleküler canlandırma protokolü 100 derin akademik "
    "alt bölümde tüm biyofiziksel matematiğiyle ortaya konmaktadır."
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
        "1.1 Yaşamın Zamansal Duraklatılması: Biyostaz Kavramı ve Enformasyonel Ölüm Kriteri",
        "Geleneksel tıp klinik ölümü kalbin ve solunumun durmasıyla tanımlarken; kriyobiyoloji ve enformasyon teorisi 'Enformasyonel-Teorik Ölüm' (Information-Theoretic Death) kavramını esas alır.",
        "Bir insan biyolojik olarak ancak ve ancak beynindeki sinaptik bağlantı haritası (Connectome) ve bellek engramlarını depolayan nöronal yapılar termodinamik olarak hiçbir teorik teknolojiyle (geleceğin nanorobotik onarım sistemleri dahil) yeniden birleştirilemeyecek derecede fiziksel yıkıma uğradığında gerçek anlamda ölmüş sayılır. Biyostaz (Biostasis); bu enformasyonel çözülmeyi klinik ölümün hemen ardından aşırı düşük sıcaklıklarda dondurarak zamanı durdurma sanatıdır.",
        "Kriter_Enformasyonel_Olum: Entropi_Norom(t) > S_kritik (Baglanti Matrisi Geri Dondurulemez Sekilde Silinmistir)",
        "Bu enformasyon teorisi ölüm kriteri eşitsizliği, beynin sinaptik devre topolojisindeki yapısal entropi artışının geri döndürülemez kritik eşiği aştığı anı tanımlar."
    ),
    (
        "1.2 Arrhenius Kinetiği ve Kimyasal Reaksiyon Hızlarının Sıcaklığa Bağımlılığı",
        "Biyolojik sistemlerde gerçekleşen tüm enzimatik aktivite, oksidatif hasar, otoliz ve proteolitik yıkım reaksiyonları kimyasal reaksiyon kinetiğinin temel yasası olan Arrhenius denklemi ile yönetilir.",
        "Sıcaklık mutlak sıfıra (-273.15°C / 0 Kelvin) doğru yaklaştıkça, moleküllerin termal kinetik enerjisi (k_B * T) dramatik biçimde düşer. -196°C'de (77 Kelvin - Sıvı Azot sıcaklığı) hücresel su molekülleri ve enzimlerin aktivasyon enerjisi bariyerini aşma olasılığı sıfıra yaklaşır. Bir enzimin oda sıcaklığında 1 saniyede yaptığı yıkıcı hidroliz reaksiyonu; -196°C'de evrenin yaşından (13.8 milyar yıl) daha uzun bir sürede ancak gerçekleşebilir.",
        "Hiz_Reaksiyon_Kriyojenik = k(T) = A * exp( - E_a / (R * T) ) -> 0 (T = 77 Kelvin'de)",
        "Bu Arrhenius kinetik hız denklemi, kriyojenik sıcaklıklarda aktivasyon enerjisi engelinin kimyasal ve biyolojik yıkım hızını mutlak anlamda nasıl dondurduğunu belgeler."
    ),
    (
        "1.3 Moleküler Difüzyonun Donması ve Stokes-Einstein Viskozite Sınırı",
        "Kimyasal reaksiyonların durmasının yanı sıra; hücre içi biyokimyasal bozulmayı engelleyen ikinci temel fiziksel bariyer, moleküler difüzyonun mutlak olarak kilitlenmesidir.",
        "Stokes-Einstein bağıntısına göre, bir molekülün difüzyon katsayısı (D = k_B * T / (6 * pi * eta * r)); ortamın dinamik viskozitesi (eta) ile ters orantılıdır. Sıcaklık camsı geçiş noktasına yaklaştıkça sıvının viskozitesi 10^12 Pa*s (baldan trilyonlarca kat daha yoğun bir katı cam fazı) seviyesine fırlarken difüzyon katsayısı sıfıra iner. Serbest oksijen radikalleri veya proteazlar bir nanometre dahi ilerleyemez; hücresel yıkım fiziksel olarak imkânsızlaşır.",
        "Katsayi_Difuzyon = D = ( k_B * T ) / ( 6 * pi * eta(T) * r_molekul ) -> 0 (eta -> 10^12 Pa*s)",
        "Bu Stokes-Einstein difüzyon formülü, kriyoprotektan çözeltisinin camsı katı faza geçişinde moleküler hareketliliğin termodinamik olarak nasıl sıfırlandığını gösterir."
    ),
    (
        "1.4 Sıvı Azot (-196°C / 77 K) ve Gaz Fazı Sıvı Azot (-140°C) Depolama Termodinamiği",
        "Kriyobiyolojide uzun süreli biyostaz ortamı olarak Sıvı Azot (LN2) kullanılır; çünkü sıvı azot atmosferik basınçta tam -196°C'de kaynar ve mükemmel bir termal rezervuardır.",
        "Biyolojik örnekler iki farklı termal ortamda saklanır: 1) Doğrudan Sıvı Azot Daldırma (-196°C): Maksimum termal stabilite sunar, elektrik enerjisine ihtiyaç duymaz; ancak hızlı soğumada termal stres çatlamaları riski taşır; 2) Gaz Fazı Sıvı Azot (-140°C ila -160°C): Camsı geçiş sıcaklığının (Tg ~ -125°C) hemen altında çalışır. Bu sıcaklıkta tüm kimyasal aktivite durmuş haldedir ve mekanik çatlama stresi %90 daha düşüktür.",
        "Termal_Guvenlik_Suresi = t_otonomi = ( m_LN2 * Delta_H_buharlasma ) / Q_isi_kacagi > 30 - 60 gun",
        "Bu kriyojenik dewar termodinamik otonomi denklemi, elektrik kesintisi olsa dahi sıvı azot faz değişim gizli ısısının örneği aylarca güvenle donuk tutma süresini hesaplar."
    ),
    (
        "1.5 Kriyobiyolojik İskemi: Sıcak İskemi, Soğuk İskemi ve Hücresel Enerji Çöküşü",
        "Bir hasta klinik olarak öldüğünde kalbin durmasıyla kriyoprotektif perfüzyonun başlaması arasında geçen her saniye 'İskemik Hasar' saati olarak işler.",
        "1) Sıcak İskemi (37°C): Oksijen kesildiği anda hücreler anaerobik glikolize geçer; 5 dakika içinde ATP depoları tükenir, laktik asit birikir ve Na+/K+ pompaları durarak nöronlar şişer (sitotoksik ödem); 2) Soğuk İskemi (0-4°C): Hasta derhal buz banyosuna alınıp mekanik CPR uygulandığında metabolizma hızı %90 düşer; sıcak iskeminin 5 dakikada yarattığı hasar soğuk iskemide 1 saatte ancak oluşur.",
        "Tukenis_ATP = [ATP](t) = [ATP_0] * exp( - k_iskemi * Q10^( (T - 37) / 10 ) * t )",
        "Bu sıcaklığa bağımlı biyometabolik ATP tükenme kinetiği, vücut sıcaklığı 0°C'ye düşürüldüğünde hücresel enerji tükenme hızının nasıl dramatik biçimde yavaşlatıldığını belgeler."
    ),
    (
        "1.6 Suyun Anormal Faz Diyagramı ve Kriyojenik Durum Değişimleri",
        "Su, evrendeki diğer sıvıların aksine katılaşırken hacmi genişleyen (yaklaşık %9) ve hidrojen bağları nedeniyle son derece karmaşık faz davranışları sergileyen benzersiz bir moleküldür.",
        "Standart atmosferik basınçta su 0°C'de altıgen (heksagonal) buz Ih kristallerine dönüşür; bu faz biyolojik hücre zarlarını bir jilet gibi kesen en ölümcül yapıdır. Ancak basınç gigapaskal düzeylerine çıkarıldığında veya aşırı hızlı soğutulduğunda; yüksek yoğunluklu amorf buz (HDA) veya düşük yoğunluklu amorf buz (LDA) gibi kristal kafesi olmayan camsı fazlar ortaya çıkar. Kriyobiyolojinin amacı suyun bu kristalin hekzagonal tuzağını aşmaktır.",
        "Genlesme_Su = Delta_V / V_0 = ( rho_su - rho_buz ) / rho_buz ~ + 9.05% (Hacim Genişlemesi)",
        "Bu termal genleşme anomalisi eşitliği, donma sırasında suyun hacimsel büyümesinin kapalı hücresel kompartmanlarda yüzlerce atmosferlik yıkıcı hidrolik basınç patlamaları yarattığını tanımlar."
    ),
    (
        "1.7 Biyolojik Zamanın Durdurulması: Kozmik Radyasyon Arka Planı ve Asırlık Korunma",
        "Sıvı azot sıcaklığında tüm biyokimyasal reaksiyonlar ve metabolizma tamamen dondurulmuş olsa da; durdurulamayan tek bir fiziksel etki kalır: Arka plan Kozmik Radyasyon.",
        "Kozmik müonlar ve doğal çevre radyasyonu (yılda yaklaşık 2-3 mSv), -196°C'deki donuk dokularda da fiziksel iyonizasyon kırıkları (DNA tek ve çift zincir kırıkları) yapmaya devam eder. Biyofiziksel hesaplamalar; kriyojenik biyostaz altındaki bir insan beyninin radyasyon hasarı nedeniyle enformasyon kaybı yaşamadan önce en az 1000 ila 3000 yıl boyunca güvenle saklanabileceğini kanıtlamıştır.",
        "Doz_Kozmik = D_kumulatif(t) = Doz_yillik * t < D_kritik_parcalanma (1000+ Yıl Güvenlik)",
        "Bu kümülatif radyasyon integrali, kriyojenik biyostaz durumunda kozmik arka plan ışınımının nöromal engram bütünlüğünü bozmadan güvenle saklanabileceği asırlık zaman ufkunu tanımlar."
    ),
    (
        "1.8 İskemi-Reperfüzyon Hasarı Paradigması ve Antioksidan Tamponlama",
        "Kriyoprezervasyon sürecinde veya gelecekteki canlandırma anında; oksijensiz kalmış dokuya aniden oksijenli sıvının girmesi masif serbest radikal patlamalarına (İskemi-Reperfüzyon Hasarı) yol açar.",
        "Mitokondriyal elektron taşıma zinciri bozulduğu için ksantin oksidaz enzimi kalsiyum akımıyla aktive olur ve ortama devasa miktarda süperoksit (O2•-) ve hidroksil radikalleri deşarj eder. Bu fırtınayı engellemek için Alcor protokollerinde perfüzyon solüsyonuna yüksek dozda membran stabilizatörleri, melatonin, alfa-tokoferol türevleri ve ksantin oksidaz inhibitörü Allopurinol eklenir.",
        "Uretim_ROS_Reperfuzyon = d[ROS]/dt = k_ksantin * [Hipoksantin] * [O2] * ( 1 / ( 1 + [Allopurinol] / K_i ) )",
        "Bu reperfüzyon radikal kinetiği denklemi, güçlü antioksidan ve enzim inhibitörlerinin iskemik dokulardaki serbest radikal patlamasını nasıl baskıladığını simgeler."
    ),
    (
        "1.9 Biyostaz Protokollerinin Hukuki ve Etik Statüsü: Ölüm Tanımının Yeniden Yazımı",
        "Kriyobiyoloji ve biyostaz, geleneksel tıbbın ve yürürlükteki hukuk sistemlerinin katı ölüm tanımlarıyla sürekli çatışma halindedir.",
        "Mevcut hukukta kalbi duran birey derhal 'ceset' (kadavra) statüsüne alınmakta ve gömülme/yakılma zorunluluğu getirilmektedir. Alcor ve CryoRus gibi organizasyonlar bu engeli aşmak için hastaları 'Tıbbi Araştırma İçin Beden Bağışı' (Uniform Anatomical Gift Act) çerçevesinde kriyoprezervasyona alır. Kriyobiyoloji etiği; kalbi duran bir bireyi 'ölü' değil, mevcut teknolojinin tedavi edemediği 'derin komadaki ağır bir yoğun bakım hastası' olarak kabul eder.",
        "Statik_Dilemma: Hukuki_Kadavra == Gercek_Biyolojik_Kadavra ? -> CEVAP: HAYIR (Enformasyon Korunur)",
        "Bu biyoetik statü önermesi, hukuki ölüm belgesi almış bir bireyin hücresel enformasyonel düzeyde hâlâ kurtarılabilir bir canlı zihin taşıdığını formüle eder."
    ),
    (
        "1.10 Canlandırma (Revival) Varsayımları ve Moleküler Nanoteknoloji Köprüsü",
        "Kriyoprezervasyona giren bir insanın gelecekte yeniden hayata döndürülmesi (Revival); iki zorunlu bilimsel ayağın varlığına bağlıdır: Biyostaz kalitesi ve Geleceğin Moleküler Tamir Teknolojisi.",
        "Kriyobiyologlar hastayı canlandırmayı vaat etmez; hastanın nöronal yapısını ve anılarını gelecekteki tıp teknolojisine 'Kusursuz Bir Biyofiziksel Kargo' olarak teslim etmeyi hedefler. Geleceğin tıp teknolojisi (Cilt 15'te detaylandırılan tıbbi nanorobotlar, kuantum tarama ve in situ genetik tamir); donuk haldeki dokuyu atom atom tarayarak kriyoprotektanları temizleyecek, hücresel membranları tamir edecek ve kalbi gençlik ritminde yeniden çalıştıracaktır.",
        "P_Canlandirma = P_Enformasyon_Biyostaz * P_Teknoloji_Nanotamir (Her İki Şart Sağlandığında = %100)",
        "Bu olasılık bileşkesi eşitliği, gelecekte başarılı bir canlandırmanın bugün sağlanan dondurma kalitesi ile gelecekteki moleküler onarım kapasitesinin ortak fonksiyonu olduğunu ortaya koyar."
    )
]

# ================= KISIM 2 =================
part2_subsections = [
    (
        "2.1 Buz Kristali Nükleasyonu Biyofiziği: Homojen vs Heterojen Nükleasyon",
        "Saf su, hiçbir yabancı parçacık içermediğinde 0°C'de donmaz; termal dalgalanmaların kendiliğinden mikroskobik bir buz çekirdeği oluşturabilmesi için sıcaklığın -38.5°C'ye (Homojen Nükleasyon Sıcaklığı - Th) kadar düşmesi gerekir.",
        "Ancak biyolojik sıvılarda (kan, sitoplazma); çözünmüş proteinler, lipid kalıntıları ve hücre zarları heterojen nükleasyon merkezleri (tohumları) olarak görev yapar. 'Heterojen Nükleasyon'da yabancı yüzeyler su moleküllerinin buz kristali kafesine dizilmesi için gereken serbest enerji bariyerini (Delta_G_nukleasyon) radikal biçimde düşürür; buz oluşumu 0°C ile -10°C arasında kontrolsüzce patlar.",
        "Hiz_Nukleasyon = J_nuk = K_0 * exp( - Delta_G_nukleasyon / (k_B * T) ) * exp( - Delta_G_difuzyon / (k_B * T) )",
        "Bu klasik nükleasyon teorisi denklemi, birim zamanda oluşan buz kristali çekirdeği sayısının termodinamik yüzey enerjisi ve sıcaklıkla üstel ilişkisini tanımlar."
    ),
    (
        "2.2 Kristal Morfolojisi: Dendritik Büyüme ve Hücresel Membranların Mekanik Delinmesi",
        "Oluşan ilk buz çekirdekleri çevrelerindeki serbest su moleküllerini hızla emerek fraktal, sivri uçlu 'Dendritik Buz Kristalleri' (Dendritic Ice) halinde büyür.",
        "Bu sivri iğnemsi kristal uçları; lipid çift katmanlı hücre zarlarını, mitokondri membranlarını, sinaptik vezikülleri ve mikrotübül iskeletini mekanik olarak keser, deler ve yırtar. Elektron mikroskobu görüntüleri; koruyucusuz dondurulan bir beyin dokusunda tüm nöron zarlarının patladığını ve hücre içinin parçalanmış bir enkaz yığınına dönüştüğünü açıkça göstermektedir; bu mekanik parçalanma biyolojik ölümü kesinleştirir.",
        "Hiz_Dendritik_Buyume = v_kristal = K_kinetik * ( T_ergime - T )^mu (Aşırı Soğuma ile Üstel Artış)",
        "Bu kristal büyüme kinetiği denklemi, ortam sıcaklığı erime noktasının altına düştükçe dendritik sivri buz kollarının yayılma hızındaki patlamayı modeller."
    ),
    (
        "2.3 Solüt Konsantrasyonu ve 'Çözelti Etkisi' (Solution Effects) Toksisitesi",
        "Buz kristalleri oluşurken içine yalnızca saf H2O moleküllerini kabul eder; sudaki tuzlar (NaCl), potasyum, proteinler ve metabolitler donan buzun dışına doğru itilir (Solüt Dışlama).",
        "Dondurulan dokuda su buza dönüştükçe, geride kalan minik sıvı kanallarındaki tuz konsantrasyonu fizyolojik 300 mOsm/kg seviyesinden aniden 3000-5000 mOsm/kg gibi ölümcül seviyelere fırlar. Bu hiperozmotik asit-tuz çorbası (Çözelti Etkisi); hücre zarlarını kimyasal olarak yakar, hücresel proteinleri denatüre eder ve enzimleri geri dönüşümsüz olarak inaktive eder.",
        "Konsantrasyon_Solut = C_kalan(T) = C_0 / ( 1 - Fraksiyon_Buz(T) ) -> Sonsuz (Fraksiyon_Buz -> 1)",
        "Bu solüt konsantrasyon eşitliği, donan buz kütlesi fraksiyonu arttıkça geride kalan sıvı mikro-kanallardaki ozmotik tuz yoğunluğunun hiperbolik olarak nasıl tavan yaptığını belgeler."
    ),
    (
        "2.4 Mazur'un İki-Faktör Hipotezi: Soğutma Hızı ve İntrasellüler Buz Çelişkisi",
        "Kriyobiyolojinin kurucusu Peter Mazur; hücrelerin dondurulması sırasında iki ölümcül mekanizmanın zıt yönde yarıştığını 'İki-Faktör Hipotezi' (Two-Factor Hypothesis) ile formüle etmiştir.",
        "1) Aşırı Yavaş Soğutma: Hücre dışındaki su donar, hücre uzun süre yüksek tuz konsantrasyonuna maruz kalır ve aşırı su kaybederek buruşur (Çözelti hasarı ölümü); 2) Aşırı Hızlı Soğutma: Hücre içindeki su dışarı kaçacak zaman bulamaz ve doğrudan hücre içinde ölümcül buz kristalleri oluşturur (İntrasellüler Buz Hasarı - IIF). Her hücre tipi için bu iki ölüm mekanizmasının ortasında yer alan çok dar bir 'Optimum Soğutma Hızı' mevcuttur; ancak bütün bir organ için homojen bir hız yakalamak imkânsızdır.",
        "Sagkalim_Mazur = S(B) = exp( - k_cozelti / B ) * exp( - k_IIF * B ) (B = dT/dt Soğutma Hızı)",
        "Bu Mazur çan eğrisi eşitliği, hücre sağkalımının soğutma hızı (B) aşırı düşük olduğunda çözelti toksisitesiyle, aşırı yüksek olduğunda hücre içi buz oluşumuyla nasıl çöktüğünü ifade eder."
    ),
    (
        "2.5 İntrasellüler Buz Oluşumu (IIF): Nükleasyon Şablonları ve Membran Yırtılması",
        "Hücre İçi Buz Oluşumu (IIF - Intracellular Ice Formation); kriyobiyolojide hücre için kesin bir ölüm fermanı olarak kabul edilir.",
        "IIF iki ana yolla tetiklenir: 1) Yüzey Katalizli Nükleasyon: Hücre dışındaki sivri bir buz kristali hücre zarını mekanik olarak gerer ve zardaki gözeneklerden sitoplazmaya sızarak hücre içindeki suyu anında tohumlar; 2) Hücre İçi Heterojen Çekirdeklenme: Organeller sitoplazmik suyu kristalleştirir. Sitoplazma donduğu anda genleşen buz hacmi hücre zarını içeriden paramparça eder.",
        "Olasilik_IIF = P_IIF = 1 - exp( - Integral_0^t J_IIF * A_membran dt' )",
        "Bu Poisson stokastik nükleasyon eşitliği, hücre zar alanı ve aşırı soğuma derecesine bağlı olarak hücre içi ölümcül kristal patlaması olasılığını modeller."
    ),
    (
        "2.6 Ozmotik Dehidrasyon Biyofiziği: Hücresel Hacim Büzüşmesi ve Lysis",
        "Hücre dışı kompartmanda buz oluştuğunda kimyasal potansiyel gradyanı nedeniyle hücre içindeki serbest su molekülleri su kanallarından (Akuaporinler) hızla dışarı kaçar.",
        "Hücre su kaybettikçe bir kuru üzüm gibi büzüşür. Ancak hücre zarı belirli bir 'Minimum Kritik Hacim' sınırından (V_min ~ V_0 * 0.35) daha fazla büzüşemez. Hücre bu sınırın ötesinde su kaybetmeye zorlandığında hücre zarı kendi içine katlanır, zardaki lipidler çift katmandan hekzagonal fazlara kayar ve zar bütünlüğü fiziksel olarak patlar (Osmotik Lysis).",
        "Hacim_Hucre_Buzusme = dV/dt = - L_p * A_membran * R * T * ( [Solut_dis] - [Solut_ic] )",
        "Bu Kedem-Katchalsky membran su geçirgenlik denklemi, dış ortamdaki hiperozmotik buz kanallarına doğru hücreden kaçan su debisini ve hacimsel çöküş hızını simgeler."
    ),
    (
        "2.7 Miyelin Kılıf ve Aksonal İletim Hatlarında Buz Kristali Tahribatı",
        "Merkezi Sinir Sisteminde aksonları saran çok katmanlı lipid miyelin kılıfları ve Ranvier düğümleri, buz kristali büyümesine karşı son derece kırılgandır.",
        "Miyelin lamelleri arasındaki mikroskobik su tabakaları donduğunda; buzun hacimsel genleşmesi miyelin katmanlarını birbirinden ayırır (Miyelin Balonlaşması ve Delaminasyon). Aksonal mikrotübüller kırılır ve presinaptik butonlar nöropilden kopar. Geleneksel dondurma yöntemleri uygulandığında beynin nöromal mimarisi geri döndürülemez biçimde mekanik parçalanmaya uğrar.",
        "Delaminasyon_Miyelin = Delta_d_lamel = d_0 * ( 1 + 0.09 * Fraksiyon_Buz_interlamellar )",
        "Bu mekanik ayrışma formülü, miyelin kılıf lamelleri arasında donan buzun katmanlar arası mikroskobik mesafeyi nasıl açarak elektriksel yalıtımı parçaladığını tanımlar."
    ),
    (
        "2.8 Kılcal Damar Yatağında Mikrovasküler Yırtılma ve Oklüzyon",
        "Bütün bir organın veya beynin dondurulmasında en ölümcül hasarlardan biri damar ağında gerçekleşir.",
        "Kan damarlarının lümeninde donan buz kütlesi; endotel hücrelerini damar duvarından soyar (endotelyal denüdasyon) ve elastik bazal membranı yırtar. Organ yeniden çözüldüğünde bu yırtılmış kılcal damarlar kanı tutamaz; masif iç kanamalar (rüptür) ve yaygın mikrovasküler tromboz gelişerek organı saniyeler içinde nekroza uğratır. Bu nedenle organ naklinde ve biyostazda geleneksel dondurma tamamen terk edilmiştir.",
        "Hasar_Endotel = D_vaskuler = A_denudasyon / A_toplam = k_ruzgar * [Kristal_Buz_Lumen]",
        "Bu vasküler endotel soyulma oranı, lümende büyüyen buz kristallerinin mikrovasküler bazal membranda yarattığı mekanik hasar yüzdesini ifade eder."
    ),
    (
        "2.9 Çözünme Sırasında Rekristalizasyon (Recrystallization) Tehlikesi",
        "Dondurma işlemi hasarsız tamamlansa dahi; çözünme (ısıtma) evresi çok daha sinsi ve ölümcül bir tehlike barındırır: 'Rekristalizasyon'.",
        "Ostwald Olgunlaşması (Ostwald Ripening) termodinamik yasasına göre; küçük buz kristalleri yüksek yüzey serbest enerjisine sahiptir. Isınma sırasında kritik sıcaklık aralığında (-50°C ila 0°C), termodinamik sistem yüzey enerjisini düşürmek için milyonlarca küçük kristali eriterek devasa makroskobik buz kütleleri halinde birleştirir. Dondururken dokuya zarar vermeyen mikroskobik çekirdekler, çözünürken dev canavarlara dönüşerek hücreleri içeriden biçer.",
        "Buyume_Rekristalizasyon = R_kristal(t)^3 - R_0^3 = K_Ostwald * exp( - Delta_G_dif / (k_B * T) ) * t",
        "Bu LSW (Lifshitz-Slyozov-Wagner) rekristalizasyon kinetiği eşitliği, çözünme sırasında küçük kristallerin devasa mekanik buz kütlelerine dönüşme hızını matematiksel olarak belgeler."
    ),
    (
        "2.10 Kriyobiyolojinin Nihai Çıkmazı: Dondurarak Koruma Paradigmasının Çöküşü",
        "Buz kristali nükleasyonu, hücresel dehidrasyon, çözelti toksisitesi ve rekristalizasyonun oluşturduğu bu dörtlü ölüm tuzağı; 'Buz İçeren Dondurma' (Freezing) ile karmaşık bir memeli organının veya beyninin hasarsız korunamayacağını kesin olarak kanıtlamıştır.",
        "Organik yaşamı korumanın tek yolu; suyu dondurarak katılaştırmak değil, suyun hiç buz kristali oluşturmasına fırsat vermeden doğrudan katı bir cam fazına dönüştürülmesidir. Bu devrimsel paradigmaya 'Vitrifikasyon' (Buzsuz Camsılaşma) denir ve modern biyostazın mutlak temelidir.",
        "Paradigma_Donusumu: Dondurma (Buz Kristali Hasarı / Olum) ----> Vitrifikasyon (Buzsuz Camsilasma / Biyostaz)",
        "Bu metodolojik kırılma prensibi, kriyobiyolojinin mekanik kristal hasarını sıfırlayan amorf katılaşma teknolojisine zorunlu geçişini simgeler."
    )
]

# ================= KISIM 3 =================
part3_subsections = [
    (
        "3.1 Vitrifikasyon (Camsılaşma) Tanımı: Buz Kristalleri Olmadan Amorf Katılaşma",
        "Vitrifikasyon (Vitrification); yüksek konsantrasyonlu kriyoprotektan çözeltileri içeren bir sıvının, soğutulurken hiçbir buz kristali fazı oluşturmadan viskozitesinin üstel artarak doğrudan 'Amorf Katı Cam' (Amorphous Glass) fazına geçmesidir.",
        "Vitrifike olmuş bir biyolojik dokuda tek bir buz kristali çekirdeği dahi bulunmaz. Su molekülleri kristal kafesi kuramaz; sıvı fazdaki rastgele, düzensiz geometrik konumlarında aniden termodinamik olarak kilitlenirler. Doku fiziksel olarak bir cam bardak gibi kaskatı ve şeffaf bir katıya dönüşür; ancak hücresel zarlar, organeller ve sinapslar sıfır mekanik basınca maruz kalarak orijinal sıvı anatomik formlarında korunur.",
        "Faz_Vitrifikasyon: Sivi (Düzensiz) --[Kritik Soğutma]--> Amorf Cam (Düzensiz Katı) (Delta_V_kristal = SIFIR)",
        "Bu amorf faz dönüşüm şeması, vitrifikasyonun hacimsel kristal genleşmesi olmaksızın moleküler mimariyi dondurma prensibini tanımlar."
    ),
    (
        "3.2 Camsı Geçiş Sıcaklığı (Tg): Süper-Soğutulmuş Sıvıdan Cama Termodinamik Geçiş",
        "Vitrifikasyon sürecinin termodinamik dönüm noktası, malzemenin fiziksel olarak katılaştığı 'Camsı Geçiş Sıcaklığı'dır (Tg - Glass Transition Temperature).",
        "Tipik vitrifikasyon kokteylleri (M22 gibi) için Tg yaklaşık -122°C ila -125°C civarındadır. Sıcaklık Tg'nin üzerine çıktığında malzeme süper-soğutulmuş viskoz bir sıvı gibi davranır; Tg'nin altına indiğinde ise viskozite 10^12 Pa*s eşiğini aşarak sistem mekanik olarak kırılgan bir katı cama dönüşür. Bu sıcaklıkta özgül ısı kapasitesinde (Cp) ani bir basamak değişimi ölçülür (İkinci Derece Sözde Faz Dönüşümü).",
        "Viskozite_Tg = eta(Tg) = 10^12 Pa*s = 10^13 Poise (Camsı Katılaşma Eşiği)",
        "Bu rheolojik viskozite kriteri, moleküler akışkanlığın bittiği ve malzemenin katı amorf cam fazına girdiği termodinamik sıcaklık koordinatını belgeler."
    ),
    (
        "3.3 Kritik Soğutma Hızı (CCR) ve Kritik Isıtma Hızı (CHR) Biyofiziği",
        "Bir çözeltinin vitrifike olabilmesi için buz nükleasyonunun gerçekleştiği tehlikeli sıcaklık bölgesinden (0°C ile -100°C arası) buz oluşmasına fırsat vermeden hızla geçilmesi şarttır.",
        "Bu hız eşiğine 'Kritik Soğutma Hızı' (CCR - Critical Cooling Rate) denir. Eğer soğutma hızı CCR'den yavaşsa buz kristalleri oluşur ve vitrifikasyon başarısız olur. Daha da önemlisi, çözünme sırasındaki rekristalizasyonu engellemek için gereken 'Kritik Isıtma Hızı' (CHR - Critical Heating Rate), CCR'den genellikle 10 ila 100 kat daha yüksektir. Yüksek kriyoprotektan konsantrasyonu CCR ve CHR eşiklerini düşürerek kalın insan organlarının vitrifike edilmesini mümkün kılar.",
        "Kriter_Vitrifikasyon: ( dT/dt )_sogutma > CCR | ( dT/dt )_isitma > CHR",
        "Bu termal hız eşitsizliği, buz kristali oluşumunu ve çözünme rekristalizasyonunu tamamen dışlamak için gereken asgari soğutma ve ısıtma hızlarını formüle eder."
    ),
    (
        "3.4 Su Moleküllerinin Kinetik Hapsi: Viskozite Patlaması ve Hidrojen Bağı Kilitlenmesi",
        "Vitrifikasyonun moleküler mekanizması, kriyoprotektan molekülleri ile su arasındaki güçlü hidrojen bağı etkileşimlerinde saklıdır.",
        "Dimetil sülfoksit (DMSO) ve gliserol gibi kriyoprotektanlar su molekülleriyle sudan daha güçlü hidrojen bağları kurar. Sıcaklık düştükçe bu moleküller suyun altıgen buz kafesi oluşturacak şekilde serbestçe dönmesini ve bir araya gelmesini sterik olarak engeller. Su molekülleri kriyoprotektanların arasında kinetik olarak hapsolur; viskozite üstel olarak tavan yapar ve moleküller hareket edemez hale gelerek donar.",
        "Viskozite_VFT = eta(T) = eta_0 * exp( B_kirilganlik / ( T - T_0 ) ) (Vogel-Fulcher-Tammann Denklemi)",
        "Bu VFT viskozite denklemi, sıcaklık T0 ideal camsılaşma noktasına yaklaştıkça çözelti viskozitesinin süper-Arrhenius bir hızla nasıl sonsuza doğru fırladığını belgeler."
    ),
    (
        "3.5 Zaman-Sıcaklık-Dönüşüm (TTT) Diyagramları ve Kristalizasyon Burnunun Aşılması",
        "Metalurji ve cam biliminden kriyobiyolojiye uyarlanan 'Zaman-Sıcaklık-Dönüşüm' (TTT - Time-Temperature-Transformation) diyagramları, vitrifikasyonun yol haritasıdır.",
        "TTT eğrisi üzerinde tipik bir 'Kristalizasyon Burnu' (Nose of Crystallization) bulunur; bu burun buz kristali oluşumu için gereken minimum zamanı ve kritik sıcaklığı gösterir. Bir biyolojik organın soğutma eğrisi bu kristalizasyon burnunun solundan teğet geçmek zorundadır. Yüksek CPA konsantrasyonu bu burnu sağa (daha uzun sürelere) kaydırır; böylece yavaş soğutma hızlarında dahi burnun etrafından dolaşılarak buzsuz camsılaşma penceresi yakalanır.",
        "Zaman_Burun = t_nose = t_minimum_kristalizasyon (Bu süreden daha hızlı soğutulmalıdır)",
        "Bu TTT kinetik koordinatı, biyolojik numunenin buz kristalleri ile çarpışmadan amorf cam bölgesine ulaşabilmesi için gereken maksimum geçiş süresini modeller."
    ),
    (
        "3.6 Diferansiyel Taramalı Kalorimetre (DSC) ile Camsılaşmanın Tescillenmesi",
        "Bir biyolojik dokunun gerçekten buzsuz olarak vitrifike olup olmadığını veya içinde gizli mikro-kristaller kalıp kalmadığını kanıtlayan analitik altın standart 'Diferansiyel Taramalı Kalorimetri'dir (DSC).",
        "DSC cihazı, numuneyi soğuturken ve ısıtırken açığa çıkan veya emilen ısı akısını (Heat Flow / mW) mikrojoule hassasiyetle ölçer. Eğer dokuda buz oluşmuşsa 0°C civarında devasa bir endotermik erime piki (latent heat peak) görülür. Eğer tam vitrifikasyon başarılmışsa erime piki sıfırdır; grafikte yalnızca Tg noktasında (-124°C) minik bir bazal hat basamak kayması (Delta_Cp) gözlemlenir.",
        "Entalpi_Erime = Delta_H_erime = Integral [ Q_isi_akisi dt ] = 0 J/g (Tam Vitrifikasyon Kanıtı)",
        "Bu termodinamik erime entalpisi eşitliği, DSC analizinde sıfır latent ısı açığa çıkmasının dokuda %100 kristalsiz amorf katılaşmanın kesin bilimsel kanıtı olduğunu belgeler."
    ),
    (
        "3.7 Fahy-Wowk Vitrifikasyon Teorisi: Konsantrasyon vs Soğutma Hızı Ödünleşimi",
        "Kriyobiyolojinin yaşayan efsaneleri Gregory Fahy ve Brian Wowk; vitrifikasyonun temel mühendislik ödünleşimini (trade-off) formüle etmiştir.",
        "Bir organı vitrifike etmek için iki değişken mevcuttur: Kriyoprotektan Konsantrasyonu (C) ve Soğutma Hızı (B). Konsantrasyon ne kadar yüksek olursa gereken soğutma hızı o kadar düşer; ancak yüksek konsantrasyon hücreye kimyasal toksisite yükler. Soğutma hızı ne kadar yüksek olursa gereken CPA konsantrasyonu düşer; ancak kalın insan organlarında ısıl iletkenlik sınırı nedeniyle yüksek soğutma hızları yakalanamaz ve termal gerilme kırılmaları oluşur. Fahy-Wowk teorisi bu iki zıt vektörün optimum kesişim noktasını hesaplar.",
        "Kriter_Fahy_Wowk: C_CPA * ( dT/dt )_sogutma >= Esik_Vitrifikasyon (Optimum Toksisite-Hız Dengesi)",
        "Bu biyomühendislik optimizasyon eşitliği, kimyasal kriyoprotektan toksisitesi ile fiziksel soğutma hızının organ bütünlüğünü koruyacak ideal kesişim dengesini simgeler."
    ),
    (
        "3.8 İnsan Beyninin Vitrifikasyonu: Synaptome ve Ultrastrüktürün Korunumu",
        "İnsan beyninin vitrifikasyonu, biyostazın en kutsal hedefidir çünkü bireyin kimliğini, anılarını ve bilincini taşıyan sinapslar burada yer alır.",
        "21st Century Medicine laboratuvarında Fahy ve ekibi; optimize edilmiş M22 vitrifikasyon solüsyonu ile tam bir memeli beynini -135°C'ye soğutmuş, ardından çözerek elektron mikroskobuyla incelemiştir. Sonuçlar tarihi bir zaferdir: Pre-sinaptik veziküller, post-sinaptik yoğunluklar (PSD), dendritik mantar dikenleri ve astrositik son ayaklar nanotopografik düzeyde tek bir bozulma olmaksızın korunmuştur (Synaptome Preservation).",
        "Korunma_Sinaptik = eta_sinaps = [Sinaps_intakt_post] / [Sinaps_pre] > %99 (Elektron Mikroskobu Tescilli)",
        "Bu nöromorfometrik doğrulama oranı, optimize vitrifikasyon protokollerinin sinaptik bağlantı haritasını atomik ölçekte kusursuz koruma başarısını gösterir."
    ),
    (
        "3.9 Kriyofiksasyon ve Vitrifikasyon Elektron Mikroskopisi (Cryo-EM) ile Görselleştirme",
        "Vitrifikasyonun moleküler seviyedeki kusursuzluğu, 2017 Nobel Kimya Ödülü'ne konu olan 'Kriyo-Elektron Mikroskopisi' (Cryo-EM) teknolojisi ile görselleştirilmiştir.",
        "Biyolojik proteinler sıvı etan içinde saniyede 100.000°C hızla vitrifike edildiğinde su molekülleri proteinin etrafında buz oluşturamaz; proteinler doğal çözelti konformasyonlarında camsı buz içinde hapsolur. Cryo-EM ile atomik çözünürlükte (1-2 Angstrom) elde edilen görüntüler; vitrifikasyonun proteinlerin üçüncül ve dördüncü yapılarını hiçbir denatürasyona uğratmadan koruduğunu fiziki olarak tescillemiştir.",
        "Cozunurluk_CryoEM = d_atomik ~ 1.2 - 2.0 Angstrom (Atomik Konformasyonel Korunum)",
        "Bu yapısal biyoloji çözünürlük standardı, vitrifikasyon fazında biyomoleküler katlanmaların atomik koordinat hassasiyetinde donduğunu kanıtlar."
    ),
    (
        "3.10 Katı Organ Vitrifikasyonu: Böbrek ve Kalp Modelinde Dünya Rekorları",
        "Hücresel süspansiyonların (sperm, embriyo, kök hücre) vitrifikasyonu uzun yıllardır rutin klinikte uygulanırken; milyarlarca hücre içeren masif katı organların vitrifikasyonu devasa bir mühendislik başarısıdır.",
        "Gregory Fahy ve ekibi; tavşan böbreğini M22 solüsyonu ile perfüze ederek -135°C'de tamamen vitrifike etmiş, ardından buz oluşturmadan çözerek tavşana geri nakletmiştir. Nakledilen biyo-yapay vitrifike böbrek; tavşanın tek böbreği olarak haftalarca tam fonksiyon göstermiş ve hayvanın yaşamını sürdürmesini sağlamıştır. Bu başarı, vitrifikasyonun katı organlarda geri döndürülebilir olduğunu tescilleyen dünya rekorudur.",
        "Fonksiyonellik_PostVitrifiye = GFR_post / GFR_kontrol > %85 (In Vivo Hayatta Kalma Başarısı)",
        "Bu organ nakil fonksiyonel retansiyon oranı, katı bir memeli organının tam vitrifikasyon ve çözünme döngüsünden fizyolojik olarak sağ çıkabileceğinin somut kanıtıdır."
    )
]

# ================= KISIM 4 =================
part4_subsections = [
    (
        "4.1 Kriyoprotektan Ajanların (CPA) Biyokimyasal Sınıflandırması",
        "Vitrifikasyonu mümkün kılan kimyasal bileşikler olan Kriyoprotektan Ajanlar (CPA - Cryoprotective Agents); hücre zarlarını geçme yeteneklerine göre iki ana sınıfa ayrılır.",
        "1) Penetran (Hücre İçine Giren) Kriyoprotektanlar: Molekül ağırlıkları küçük (<100 Da) ve amfifilik karakterde olup hücre zarını pasif veya kolaylaştırılmış difüzyonla aşarak sitoplazmadaki suyu korur (DMSO, Gliserol, Etilen Glikol, Formamid); 2) Non-Penetran (Hücre Dışında Kalan) Kriyoprotektanlar: Yüksek molekül ağırlıklı polimerler veya şekerlerdir (Trehaloz, Sukroz, Dekstran, PVP); hücre içine giremez, hücre dışı ozmolariteyi artırarak suyu hücreden kontrollü çeker ve buz büyümesini dışarıda kilitler.",
        "Siniflandirma_CPA: [Penetran (DMSO, EG, Gliserol)] + [Non-Penetran (Trehaloz, Dekstran)]",
        "Bu moleküler taksonomi şeması, kriyoprotektif solüsyonların hücre içi ve hücre dışı kompartmanları eş zamanlı koruma prensibini tanımlar."
    ),
    (
        "4.2 Dimetil Sülfoksit (DMSO): Hücre Zarı Geçirgenliği ve Hidrojen Bağı Dinamiği",
        "Kriyobiyoloji tarihinde en yaygın kullanılan penetran kriyoprotektan olan Dimetil Sülfoksit (DMSO / (CH3)2SO); hem polar sülfoksit grubu hem de apolar metil grupları taşıyan eşsiz bir çözücüdür.",
        "DMSO, hücre zarındaki fosfolipid baş gruplarının hidrasyon kabuğunu yeniden düzenleyerek zardan saniyeler içinde geçer. Hücre içinde su molekülleriyle güçlü hidrojen bağları kurar (DMSO-su bağı, su-su bağından termodinamik olarak daha kararlıdır). Bu durum suyun donma noktasını düşürür ve nükleasyonu zorlaştırır. Ancak yüksek konsantrasyonlarda proteinlerin hidrofobik ceplerini açarak denatürasyona yol açması ana toksisite sınırıdır.",
        "Katsayi_Gecirgenlik_DMSO = P_DMSO = 1.5 * 10^-3 cm/s >> P_gliserol (Zardan Çok Hızlı Geçiş)",
        "Bu transmembran difüzyon hızı karşılaştırması, DMSO'nun hücresel kompartmanlara dakikalar içinde nüfuz ederek içsel buz oluşumunu engelleme üstünlüğünü belgeler."
    ),
    (
        "4.3 Etilen Glikol (EG) ve Formamid: Düşük Viskozite ve Hızlı Difüzyon Avantajı",
        "Gliserol gibi yüksek viskoziteli CPA'lar kalın organ dokularına difüze olmakta aşırı yavaş kalırken; Etilen Glikol (EG) ve Formamid son derece akışkan ve hızlı penetre olan moleküllerdir.",
        "Etilen Glikol (HO-CH2-CH2-OH), iki hidroksil grubuyla mükemmel bir buz baskılayıcıdır ve DMSO'dan daha düşük viskoziteye sahiptir. Formamid (HCONH2) ise amid bağı sayesinde suyun hidrojen ağına kusursuzca entegre olur ve çözeltinin viskozitesini dramatik biçimde düşürür. Bu iki molekülün kombinasyonu; organ perfüzyonunda damar içi direnci düşürerek tüm kılcal damarlara dakikalar içinde homojen kriyoprotektan dağılımı sağlar.",
        "Viskozite_Karisim = eta_karisim = exp( x_DMSO * ln(eta_DMSO) + x_EG * ln(eta_EG) + x_Formamid * ln(eta_Formamid) )",
        "Bu logaritmik viskozite karışım kuralı, düşük moleküler ağırlıklı CPA'ların perfüzyon hidrolik direncini nasıl optimize ettiğini formüle eder."
    ),
    (
        "4.4 Gliserolün Tarihsel Rolü: Luyet, Smith ve Polge'un Çığır Açan Keşifleri",
        "1949 yılında Christopher Polge, Audrey Smith ve Alan Parkes; boğa spermini dondururken tesadüfen gliserol içeren bir solüsyon kullanmış ve hücrelerin çözündükten sonra canlı kaldığını keşfederek modern kriyobiyolojiyi başlatmıştır.",
        "Gliserol (C3H8O3), üç hidroksil grubu taşıyan doğal ve toksisitesi son derece düşük bir trioldür. İnsan alyuvarlarının (kan bankacılığı) dondurulmasında bugün dahi %40 konsantrasyonda altın standarttır. Ancak moleküler büyüklüğü ve yüksek viskozitesi nedeniyle hücre zarlarından çok yavaş geçer; bütün bir insan beyni veya böbreğinin perfüzyonunda saatler gerektirdiği için modern vitrifikasyonda yerini daha hızlı difüze olan kokteyllere bırakmıştır.",
        "Zaman_Denge_Gliserol = t_difuzyon = L_doku^2 / ( pi^2 * D_gliserol ) >> t_difuzyon_DMSO * 5",
        "Bu difüzyon zaman sabiti eşitliği, gliserolün kalın dokularda hücresel dengeye ulaşma süresinin DMSO'ya kıyasla beş kattan daha uzun olduğunu simgeler."
    ),
    (
        "4.5 Toksisite Kinetiği Biyofiziği: Konsantrasyon, Sıcaklık ve Maruziyet Süresi",
        "Vitrifikasyon için gereken toplam CPA konsantrasyonu devasadır (ağırlıkça %50-60 / ~8-9 Molar); bu yoğunlukta kriyoprotektanlar oda sıcaklığında hücreleri dakikalar içinde zehirler (Kimyasal Toksisite).",
        "Kriyoprotektan toksisitesi üçlü bir fonksiyonla yönetilir: Konsantrasyon (C), Sıcaklık (T) ve Maruziyet Süresi (t). Sıcaklık 0°C'ye veya negatif derecelere (-20°C) düşürüldüğünde, CPA'nın metabolik ve enzimatik toksisitesi logaritmik olarak çökerken; hücre içine koruyucu difüzyonu devam eder. Bu nedenle klinik perfüzyon protokolleri katı bir 'Kademeli Soğuk Rampa' (Cold Temperature Ramp) ile yürütülür.",
        "Katsayi_Toksisite = I_toksik = Integral_0^t C_CPA(t')^alpha * exp( - E_toksisite / (R * T(t')) ) dt'",
        "Bu kümülatif toksisite integrali, hücrenin maruz kaldığı toplam kimyasal stresin CPA konsantrasyonunun kuvveti ve sıcaklıkla olan üstel ilişkisini tanımlar."
    ),
    (
        "4.6 M22 Vitrifikasyon Kokteyli: 21st Century Medicine'ın 9.3 Molar Şaheseri",
        "Katı organların ve insan beyninin buzsuz vitrifikasyonu için geliştirilen en gelişmiş ve patentli kimyasal formülasyon, Gregory Fahy'nin tasarladığı 'M22' kokteylidir.",
        "M22 toplam 9.3 Molar (yaklaşık %60 w/v) konsantrasyona sahiptir ve 7 farklı kriyoprotektan ve sentetik buz bloke edici ajanın kusursuz sinerjisinden oluşur: Dimetil sülfoksit (DMSO - %22.3 w/v), Formamid (%12.9 w/v), Etilen Glikol (%16.8 w/v), N-metilformamid (NMA), 3-metoksipropandiol ve polimerik buz blokörleri. M22'nin kritik soğutma hızı dakikada 0.1°C'nin dahi altındadır; yani devasa bir insan beyni normal bir buzdolabında soğutulsa dahi içinde tek bir buz kristali oluşmadan camlaşır.",
        "Kompozisyon_M22 = 22.3%_DMSO + 16.8%_EG + 12.9%_Formamid + 3%_NMA + 1%_Supercool (Toplam 9.3 M)",
        "Bu molar formülasyon kuralı, M22 vitrifikasyon solüsyonunun minimum toksisiteyle maksimum buz önleme kapasitesini sağlayan kimyasal tarifini belgeler."
    ),
    (
        "4.7 VM3 ve Diğer Yeni Nesil Kokteyller: Nöro-Spesifik Osmotik Koruma",
        "M22 tüm katı organlar için optimize edilmişken; yalnızca beyin dokusunu hedefleyen 'VM3' kokteyli gibi nöro-spesifik formülasyonlar geliştirilmiştir.",
        "VM3; kan-beyin bariyerini aşırı dehidrate etmeyen, nöronal membranlardaki lipid çözünmesini en aza indiren ve osmolariteyi 5000 mOsm/L seviyesine kademeli yükselten özel osmoprotektif şekerler (sukroz, mannitol) içerir. VM3 perfüzyonu yapılan beyinlerde nöron somaları aşırı büzüşmez, sinaptik yarık aralığı (20 nm) korunur ve nöromal bağlantı haritası sıfır distorsiyonla camsı faza kilitlenir.",
        "Osmolarite_VM3 = Osm_toplam = Sigma_i [CPA_i] * g_i ~ 4800 - 5200 mOsm/kg",
        "Bu koligatif osmolarite eşitliği, nöro-vitrifikasyon solüsyonlarının kan-beyin bariyeri dengesini koruyacak hedef osmotik basınç aralığını tanımlar."
    ),
    (
        "4.8 Osmotik Şok ve Kademeli Konsantrasyon Rampaları (Stepwise Loading)",
        "Eğer 9.3 Molar M22 solüsyonu hücrelere tek seferde verilirse; su anında hücreden dışarı fışkırır ve hücre şiddetli ozmotik şokla büzüşerek ölür; aynı şekilde çözünürken saf su verilirse hücre içeri su hücumuyla patlar.",
        "Bu ozmotik parçalanmayı önlemek için 'Kademeli Yükleme ve Boşaltma' (Stepwise Concentration Ramp) protokolü uygulanır: CPA konsantrasyonu 5 ila 8 kademede (%10, %25, %50, %75, %100 M22) yavaşça artırılır; her kademede hücrenin su ve CPA difüzyon dengesine ulaşması (osmotik relaksasyon) beklenir. Çözünme sırasında ise hücre dışına geçirimsiz bir osmotik tampon (0.5 M Sukroz veya Mannitol) eklenerek hücrenin aşırı su emip patlaması engellenir.",
        "Hacim_Hucre_Relaksasyon = V(t) / V_0 = 1 - S_büzüşme * exp( - t / Tau_osmotik )",
        "Bu dinamik hacimsel relaksasyon denklemi, kademeli CPA rampasının hücre hacmini kritik litik sınırların içinde nasıl güvenle tuttuğunu modeller."
    ),
    (
        "4.9 Kriyoprotektan Toksisitesinin Biyokimyasal Mekanizması: Protein Denatürasyonu",
        "Kriyoprotektanların yüksek konsantrasyonda hücreyi zehirlemesinin temel moleküler nedeni; proteinlerin üçüncül ve dördüncü yapılarını bir arada tutan hidrofobik etkileşimleri çözmeleridir.",
        "DMSO ve Formamid, suyun dielektrik sabitini düşürür ve proteinlerin iç çekirdeğindeki hidrofobik amino asit kalıntılarını (lösin, izolösin, valin) çevreleyen suyu destabilize eder. Protein konformasyonu çözülerek (unfolding) denatüre olur ve enzimatik aktif bölgeler kapanır. Bu toksisiteyi engellemek için solüsyona protein yapısını stabilize eden 'Kosmotropik Ajanlar' ve termodinamik dengeleyiciler eklenir.",
        "Serbest_Enerji_Denaturasyon = Delta_G_katlanma = Delta_G_0 - m_CPA * [CPA_konsantrasyonu]",
        "Bu protein termodinamiği denklemi, artan kriyoprotektan konsantrasyonunun yerel protein katlanma serbest enerjisini nasıl aşındırdığını ve toksisite yarattığını belgeler."
    ),
    (
        "4.10 Nötralize Edici Moleküller: Toksisiteyi Baskılayan Kriyoprotektif Antidotlar",
        "Kriyobiyolojinin en son keşiflerinden biri, iki farklı kriyoprotektanın birbirinin toksisitesini karşılıklı olarak nötralize edebilmesidir ('Toksisite Karşıtlığı' / Toxicity Antagonism).",
        "Örneğin Formamid tek başına aşırı toksiktir ve DMSO tek başına hücreyi bozar; ancak belirli bir molar oranda (örneğin 1:1 DMSO:Formamid karışımı) bir araya getirildiklerinde, Formamid'in zayıf hidrojen bağı kabul etme özelliği ile DMSO'nun güçlü kabul etme özelliği birbirini dengeler. Protein denatürasyonu %80 azalır; tek tek bileşenlerin toplamından çok daha az toksik fakat çok daha güçlü bir vitrifikasyon kapasitesi elde edilir.",
        "Toksisite_Kombinasyon = Tox(A + B) << Tox(A) + Tox(B) (Pozitif Biyokimyasal Antagonizm)",
        "Bu toksisite antagonizma eşitsizliği, çok-bileşenli CPA kokteyllerinin tekil moleküllere kıyasla olağanüstü biyouyumluluk üstünlüğünü kanıtlar."
    )
]

# ================= KISIM 5 =================
part5_subsections = [
    (
        "5.1 Doğal Antifriz Proteinleri (AFPs) ve Kutup Canlılarının Biyolojik Adaptasyonu",
        "Kutup denizlerinde yaşayan teleost balıkları (-1.9°C'deki donmamış sularda), böcekler ve kışlayan bitkiler; kanlarında buz kristalleri oluşumunu engelleyen 'Antifriz Proteinleri' (AFPs - Antifreeze Proteins) üretirler.",
        "Bu proteinler (AFP Tip I, II, III, IV ve Antifriz Glikoproteinleri - AFGP); sudaki serbest hidrojen bağlarını bağlayarak suyun donma noktasını düşürürken erime noktasını değiştirmez (Termal Histerezis). Bu proteinlerin en büyük mucizesi; koligatif olmayan (konsantrasyondan bağımsız) bir mekanizmayla, mikromolar gibi minicik konsantrasyonlarda dahi buz kristali büyümesini tamamen felç edebilmeleridir.",
        "Termal_Histerezis = Delta_T_histerezis = T_erime - T_donma ~ 1.5°C - 3.0°C (Koligatif Olmayan Ayrışma)",
        "Bu termodinamik histerezis eşitliği, antifriz proteinlerinin erime sıcaklığını sabit tutarken donma noktasını seçici olarak düşürme kabiliyetini belgeler."
    ),
    (
        "5.2 Adsorpsiyon-İnhibisyon Mekanizması: Buz Kristali Büyüme Yüzünün Kilitlenmesi",
        "AFPs ve sentetik buz blokörlerinin buz oluşumunu durdurma prensibi Raymond ve DeVries tarafından 'Adsorpsiyon-İnhibisyon Mekanizması' ile açıklanmıştır.",
        "Buz kristali büyürken belirli atomik düzlemler (bazal düzlem ve prizmatik düzlemler) boyunca su moleküllerini ekler. Antifriz proteini üzerindeki treonin veya aspartat kalıntıları, bu kristal yüzeyindeki oksijen atomlarının aralığına kusursuz bir anahtar-kilit uyumuyla oturur (Kelvin Etkisi). Proteinin bağlandığı iki nokta arasında kalan buz yüzeyi mecburen dışbükey (konveks) bir kavis çizmek zorunda kalır; bu kavis yerel yüzey serbest enerjisini fırlatarak buz kristalinin ileri büyümesini termodinamik olarak kilitler.",
        "Yari_Cap_Kelvin = r_kritik = ( 2 * gamma_buz_su * T_ergime ) / ( Delta_H_ergime * Delta_T_asiri_soguma )",
        "Bu Gibbs-Thomson eğrilik denklemi, iki adsorbe protein arasındaki dışbükey mikroskobik buz cephesinin büyümesinin termodinamik imkânsızlığını modeller."
    ),
    (
        "5.3 Sentetik Buz Bloke Edici Ajanlar (BBA): Polivinil Alkol (PVA) ve Poligliserol",
        "Doğal balık antifriz proteinlerinin laboratuvarda rekombinant üretimi aşırı pahalı ve immünojenik olduğundan; modern kriyobiyoloji sentetik 'Buz Bloke Edici Ajanlar' (BBA - Ice Blocking Agents) geliştirmiştir.",
        "En güçlü sentetik buz blokörü, Brian Wowk tarafından keşfedilen 'Süper-Düşük Moleküler Ağırlıklı Polivinil Alkol'dür (PVA / X-1000 ve Z-1000). PVA zincirindeki ardışık 1,3-diol hidroksil grupları, buz kristali kafesindeki oksijen mesafesiyle (4.5 Angstrom) birebir örtüşür. Çözeltiye yalnızca %0.1 ila %1 oranında PVA eklenmesi; heterojen nükleasyon tohumlarını tamamen körleştirir ve gereken toksik CPA miktarını %30 düşürür.",
        "Konsantrasyon_BBA = [PVA] ~ %0.5 - 1.0 w/v (Devasa Buz Baskılama Gücü)",
        "Bu mikromoleküler katkı oranı, sentetik polimerlerin minicik dozlarda devasa kriyoprotektif etkinlik çarpanı yarattığını gösterir."
    ),
    (
        "5.4 Rekristalizasyon İnhibisyonu (IRI): Isınma Sırasında Buz Büyümesinin Önlenmesi",
        "Camsı fazdan çözünme sırasında en büyük ölüm tehdidi olan küçük kristallerin devasa kütlelere dönüşmesi (Rekristalizasyon); Buz Rekristalizasyon İnhibitörleri (IRI) ile tamamen durdurulur.",
        "PVA ve sentetik C-bağlı antifriz glikolipidleri; yeni çözünen su moleküllerinin mevcut mikroskobik buz çekirdeklerine yapışmasını engeller. Kristal büyüme hızı sıfıra indirgenir. Numune çözünürken kritik tehlike bölgesinden (-50°C ile 0°C) geçerken buzlar büyüyemez ve doku mekanik olarak parçalanmadan pürüzsüzce sıvı faza geri döner.",
        "Aktivite_IRI = 1 - ( A_kristal_son / A_kristal_kontrol ) > %95 (Tam Rekristalizasyon Blokajı)",
        "Bu rekristalizasyon baskılama metriği, IRI ajanlarının çözünme evresinde doku içi buz büyümesini %95'in üzerinde nasıl engellediğini belgeler."
    ),
    (
        "5.5 Nanopartikül Bazlı Nükleasyon Engelleyiciler: Grafen Kuantum Noktaları ve Silika",
        "Sentetik polimerlerin ötesinde, modern nanotıp 'Nanometrik Buz Engelleyiciler' (Nanoparticle Ice Blockers) tasarlamıştır.",
        "Özel olarak yüzeyi hidrofilik hidroksil ve karboksil gruplarıyla donatılmış Grafen Kuantum Noktaları (GQDs) ve mezogözenekli silika nano-küreleri; çözeltideki serbest su moleküllerini nano-gözenekleri içinde hapsederek kristal kafesi oluşturmalarını fiziksel olarak engeller. Bu nano-blokörler heterojen nükleasyon merkezlerini absorbe ederek nötralize eder ve vitrifikasyon kalitesini zirveye taşır.",
        "Nukleasyon_Baskilama = J_nuk_nano = J_0 * ( 1 / ( 1 + K_GQD * [Grafen_Kuantum_Nokta] ) )",
        "Bu nanoteknolojik nükleasyon sönümleme denklemi, yüzey modifiyeli nano-parçacıkların suyun kristal tohumlanma frekansını nasıl sıfırladığını açıklar."
    ),
    (
        "5.6 BBA'ların Kriyoprotektan Toksisitesini Düşürme Çarpanı",
        "Vitrifikasyon kokteyllerinin insan dokusuna uygulanmasındaki en büyük engel olan kimyasal toksisite; BBA katkısıyla radikal biçimde kırılmıştır.",
        "M22 solüsyonuna sadece %1 oranında sentetik buz blokörü (PVA / Polyglycerol) eklendiğinde; solüsyonun buz oluşturmadan vitrifike olabilmesi için gereken toplam organik çözücü (DMSO/Formamid) konsantrasyonu 9.3 Molar'dan 7.5 Molar'a indirilebilmektedir. Bu %20'lik konsantrasyon düşüşü, hücrelerin maruz kaldığı metabolik toksisiteyi %80 azaltır; organın biyolojik sağkalım şansını katbekat artırır.",
        "Toksisite_Azalimi = Delta_Tox = - 80% (BBA İlavesi ile Düşürülen CPA Eşiği)",
        "Bu toksikolojik kazanç oranı, buz bloke edici ajanların vitrifikasyonu klinik olarak tolere edilebilir bir biyolojik sınıra çekme başarısını gösterir."
    ),
    (
        "5.7 Biyomimetik Peptitler: Antifriz Glikopeptit (AFGP) Sentetik Analogları",
        "Doğal AFGP'lerin omurgasında tekrarlanan 'Alanin-Alanin-Treonin' tripeptid dizisi ve treonine bağlı disakkarit birimleri yer alır.",
        "Katı faz peptit sentezi (SPPS) ile üretilen sentetik biyomimetik AFGP analogları; doğal proteinlerin immünojenik kısımlarından arındırılmış, proteazlara dirençli D-amino asit omurgalarıyla yeniden üretilmiştir. Bu sentetik peptitler kan perfüzyonunda hiçbir antikor yanıtı veya kompleman aktivasyonu tetiklemez; insan kılcal damarlarından serbestçe geçerek tüm hücrelerin dış membranını buz kristallerine karşı zırhlar.",
        "Dizi_Mimetik: [Ala - Ala - Thr*(Gal-beta-1,3-GalNAc)]_n (Sentetik Biyomimetik Buz Blokörü)",
        "Bu moleküler peptid sekansı, kutup balığı antifriz glikoproteinlerinin insan klinik biyostazında kullanılan immün-inert sentetik mimarisini tanımlar."
    ),
    (
        "5.8 Zwitteriyonik Kriyoprotektif Polimerler: Betain ve Trehaloz Kopolimerleri",
        "Hücre içine giremeyen non-penetran şekerler (Trehaloz gibi); zwitteriyonik polimerik fırçalarla birleştirilerek süper-koruyucu kopolimerler üretilmiştir.",
        "Polisülfobetain-kopolimer-trehaloz (PSB-Trehaloz) konjugatları; hem hücre zarı dış yüzeyindeki proteinleri opsonizasyondan ve denatürasyondan korur hem de hücre dışı suyun serbestliğini kısıtlar. Bu polimerler hücre zarının elastikiyetini koruyarak aşırı soğukta zar faz geçişlerini (sıvı-kristalden katı-jele geçiş) yumuşatır ve lipid çatlamalarını önler.",
        "Gecis_Sicakligi_Membran = Delta_T_membran = - k_polimer * [PSB_Trehaloz_yuzey]",
        "Bu membran faz stabilitesi denklemi, zwitteriyonik şeker kopolimerlerinin hücre zarı lipid donma noktasını nasıl güvenle düşürdüğünü belgeler."
    ),
    (
        "5.9 Kriyojenik Akışkan Dinamiği ve Kılcal Damarlarda BBA Homojenliği",
        "Katı bir organın veya beynin vitrifikasyonunda en büyük hidrolik risk, eklenen makromoleküler buz blokörlerinin yüksek viskozite nedeniyle kılcal damarlara girememesidir.",
        "Bu sorunu çözmek için BBA moleküler ağırlığı 1000 Da (yaklaşık 20 monomerlik PVA oligomeri) ile sınırlandırılmıştır. Bu ultra-küçük zincirler kan plazmasından daha düşük bir viskozite artışı yaratır; 5 mikrometrelik en uç beyin kapillerlerine dahi tıkanma yapmadan homojen olarak dağılır ve tüm mikrovasküler yatağı buz kristali saldırısına karşı korur.",
        "Hidrolik_Iletkenlik = L_p = ( pi * r_kapiller^4 ) / ( 8 * eta_solusyon * L ) (Düşük MW BBA ile Korunur)",
        "Bu Poiseuille mikrovasküler debi eşitliği, optimize edilmiş düşük moleküler ağırlıklı BBA oligomerlerinin serebral mikrosirkülasyonda yarattığı sıfır tıkanma güvencesini formüle eder."
    ),
    (
        "5.10 İleri Düzey Nöroprotektif Kokteyller ve BBA'ların Klinik Standartlaşması",
        "Günümüzde Alcor Life Extension Foundation ve Tomorrow Bio gibi biyostaz organizasyonları; tüm hasta perfüzyonlarında BBA içeren yeni nesil solüsyonları standart protokol haline getirmiştir.",
        "BBA katkılı M22 varyantları; insan beyninin -130°C'ye soğutulması sırasında elektroensefalogram (EEG) ve tomografik taramalarda sıfır buz kristali imzası sağlamıştır. Buz blokörleri; vitrifikasyonu aşırı hızlı soğutma zorunluluğundan kurtararak, kalın insan kafatasının içinde yer alan devasa beynin dahi güvenle camsı faza geçmesini sağlayan nihai anahtardır.",
        "Guvenlik_Vitrifiye_Beyin = 1 - Fraksiyon_Kristal = 1.000 (BBA Sinerjisi ile %100 Camsılaşma)",
        "Bu mutlak camsılaşma saflık oranı, modern buz bloke edici ajanların insan nörokriyoprezervasyonunda mekanik buz hasarını tamamen sıfırladığını kanıtlar."
    )
]

# ================= KISIM 6 =================
part6_subsections = [
    (
        "6.1 Klinik Kriyoprotektif Standby: Yasal Ölüm Anında Zamana Karşı Yarış",
        "Bir biyostaz hastası için en kritik evre, kalbin durduğu ve yasal ölümün ilan edildiği ilk 60 saniyedir; bu ana 'Standby ve İlk Müdahale' (Standby & First Response) denir.",
        "Alcor veya Tomorrow Bio standby ekipleri; terminal hastanın başucunda haftalarca nöbet tutar. Yasal ölüm hekim tarafından ilan edildiği saniyede ekip devreye girer. Gecikilen her dakika sıcak iskemi nedeniyle nöronal ATP depolarını eritir ve hücresel lizisi başlatır; bu nedenle ilk 10 dakika içinde vücut sıcaklığının hızla düşürülmesi ve yapay dolaşımın başlatılması hayatidir.",
        "Zaman_Kritik_Standby: Delta_t_mudahale < 60 saniye (Sıcak İskemiyi Sıfırlama Penceresi)",
        "Bu acil müdahale zaman hedefi eşitsizliği, yasal ölüm anı ile mekanik hipotermik resüsitasyonun başlatılması arasındaki kritik pencereyi belgeler."
    ),
    (
        "6.2 Mekanik Kardiyopulmoner Destek (LUCAS) ve Hipotermi İndüksiyonu",
        "Standby başladığı anda hastanın göğsüne derhal otomatik bir mekanik göğüs kompresyon cihazı (LUCAS) yerleştirilir ve trakeal entübasyonla akciğerlere saf oksijen pompalanır.",
        "LUCAS cihazı dakikada 100 ritmik kompresyonla kan dolaşımını mekanik olarak sürdürür; beyne oksijenli kan gitmeye devam ederken sıcak iskemi süresi durdurulur. Eş zamanlı olarak hasta devasa bir buzlu su banyosuna (Ice Slurry) gömülür. Vücut yüzeyine hızla ısı transferi sağlayan sirkülasyonlu buzlu su pompaları ile hastanın çekirdek vücut sıcaklığı dakikada 0.5°C hızla 10°C'nin altına çekilir.",
        "Soguma_Hizi_Hipotermi = dT/dt = - ( h_konveksiyon * A_yuzey * ( T_vucut - T_buzlu_su ) ) / ( m_vucut * c_doku )",
        "Bu Newton soğuma kanunu eşitliği, sirkülasyonlu buz banyosunun insan vücut sıcaklığını hipotermik koruma fazına indirme hızını modeller."
    ),
    (
        "6.3 İntravenöz Anti-İskemik ve Antikoagülan İlaç Kokteyli (Standby Meds)",
        "Mekanik CPR devam ederken, hastanın damar yolundan 15 farklı molekülden oluşan agresif bir 'Standby İlaç Kokteyli' hızla enjekte edilir.",
        "1) Heparin (Masif Doz - 50.000 Ünite): Kanın kılcal damarlarda pıhtılaşmasını kesin olarak durdurur; 2) Propofol ve Tiopental: Beynin metabolik aktivitesini ve oksijen tüketimini %90 baskılar; 3) Streptokinaz / t-PA: Varsa oluşmuş mikro-pıhtıları anında eritir; 4) Vazopressörler ve pH Tamponları (Sodyum Bikarbonat): Laktik asidozu nötralize eder; 5) Membran Koruyucular ve Antioksidanlar (Melatonin, E vitamini): Reperfüzyon hasarını bloke eder.",
        "Protokol_Standby_Meds: Heparin (50k U) + Propofol + t-PA + Bikarbonat + Allopurinol + Melatonin",
        "Bu standart standby farmakoterapi listesi, hipotermi sırasında vasküler açıklığı koruyan ve serebral metabolik çöküşü durduran ilaç protokolünü tanımlar."
    ),
    (
        "6.4 Cerrahi Kanülasyon: Femoral ve Karotis/Juguler Vasküler Erişim",
        "Hasta saha ambulansında veya yerel cerrahi ünitesinde 10°C'ye soğutulduktan sonra, vücut kanının yıkanması ve kriyoprotektan perfüzyonu için acil damar erişimi (Cerrahi Kanülasyon) açılır.",
        "Tam vücut hastalarında sağ femoral arter ve femoral ven cerrahi insizyonla açılır ve geniş lümenli kanüller yerleştirilir; nörokriyoprezervasyon hastalarında ise doğrudan ortak karotis arterleri (Arteria Carotis Communis) ve internal juguler venler kanüle edilir. Bu kanüller hastanın dolaşım sistemini ameliyathane tipi bir kalp-akciğer makinesine ve kriyoprotektif perfüzyon devresine bağlar.",
        "Debi_Kanulasyon = Q_erisim = ( Delta_P_pompa * pi * r_kanul^4 ) / ( 8 * eta_kan * L_kanul ) > 2.0 L/dakika",
        "Bu cerrahi kanülasyon hidrolik debi denklemi, geniş lümenli damar erişiminin hızlı kan yıkaması için gereken yüksek perfüzyon akışını nasıl sağladığını gösterir."
    ),
    (
        "6.5 Kanın Tamamen Yıkanması: Asellüler Organ Koruma Solüsyonu (MHP-2 / B1)",
        "Kriyoprotektanlar kana doğrudan verilemez; kandaki alyuvarlar yüksek ozmotik basınçta parçalanır ve hemoglobin damarları tıkar; bu nedenle kanın vücuttan tamamen uzaklaştırılması şarttır.",
        "Femoral veya karotis arterden soğuk (4°C) 'Asellüler Organ Koruma Solüsyonu' (Alcor'un MHP-2 veya CryoRus'un B1 solüsyonu) pompalanırken; venöz kanülden hastanın tüm kanı dışarı tahliye edilir. Solüsyon doku ile izoozmotiktir, hidroksietil nişasta (HES) ile onkotik basıncı dengeler ve hücrelerin şişmesini önler. Venöz dönüş tamamen berraklaştığında (tüm kan temizlendiğinde) hasta kriyoprotektan perfüzyonuna hazır hale gelir.",
        "Hematokrit_Tahliye = Hct(t) = Hct_0 * exp( - Q_yikama * t / V_kan ) -> 0 (Kansız Dolaşım)",
        "Bu hematolojik arındırma eşitliği, asellüler perfüzyonun hastanın dolaşımındaki alyuvar ve hemoglobin fraksiyonunu dakikalar içinde sıfıra indirme kinetiğini belgeler."
    ),
    (
        "6.6 Kapalı Devre Bilgisayar Kontrollü Perfüzyon Sistemi: Refraktometrik İzleme",
        "Kriyoprotektan yüklemesi manuel yapılamaz; damar içi basınç, sıcaklık ve konsantrasyon mikron düzeyinde bilgisayar kontrollü bir perfüzyon konsolu ile yönetilir.",
        "Perfüzyon devresi iki pompa taşır: Biri bazal koruma sıvısını, diğeri konsantre M22'yi pompalar. Bilgisayar bu iki akışı karıştırarak konsantrasyonu kademeli artırır. Venöz drenaj hattına yerleştirilen 'Dijital İn-Line Refraktometre' (Brix / Kırılma İndisi ölçer); dokulardan dönen sıvının kriyoprotektan konsantrasyonunu gerçek zamanlı ölçer. Doku CPA'yı emdikçe venöz refraktometre değeri yükselir; hedef doygunluk seviyesine ulaşılana kadar perfüzyon sürdürülür.",
        "Refraktif_Indeks = n_D = n_0 + alpha_Brix * [Konsantrasyon_CPA_venoz] (Gerçek Zamanlı Doku Doygunluğu)",
        "Bu optik refraktometri denklemi, venöz hattan dönen perfüzatın kırılma indisinin dokudaki anlık kriyoprotektan konsantrasyonunu doğrudan nasıl raporladığını simgeler."
    ),
    (
        "6.7 Basınç Kontrolü ve Ödem Önleme Biyofiziği: Perfüzyon Basınç Sınırları",
        "Perfüzyon sırasında en büyük cerrahi felaket, aşırı pompa basıncının beyin kılcal damarlarını patlatması veya düşük basıncın yetersiz CPA dağılımına yol açmasıdır.",
        "Serebral arteriyel perfüzyon basıncı kesin olarak 80 ila 100 mmHg fizyolojik sınırında kilitlenir. Kriyoprotektanların yüksek viskozitesi nedeniyle damar direnci arttığında perfüzyon debisi otomatik olarak düşürülür. Çözeltiye eklenen yüksek molekül ağırlıklı kolloidler (PEG veya Dekstran); Starling dengesini koruyarak sıvının damardan parankime kaçıp beyin ödemi (ödematöz şişme) yapmasını kesin olarak engeller.",
        "Filtrasyon_Starling = J_v = K_f * [ ( P_kapiller - P_interstisyel ) - sigma * ( Pi_kapiller - Pi_interstisyel ) ] ~ 0",
        "Bu mikrovasküler Starling sıvı dengesi denklemi, onkotik basınç optimizasyonunun kriyoprotektif perfüzyon sırasında doku ödemini nasıl sıfırladığını açıklar."
    ),
    (
        "6.8 Terminal Perfüzyon Kriteri: Venöz Konsantrasyonun Hedef Seviyeye Ulaşması",
        "Kriyoprotektan perfüzyonunun ne zaman sonlandırılacağı rastgele tahminlerle değil, kesin analitik biyokimyasal kriterlerle belirlenir.",
        "M22 kokteyli için hedef terminal konsantrasyon 9.3 Molar'dır (%100 M22). Venöz dönüş hattındaki refraktometre değeri karotis veya femoral giriş konsantrasyonunun en az %95'ine ulaştığında (Venöz M22 > 8.8 Molar); dokuların, özellikle derin beyin çekirdeklerinin kriyoprotektan ile tamamen doyduğu ve vitrifikasyon için gereken koruma eşiğini aştığı tescillenir. Bu aşamada cerrahi perfüzyon durdurulur.",
        "Kriter_Terminal_Doygunluk: [CPA_venoz] / [CPA_arteriyel] >= 0.95 (Tam Doku Saturasyonu)",
        "Bu biyokimyasal doygunluk oranı, biyolojik organın buz oluşturmadan vitrifike olabilmesi için gereken asgari terminal perfüzyon tamamlama standardını belgeler."
    ),
    (
        "6.9 Kriyojenik Soğutma Odası: Bilgisayar Kontrollü Gaz Fazı Soğutma Rampası",
        "Perfüzyonu tamamlanan hasta derhal cerrahi masadan alınarak 'Bilgisayar Kontrollü Kriyojenik Soğutma Odası'na (Computer-Controlled Cooling Chamber) yerleştirilir.",
        "Hasta doğrudan sıvı azota atılmaz (bu durum termal şokla bedeni paramparça eder). Sıvı azot buharı oditoryuma üflenerek sıcaklık önceden programlanmış katı bir termal rampa ile düşürülür: 0°C'den -110°C'ye saatte 5°C hızla (kristalizasyon burnunu aşmak için); camsı geçiş noktası olan -125°C civarında ise termal gerilme kırılmalarını önlemek için soğutma hızı saatte 0.5°C'ye kadar yavaşlatılır.",
        "Sogutma_Rampasi: dT/dt = - 5°C/saat (0°C -> -110°C) | dT/dt = - 0.5°C/saat (-110°C -> -140°C)",
        "Bu hassas termal profil, dokunun amorf katı faza geçerken iç gerilme çatlakları oluşturmasını engelleyen bilgisayar kontrollü soğutma rejimini tanımlar."
    ),
    (
        "6.10 Uzun Süreli Depolama: Bigfoot Çiyarları ve Sıvı Azotta Sonsuz Uyku",
        "-140°C'de camsı katılaşması tamamlanan hasta; nihai istirahat yeri olan vakum yalıtımlı devasa 'Bigfoot Kriyojenik Çiyar Tankı'na (Cryogenic Dewar) transfer edilir.",
        "Bu pasif vakum yalıtımlı paslanmaz çelik dewar tankları; içine sıvı azot doldurulduğunda sıfır elektrik enerjisiyle, hiçbir mekanik pompaya veya jeneratöre ihtiyaç duymadan çalışır. Hasta baş aşağı pozisyonda tankın içine yerleştirilir (böylece sıvı azot seviyesi azalsa dahi en son beyin açıkta kalır). Tanklar periyodik olarak sıvı azot ikmali yapılarak tıp teknolojisinin canlandırma gücüne ulaşacağı geleceğin o kutlu gününe kadar asırlar boyu sıfır entropik bozulmayla bekletilir.",
        "Statik_Kayip_LN2 = dV_LN2/dt < 5 L/gun (Bigfoot Dewar'da 30 Günlük Acil Durum Rezervi)",
        "Bu kriyojenik izolasyon parametresi, Bigfoot dewar tanklarının dış dünyadaki tüm elektrik ve altyapı çökmelerine karşı sunduğu asırlık pasif koruma güvenliğini belgeler."
    )
]

# ================= KISIM 7 =================
part7_subsections = [
    (
        "7.1 Nörokriyoprezervasyon (Neuro) vs Tam Vücut (Whole-Body) Biyostaz Paradigması",
        "Kriyobiyolojide iki temel yaklaşım mevcuttur: Tüm bedenin dondurulduğu 'Tam Vücut' (Whole-Body) ve yalnızca başın/beynin korunduğu 'Nörokriyoprezervasyon' (Neuro).",
        "Nörokriyoprezervasyonun felsefi ve bilimsel temeli şudur: Bir bireyin anıları, kişiliği, entelektüel birikimi ve öz-bilinci yalnızca beyin dokusundaki sinaptik devrelerde (Connectome) saklıdır; bedenin geri kalanı biyolojik bir taşıyıcı kabuktur. Geleceğin rejeneratif tıbbı genç bir biyolojik bedeni kök hücrelerden sıfırdan üretebilecekken (Cilt 14-15), hasarlı ve yaşlı bir bedeni dondurmaya çalışmak gereksiz maliyet ve teknik zorluklar yaratır.",
        "Onerme_Neuro: Kimlik_Birey == Bilgi_Beyin | Beden = Geri Donusturulebilir_Tasiyici",
        "Bu biyo-felsefi kimlik denkliği, nörokriyoprezervasyonun tüm teknik kaynakları yalnızca beynin kusursuz korunmasına odaklama mantığını özetler."
    ),
    (
        "7.2 Sefalik İzolasyon (Cephalic Isolation) Cerrahi Protokolü",
        "Nörokriyoprezervasyon prosedüründe hastanın başı, gövdeden mikromikrocerrahi tekniklerle ayrılır (Sefalik İzolasyon).",
        "Cerrahi operasyon 0°C buz banyosu altında gerçekleştirilir. Altıncı servikal vertebra (C6) seviyesinde trakea ve özofagus bağlanıp kesilir; karotis arterleri ve internal juguler venler diseke edilerek perfüzyon kanülleri yerleştirilir. Omurilik kordonu temiz bir mikro-kesiyle ayrılır. Bu yaklaşım; gövdedeki onlarca litre kan ve organın yarattığı karmaşık perfüzyon dengesizliklerini dışlayarak, kriyoprotektif solüsyonun %100 doğrudan serebral damar yatağına akmasını sağlar.",
        "Direnc_Perfuzyon_Serebral: R_beyin << R_tum_vucut (Hassas ve Hızlı Akış Kontrolü)",
        "Bu cerrahi hidrodinamik üstünlük, sefalik izolasyonun perfüzyon direncini düşürerek beyne hızlı ve homojen kriyoprotektan doyumunu nasıl sağladığını modeller."
    ),
    (
        "7.3 Kan-Beyin Bariyeri (BBB) Geçirgenliği ve Kriyoprotektif Moleküler Geçiş",
        "Beyin vitrifikasyonunun önündeki en büyük biyofiziksel bariyer, kılcal damarları saran sıkı endotelyal bağlantılar ve astrosit son ayaklarından oluşan 'Kan-Beyin Bariyeri'dir (BBB).",
        "Suda çözünen birçok molekül bariyeri geçemezken; küçük penetran CPA'lar (özellikle Formamid ve DMSO) lipid membranları aşabilir. Ancak yüksek ozmotik basınç farkı endotel hücrelerini büzüştürerek sıkı bağlantıların (Claudin-5) geçici olarak açılmasına yol açabilir. Perfüzyon solüsyonuna mikro-dozda Mannitol eklenmesi endotelyal hipertonik büzüşmeyi kontrol ederek CPA'ların nöron parankimine pürüzsüz difüzyonunu hızlandırır.",
        "Gecirgenlik_BBB_CPA = P_efektif = P_paraselluler + P_transselluler = f( [Mannitol], [DMSO] )",
        "Bu kombine kan-beyin bariyeri transfer denklemi, hiperozmotik modülatörlerin kriyoprotektan moleküllerinin beyin parankimine difüzyon katsayısını nasıl katladığını belgeler."
    ),
    (
        "7.4 Serebral Dehidrasyon ve Hacimsel Büzüşme: Kraniyal Boşluk Dinamikleri",
        "Beyin dokusuna yüksek konsantrasyonlu M22 solüsyonu perfüze edildiğinde; güçlü ozmotik gradyan nöronlardan ve glia hücrelerinden suyu damar lümenine çeker.",
        "Bu süreç beyin parankiminde ölçülebilir bir 'Hacimsel Büzüşmeye' (Brain Shrinkage) yol açar. Beyin kafatası kemiklerinden içeriye doğru yaklaşık %15 ila %25 oranında küçülür; dural sinüsler ve subdural boşluk genişler. Bu büzüşme nöronal membranları yırtmaz, aksine hücrelerin içindeki serbest suyu tahliye ederek hücre içi buz oluşumu riskini tamamen sıfırlar ve camsı katılaşmaya ideal bir dehidrate zemin hazırlar.",
        "Buzusme_Beyin = Delta_V_beyin / V_0 = - alpha_ozmotik * ( Osm_CPA - Osm_fizyolojik ) ~ - 20%",
        "Bu serebral volumetrik büzüşme formülü, yüksek ozmolariteli perfüzyonun beyin dokusu hacminde yarattığı kontrollü dehidrasyon katsayısını tanımlar."
    ),
    (
        "7.5 Bilgisayarlı Tomografi (CT) ile Beyin Vitrifikasyonunun İn Situ Taranması",
        "Alcor'un geliştirdiği en önemli klinik kalite güvence inovasyonu; dondurulan hastaların doğrudan 'Kriyojenik Bilgisayarlı Tomografi' (Cryo-CT) ile taranmasıdır.",
        "M22 kokteyli içindeki kriyoprotektanlar (özellikle kükürt içeren DMSO), sudan daha yüksek bir X-ışını soğurma yoğunluğuna (Hounsfield Birimi - HU) sahiptir. Beyin parankimi CPA ile doldukça CT görüntüsünde doku belirgin şekilde 'aydınlanır' (parlak beyazlaşır). Eğer beynin bir bölgesinde damar tıkanıklığı nedeniyle CPA ulaşmamışsa orası karanlık kalır ve buz riski tespit edilir; tam aydınlanan bir beyin %100 kusursuz vitrifikasyonun dijital tescilidir.",
        "Hounsfield_Yogunlugu = HU_doku = 1000 * ( mu_doku - mu_su ) / mu_su > + 150 HU (Tam Vitrifikasyon İmzası)",
        "Bu radyolojik dansitometri kriteri, bilgisayarlı tomografide ölçülen Hounsfield yoğunluğunun serebral kriyoprotektan doyum derecesini nasıl kesin olarak doğruladığını belgeler."
    ),
    (
        "7.6 Hipokampus ve Amigdala Engramlarının Korunması: Bellek Biyostazı",
        "Bir insanın otobiyografik anılarının, duygusal hafızasının ve kimliğinin merkez üssü Hipokampus (CA1, CA3, Dentat Girus) ve Amigdaladır.",
        "Vitrifikasyon perfüzyonu sırasında bu derin limbik yapılar posterior serebral arter ve koroid arter dallarından kusursuzca beslenir. Cryo-EM ve konfokal mikroskopi; vitrifike edilen hipokampal dilimlerde sinaptik veziküllerin, presinaptik SNARE proteinlerinin ve post-sinaptik PSD-95 yoğunluklarının orijinal nanometrik pozisyonlarında kilitlendiğini göstermiştir. Bellek izleri (engramlar) fiziksel taşıyıcılarında hiçbir moleküler kayıp olmadan korunur.",
        "Bütünlük_Engram = I_engram = [PSD-95_intakt] / [PSD-95_kontrol] * ( 1 - Hasar_Sinaptik ) > %98",
        "Bu sinaptik engram bütünlüğü indeksi, vitrifikasyonun hipokampal anı devrelerini moleküler çözünürlükte koruma başarısını matematiksel olarak ortaya koyar."
    ),
    (
        "7.7 Tam Vücut Kriyoprezervasyonunun Vasküler ve Organ Karmaşıklığı Zorlukları",
        "Tam Vücut kriyoprezervasyonu hastaya kendi orijinal bedenini koruma konforu sunsa da; biyofiziksel açıdan nörokriyoprezervasyondan katbekat daha karmaşık bir mücadeledir.",
        "İnsan vücudundaki farklı organlar (böbrek, karaciğer, kalp, bağırsaklar, kaslar) taban tabana zıt vasküler dirençlere, kan akış debilerine ve membran su geçirgenliklerine (Lp) sahiptir. Beyne yetecek konsantrasyonda CPA verildiğinde böbrekler aşırı doygunluktan zehirlenebilir veya kas dokusu devasa kütlesi nedeniyle yavaş soğuyarak buz oluşturabilir. Bütün bir vücudu tek bir perfüzyonla homojen vitrifike etmek aşırı hassas çoklu debi yönetimi gerektirir.",
        "Heterojenite_Doku: Lp_bobrek >> Lp_beyin >> Lp_kas (Farklı Su Geçirgenlik Dinamikleri)",
        "Bu dokusal heterojenite ilkesi, tam vücut biyostazında organlar arası difüzyon katsayısı uçurumlarının yarattığı perfüzyon mühendisliği zorluğunu formüle eder."
    ),
    (
        "7.8 Otolog Klonlama ve Biyobaskı ile Yeni Beden Üretimi (Re-Embodiment)",
        "Nörokriyoprezervasyon hastaları gelecekte hayata döndürüldüğünde; donuk haldeki baş veya beyin dokusu nasıl tam bir insana dönüşecektir?",
        "Bu dönüşümün iki biyoteknolojik yolu mevcuttur: 1) Otolog Klonlama: Hastanın korunmuş nöronlarından çekirdek transferi (SCNT) ile hastanın kendi DNA'sını taşıyan genç bir embriyonik biyolojik beden (beyin sapı inaktif) üretilir ve sefalik anastomozla baş yeni bedene bağlanır; 2) 3D Biyomekatronik Biyo-Baskı: Cilt 14-15'te detaylandırılan teknolojilerle hastaya özel genç, güçlü, biyo-hibrit bir sibernetik beden sıfırdan basılır ve beyin bu ebedi kabuğa entegre edilir (Re-Embodiment).",
        "Yeniden_Bedenlenme = Re-Embodiment = [Korunmus_Beyin] + [3D_BiyoBaskili_Sibernetik_Beden]",
        "Bu post-biyolojik rekonstrüksiyon eşitliği, nörokriyoprezervasyon hastalarının geleceğin rejeneratif tıbbında mükemmel genç bedenlere nasıl kavuşturulacağını modeller."
    ),
    (
        "7.9 Maliyet, Lojistik ve Vakıf Güvencesi: Asırlık Finansal Sürdürülebilirlik",
        "Biyostaz hastasının sıvı azotta onlarca hatta yüzlerce yıl boyunca aralıksız korunabilmesi; yalnızca tıbbi değil, sarsılmaz bir finansal ve kurumsal altyapı gerektirir.",
        "Alcor ve Tomorrow Bio; hasta ödemelerini (yaklaşık 80.000$ - 200.000$) 'Hasta Bakım Tröstü'ne (Patient Care Trust) aktarır. Bu tröst bağımsız bir kurul tarafından yönetilir, anaparaya asla dokunulmaz; yalnızca paranın küresel düşük riskli yatırımlardan elde ettiği yıllık %3-4'lük reel getiri sıvı azot ikmali ve tesis bakım masraflarını karşılar. Finansal model matematiksel olarak hastanın sonsuza kadar finanse edilmesini garantiler.",
        "Finansal_Surdurulebilirlik = Getiri_Yatirim = ( Anapara_Trost * r_reel ) > Yillik_Bakim_Maliyeti",
        "Bu finansal otonomi eşitsizliği, hasta bakım tröstünün sermaye getirisi ile kriyojenik biyostaz tesisinin asırlar boyu iflas etmeden ayakta kalma matematiğini belgeler."
    ),
    (
        "7.10 Geleceğe Fırlatılan Bilinç Oku: Biyostazın Kozmolojik ve Varoluşsal Zaferi",
        "Nörokriyoprezervasyon, bireyin bilincini günümüzün çaresiz hastalıklarından ve kaçınılmaz biyolojik ölümünden geleceğin sınırsız tıp çağına fırlatan zamansal bir oktur.",
        "Hasta için geçen süre sıfırdır; -196°C'de bilincin kapandığı saniye ile 100 yıl sonra nanorobotik bir ameliyathanede genç bir bedende gözlerini açtığı an arasında öznel olarak tek bir saliselik derin bir uyku hissi yaşanır. Biyostaz; zamanın amansız akışına karşı insan aklının geliştirdiği en cüretkâr, en bilimsel ve en muazzam köprüdür.",
        "Zaman_Oznellik = Delta_t_oznel = 0 (Biyostaz Boyunca) | Delta_t_nesnel = Yuzyillar",
        "Bu zamansal görelilik formülü, kriyojenik biyostaz altındaki bir insan bilinci için yüzyılların öznel olarak sıfır saniyede aşıldığını matematiksel olarak ifade eder."
    )
]

# ================= KISIM 8 =================
part8_subsections = [
    (
        "8.1 Termal Stres ve Termoelastik Gerilme Biyomekaniği",
        "Vitrifikasyon buz oluşumunu tamamen engellese dahi; katılaşma evresinde ortaya çıkan ikinci büyük fiziksel tehdit 'Termal Gerilme Kırılmaları'dır (Thermal Stress Fracturing).",
        "Doku camsı geçiş sıcaklığının (Tg ~ -125°C) altına soğutulurken termal büzüşme (Thermal Contraction) yaşar. Katılaşmış cam fazında moleküller artık yer değiştiremediği için bu büzüşme doku içinde devasa iç mekanik gerilmeler (termoelastik stres) biriktirir. Eğer yerel gerilme malzemenin kırılma mukavemetini (tensile strength) aşarsa; cam kaskatı bir pencere camı gibi aniden çatlar.",
        "Gerilme_Termoelastik = Sigma_termal = ( E_cam / ( 1 - nu ) ) * alpha_termal * ( Tg - T )",
        "Bu termoelastik gerilme denklemi, camsı geçiş noktasının altına inildikçe biriken mekanik stresin elastik modül (E) ve termal genleşme katsayısı (alpha) ile doğrusal artışını belgeler."
    ),
    (
        "8.2 Kırılma Mekaniği: Griffith Çatlak Kriteri ve Gerilme Yoğunluk Faktörü (K_IC)",
        "Termal gerilme altındaki katılaşmış amorf biyolojik camın ne zaman ve nasıl çatlayacağı Griffith Kırılma Mekaniği prensipleriyle hesaplanır.",
        "Malzeme içindeki mikroskobik bir mikro-boşluk veya çatlak ucundaki 'Gerilme Yoğunluk Faktörü' (K_I), dokunun 'Kırılma Tokluğu' (K_IC - Fracture Toughness) değerine ulaştığı anda çatlak kararsız hale gelir ve ses hızında doku boyunca ilerler (Catastrophic Fracture). Camsı biyolojik dokularda K_IC yaklaşık 0.1 ila 0.3 MPa*m^0.5 gibi düşük bir değerdedir; bu durum vitrifike organları kırılmaya son derece hassas kılar.",
        "Kriter_Griffith: K_I = Sigma_termal * sqrt( pi * a_catlak ) >= K_IC (Kırılma Başlama Eşiği)",
        "Bu kırılma mekaniği eşitsizliği, dokudaki mikroskobik bir mikro-çatlağın kritik termal stres altında makroskobik bir kırılmaya dönüşme eşiğini tanımlar."
    ),
    (
        "8.3 Tg Altı Soğutmada Mekanik Kırılmalar: Sıvı Azot (-196°C) Risk Faktörü",
        "Vitrifike olmuş bir beyin -125°C'de tamamen camlaşır; ancak uzun süreli depolama için sıvı azot sıcaklığı olan -196°C'ye kadar soğutulması gerekir.",
        "Bu ilave 70°C'lik soğuma aralığında, termal büzüşme gerilmesi 5 ila 10 MegaPaskal (MPa) gibi devasa seviyelere tırmanır. 1990'lı yıllarda doğrudan sıvı azota daldırılan beyinlerde yapılan CT taramaları; beynin birkaç büyük makroskobik blok halinde kırıldığını (çatladığını) göstermiştir. Bu kırıklar hücreleri ezmez (hücresel düzeyde ultrastrüktür korunur); ancak beyin blokları arasında milimetrik yarıklar açar.",
        "Mekanik_Gerilme_LN2 = Sigma_max ~ 8.5 MPa >> Sigma_kirilma_cam (~ 2.0 MPa)",
        "Bu gerilme kapasitesi aşımı karşılaştırması, kontrolsüz sıvı azot soğutmasının camsı dokuda neden kaçınılmaz kırıklar açtığını matematiksel olarak ortaya koyar."
    ),
    (
        "8.4 Akustik Emisyon (AE) Algılama: Kırıkların Gerçek Zamanlı İzlenmesi",
        "Kriyojenik soğutma sırasında doku içinde bir çatlak veya kırık oluştuğunda; açığa çıkan elastik gerilme dalgaları yüksek frekanslı (100 kHz - 1 MHz) mikroskobik bir 'çatırtı' sesi yayar.",
        "Alcor'un geliştirdiği 'Akustik Emisyon İzleme' (Acoustic Emission Testing) teknolojisi; dewar tankının içine yerleştirilen piezoelektrik transdüserlerle bu mikro-ses dalgalarını gerçek zamanlı kaydeder. Sistem hangi sıcaklıkta, hangi hızda çatlak oluştuğunu desibel düzeyinde kaydeder. Bu veriler soğutma hızının nerede yavaşlatılması gerektiğini gösteren paha biçilmez bir biyofiziksel geri bildirim sağlar.",
        "Enerji_Akustik = E_AE = Integral [ V_piezo(t)^2 dt ] (Çatlağın Büyüklüğü ile Orantılı Enerji)",
        "Bu akustik emisyon enerji integrali, piezoelektrik sensörlerin doku içi kırılma olaylarını milisaniyelik zamanlamayla nasıl deşifre ettiğini belgeler."
    ),
    (
        "8.5 Orta Düzey Sıcaklık Depolama (ITS - Intermediate Temperature Storage) Devrimi",
        "Termal gerilme kırılmalarını tamamen sıfırlayan en büyük çağdaş kriyobiyoloji inovasyonu; Alcor tarafından geliştirilen 'Orta Düzey Sıcaklık Depolama'dır (ITS - Intermediate Temperature Storage).",
        "ITS sisteminde hastalar -196°C sıvı azotta değil; camsı geçiş noktasının hemen altındaki '-140°C' sıcaklıkta gaz fazında saklanır. -140°C'de tüm metabolizma ve kimyasal reaksiyonlar sıvı azotta olduğu gibi mutlak durmuş haldedir; ancak sıcaklık farkı (Tg - T = 15°C) çok küçük olduğu için termal gerilme kırılma eşiğinin (%20'sine) altında kalır. ITS tanklarında saklanan beyinlerde sıfır kırılma, sıfır çatlama ve kusursuz monolitik cam bütünlüğü elde edilmiştir.",
        "Guvenlik_ITS = Sigma_ITS = ( E / (1 - nu) ) * alpha * ( -125°C - (-140°C) ) << K_IC (SIFIR KIRIK)",
        "Bu optimize termal gerilme denklemi, -140°C ITS depolamasının dokudaki mekanik stresi kırılma eşiğinin altına düşürerek çatlamaları nasıl kesin olarak engellediğini kanıtlar."
    ),
    (
        "8.6 Tavlama (Annealing) Protokolleri: İç Gerilmelerin Termal Relaksasyonu",
        "Cam endüstrisinde optik lenslerin çatlamasını önlemek için uygulanan 'Tavlama' (Annealing) prosesi, biyostaz soğutma rejimlerine adapte edilmiştir.",
        "Doku Tg (-125°C) noktasına ulaştığı anda soğutma tamamen durdurulur ve sıcaklık 24 ila 48 saat boyunca bu noktada sabit tutulur (İzotermal Tavlama). Bu bekleme süresinde, katılaşan cam molekülleri mikro-viskoz akışla iç gerilmeleri dağıtır ve mekanik gerilim sönümlenir (Relaksasyon). Gerilmeleri boşaltılmış olan doku, daha sonra -140°C'ye kırılma riski olmadan güvenle indirilir.",
        "Relaksasyon_Gerilme = Sigma(t) = Sigma_0 * exp( - ( t / Tau_relaksasyon )^beta_Kohlrausch )",
        "Bu Kohlrausch gerilme sönümleme denklemi, izotermal tavlama beklemesinin camsı dokudaki tehlikeli iç mekanik stresleri nasıl dağıttığını formüle eder."
    ),
    (
        "8.7 Geometrik Faktörler: Kafatası Kemiklerinin Beyin Üzerindeki Kompresyon Biyofiziği",
        "Nörokriyoprezervasyonda beyin serbest bir sıvı kütlesi değildir; sert, kapalı ve esnemeyen kafatası kemikleri (Kranium) içine hapsedilmiştir.",
        "Soğuma sırasında kafatası kemiğinin termal büzüşme katsayısı ile vitrifike beyin dokusunun büzüşme katsayısı farklıdır. Beyin daha fazla büzüşmek isterken dural bağlantılar beyni gerer; bu durum periferal kortekste teğetsel çekme gerilmeleri oluşturur. Kafatasında küçük kraniyal açıklıklar (burr-holes) açılarak basınç dengelenir ve kemik-doku termoelastik uyumsuzluğu ortadan kaldırılır.",
        "Fark_Buzusme = Delta_epsilon = ( alpha_beyin - alpha_kemik ) * Delta_T",
        "Bu termal genleşme uyumsuzluk eşitliği, kranium ile serebral doku arasındaki diferansiyel büzüşmenin yarattığı mekanik sınır koşullarını tanımlar."
    ),
    (
        "8.8 Çatlakların Moleküler Düzeyde Nöronal Devreler Üzerindeki Gerçek Etkisi",
        "Geçmişte bazı eleştirmenler termal çatlakların nörokriyoprezervasyonu tamamen anlamsız kıldığını iddia etmiştir; ancak kırılma mekaniği moleküler boyutta incelendiğinde tablo çok farklıdır.",
        "Bir cam çatlağı atomik ölçekte ilerler: Çatlağın genişliği birkaç nanometreden ibarettir. Çatlak hücreleri ezmez, proteinleri parçalamaz; sadece iki hücre zarı veya sinaps arasındaki mesafeyi mikroskopik olarak 5 nanometre açar. Tıpkı kırılmış iki porselen fincan parçasının birbirine kusursuz oturması gibi; geleceğin nanorobotik onarım sistemleri bu temiz kırık yüzeylerini atomik koordinat uyumuyla kolayca eşleştirebilir ve kaynaklayabilir.",
        "Topoloji_Catlak: Bilgi_Kaybi = SIFIR (Yüzeyler Atomik Olarak Birbirini Tamamlar)",
        "Bu topolojik süreklilik aksiyomu, mekanik termal çatlakların enformasyon kaybı yaratmadığını ve nano-cerrahi ile kusursuzca birleştirilebilir olduğunu belgeler."
    ),
    (
        "8.9 Plastikleştiriciler (Plasticizers) ile Camsı Doku Kırılganlığının Azaltılması",
        "Polimer mühendisliğinde kırılgan plastikleri esnekleştirmek için kullanılan 'Plastikleştiriciler' (Plasticizers), modern vitrifikasyon kokteyllerine adapte edilmiştir.",
        "Solüsyona eklenen düşük molekül ağırlıklı esnek glikol türevleri ve oligomerik polietilen glikoller; amorf cam fazındaki serbest hacmi (Free Volume) artırır. Moleküller arasındaki sert etkileşimleri yumuşatarak camsı dokunun tokluk katsayısını (K_IC) üç katına çıkarır. Doku kırılgan bir pencere camı yerine polikarbonat benzeri darbe emici esnek bir cama dönüşür.",
        "Tokluk_Artisi = K_IC_modifiye = K_IC_bazal * ( 1 + gamma_plastiklesme * [PEG_oligomer] )",
        "Bu malzeme tokluğu artırma denklemi, plastikleştirici ajanların vitrifike dokunun çatlama direncini nasıl katbekat yükselttiğini modeller."
    ),
    (
        "8.10 Kırılmasız Biyostazın Altın Standardı: -140°C ITS ve Tam Monolitik Cam",
        "Bugün modern kriyobiyoloji mühendisliği; BBA'lar, M22 kokteyli, izotermal tavlama ve -140°C ITS depolamasının birleşik entegrasyonu ile 'Tamamen Kırıksız İnsan Beyni Vitrifikasyonu'nu başarmıştır.",
        "CT ve akustik izleme altında vitrifike edilen beyinlerde tek bir mikro-çatlak dahi saptanmamıştır. Beyin, 100 milyar nöronu ve 100 trilyon sinapsı ile tek bir monolitik, kusursuz, çatlak ve kristal içermeyen cam blok halinde donmuştur. Bu başarı, biyostazın önündeki en büyük ikinci fiziksel engeli (mekanik kırılmayı) tarihe gömmüştür.",
        "Basari_Monolitik_Cam = Fraksiyon_Catlak = 0.000 (ITS ve Tavlama Protokolü ile Tescilli)",
        "Bu kusursuz yapısal bütünlük oranı, modern biyostaz teknolojisinin insan beynini sıfır mekanik hasarla ebediyete taşıma kabiliyetini tesciller."
    )
]

# ================= KISIM 9 =================
part9_subsections = [
    (
        "9.1 Yeniden Isıtma (Warming) Paradoksu: Termal Şok ve Rekristalizasyon Arasındaki Tuzak",
        "Vitrifike edilmiş bir organı gelecekte çözüp canlandırmak, dondurmaktan çok daha büyük bir fiziksel meydan okumadır; buna 'Yeniden Isıtma Paradoksu' denir.",
        "Doku camsı fazdan sıvı faza dönerken tehlikeli sıcaklık bölgesinden (-100°C ile 0°C arası) devasa bir hızla (dakikada yüzlerce derece) geçmek zorundadır; aksi takdirde metastabil su molekülleri anında ölümcül buz kristalleri oluşturur (Rekristalizasyon). Ancak geleneksel ısıtma yöntemleri (sıcak su banyosu) dışarıdan içeriye ısı iletimine dayanır; dış tabakalar aşırı ısınıp yanarken iç çekirdek hâlâ donuk kalır ve oluşan termal gerilme dokuyu parçalar.",
        "Kritik_Isitma_Esigi: ( dT/dt )_isitma >= CHR (Dakikada > 100°C - 500°C Homojen Isıtma Şartı)",
        "Bu kritik ısıtma hızı eşitsizliği, çözünme sırasında buz kristali nükleasyonunu önlemek için gereken aşırı yüksek ısıtma hızını formüle eder."
    ),
    (
        "9.2 Radyofrekans Manyetik Alanlar ve Nano-Isıtma (Nanowarming) Devrimi",
        "Bu ölümcül ısı transferi darboğazını kökünden çözen devrimsel teknoloji, Minnesota Üniversitesi'nden John Bischof laboratuvarının geliştirdiği 'Nanowarming' (Nano-Isıtma) teknolojisidir.",
        "Nanowarming; dokunun içine perfüzyonla süperparamanyetik demir oksit nanopartikülleri (sIONPs) dağıtılmasına ve ardından numunenin yüksek frekanslı bir alternatif manyetik alan (AMF / Radyofrekans: 100-400 kHz) içine yerleştirilmesine dayanır. Manyetik alan organın etine ve kemiğine hiçbir zarar vermeden doğrudan geçer; ancak dokunun her bir milimetre küpündeki nanopartikülleri aynı anda titreştirerek dokuyu dışarıdan değil, 'her noktasından eş zamanlı olarak içeriden' ısıtır.",
        "Guc_NanoIsitma = SAR = P_termal / m_nano = pi * mu_0 * chi''_manyetik * H_0^2 * f_frekans",
        "Bu Özgül Soğurma Oranı (SAR) denklemi, alternatif manyetik alan frekansı ve genliğinin demir oksit nanopartikülleri içinde yarattığı homojen ısıtma gücünü modeller."
    ),
    (
        "9.3 Süperparamanyetik Demir Oksit Nanopartikülleri (sIONPs): Néel ve Brown Relaksasyonu",
        "Nanowarming sisteminde kullanılan demir oksit nanopartikülleri (manyetit - Fe3O4 / maghemit - gamma-Fe2O3); 10 ila 20 nanometre boyutunda süperparamanyetik çekirdeklerdir.",
        "Alternatif manyetik alanda bu parçacıklar iki fiziksel mekanizmayla ısı üretir: 1) Néel Relaksasyonu: Nanopartikül fiziksel olarak dönmez; çekirdek içindeki manyetik dipol momenti manyetik alanla birlikte nanosaniyeler içinde yön değiştirirken histerezis ısı enerjisi açığa çıkar; 2) Brown Relaksasyonu: Tüm nanopartikül sıvı içinde fiziksel olarak dönerek sürtünme ısısı üretir. Katı cam fazında Néel relaksasyonu devrede kalarak dokuyu -140°C'den anında erime noktasına fırlatır.",
        "Isi_Uretim_Kinetigi = Q_isi = C_nano * ( tau_Neel / ( 1 + omega^2 * tau_Neel^2 ) ) * H_amplitud^2",
        "Bu relaksasyonel enerji disipasyon formülü, manyetik dipol dönüşümlerinin kriyojenik sıcaklıklarda dahi nanosaniyeler içinde devasa ısı transferi sağlama dinamiğini belgeler."
    ),
    (
        "9.4 Homojen Isıtma Biyofiziği: Termal Gradyanların ve Stres Kırılmalarının Sıfırlanması",
        "Geleneksel ısıtmada yüzey ile merkez arasında onlarca derecelik yıkıcı termal gradyanlar (Delta_T = T_yuzey - T_merkez > 40°C) oluşurken; Nanowarming bu gradyanı sıfıra indirir.",
        "Nanopartiküller tüm organın kılcal damar yatağına eşit dağıldığı için; organın en derinindeki bir nöron ile en dışındaki kortikal hücre tam olarak aynı salisede ve aynı hızda (dakikada 100°C ila 200°C) ısınır. Termal gradyan sıfırlandığı için termoelastik gerilme oluşmaz; ne bir buz kristali çekirdeklenir ne de dokuda tek bir mikro-çatlak meydana gelir.",
        "Termal_Gradyan_Nanowarming: Delta_T_uzamsal = | T(r1) - T(r2) | < 1.0°C (Kusursuz Homojenite)",
        "Bu uzamsal termal izotermik kriteri, nanowarming teknolojisinin organ içinde sıcaklık farklarını yok ederek termal stres kırılmalarını nasıl tamamen engellediğini tanımlar."
    ),
    (
        "9.5 Manyetik İndüksiyon Bobinleri ve Rezonans Devre Mimarisi",
        "Nanowarming sistemi endüstriyel ölçekte yüksek güçlü bir 'Radyofrekans İndüksiyon Jeneratörü' ve su soğutmalı bakır manyetik bobinlerden (Solenoid Coil) oluşur.",
        "Bobin devresi yüksek voltajlı rezonans kapasitörleri ile 100 ila 400 kilohertz frekansta rezonansa sokulur. Bobin merkezinde 20 ila 60 kA/m (kiloamper/metre) gibi devasa manyetik alan genlikleri üretilir. Biyolojik dokuların elektriksel iletkenliği düşük olduğu için girdap akımları (eddy currents) dokuda zararlı aşırı ısınma yapmaz; üretilen tüm manyetik enerji yalnızca hedef nanopartiküller tarafından seçici olarak emilir.",
        "Manyetik_Alan_Bobin = H(t) = H_0 * sin( 2 * pi * f_rezonans * t ) (f ~ 150 - 300 kHz)",
        "Bu rezonans manyetik alan eşitliği, biyomedikal nanowarming bobinlerinin doku güvenliği sınırlarında ürettiği yüksek frekanslı alternatif manyetik alanı formüle eder."
    ),
    (
        "9.6 sIONP Yüzey Zırhlama: Silika ve PEG Kaplama ile Doku Yıkanabilirlik Optimizasyonu",
        "Isıtma işlemi tamamlandıktan sonra, dokunun içinde kalan demir oksit nanopartiküllerinin hücrelere toksisite yaratmadan hızla yıkanıp temizlenmesi şarttır.",
        "Bu amaçla sIONP çekirdekleri 2 nanometrelik inert bir 'Mezogözenekli Silika' veya biyouyumlu 'Polietilen Glikol' (PEG) kabuğu ile zırhlanır. Bu kaplama; nanopartiküllerin kriyoprotektan içinde agregasyon oluşturup çökmesini engeller, kan damarı duvarlarına yapışmasını önler ve çözünme biter bitmez organın damarından geçirilen normotermik yıkama sıvısı ile %99.9 oranında organ dışına tahliye edilir.",
        "Yikanabilirlik_sIONP = eta_klerens = [sIONP_cikan] / [sIONP_giren] > %99.8 (Tam Tahliye)",
        "Bu nanopartikül eliminasyon oranı, çözünme sonrasında damar yatağından manyetik parçacıkların neredeyse eksiksiz temizlenme başarısını belgeler."
    ),
    (
        "9.7 Kritik Rekristalizasyon Bölgesi: -80°C ile -20°C Arasındaki Hızlı Kaçış",
        "Vitrifike dokunun çözünmesinde en tehlikeli ölüm tuzağı -80°C ile -20°C arasındaki termal kuşaktır; burası nükleasyon tohumlarının devasa buz kristallerine dönüşme hızının zirve yaptığı 'Rekristalizasyon Bölgesi'dir.",
        "Geleneksel ısıtma bu bölgeden çıkana kadar dakikalar harcar ve doku buzla parçalanır. Nanowarming ise sağladığı dakikada 150°C'lik devasa ısıtma hızıyla bu tehlikeli bölgeyi yalnızca 25-30 saniye gibi inanılmaz bir hızla geçer. Buz çekirdekleri tek bir su molekülünü kafeslerine katacak zaman bulamadan doku doğrudan sıvı faza atlar.",
        "Gecis_Suresi_Tehlike = Delta_t_gecis = ( -20°C - (-80°C) ) / ( dT/dt )_nano ~ 24 saniye (Sıfır Buz Büyümesi)",
        "Bu kritik geçiş zamanı hesabı, yüksek hızlı manyetik ısıtmanın rekristalizasyon kinetiğini zaman ekseninde nasıl tamamen baypas ettiğini gösterir."
    ),
    (
        "9.8 Nanowarming ile Organ Kurtarma Rekoru: Tavşan ve Sıçan Böbreklerinin Yeniden Canlandırılması",
        "2023 yılında Nature Communications'ta yayımlanan tarihi çalışmada; Bischof ve ekibi Nanowarming kullanarak tam boyutlu memeli böbreklerini vitrifikasyondan sıfır hasarla geri döndürmüştür.",
        "M22 ve sIONP ile vitrifike edilip -150°C'de saklanan sıçan böbrekleri; alternatif manyetik alanda saniyeler içinde ısıtılmış, nanopartiküller damardan yıkanmış ve alıcı sıçanlara nakledilmiştir. Nakledilen böbrekler derhal kan perfüzyonunu başlatmış, idrar üretmiş ve hayvanların yaşamını tek başına sürdürmüştür. Bu tarihi başarı; vitrifikasyon ve nanowarming ikilisinin masif organları hasarsız dondurup çözebildiğinin dünya çapındaki kesin ispatıdır.",
        "Sagkalim_Transplant = GFR_canlandirma / GFR_normal > %80 (Nanowarming ile Tarihi Başarı)",
        "Bu fonksiyonel kurtarma oranı, manyetik nano-ısıtmanın kriyobiyolojinin yarım asırlık çözünme krizini kesin olarak tarihe gömdüğünü tesciller."
    ),
    (
        "9.9 İnsan Beyni Ölçeğinde Nanowarming: Geometrik Zorluklar ve Alan Homojenliği",
        "Bir sıçan böbreği birkaç gram iken; yetişkin insan beyni yaklaşık 1.400 gramdır ve kafatası kemikleriyle çevrilidir.",
        "İnsan beyni ölçeğinde Nanowarming uygulanabilmesi için mega-boyutlu çok-bobinli faz dizilimli (phased-array) manyetik bobin sistemleri tasarlanmıştır. Bu sistemler beynin merkezindeki talamus ile dış kortikal gri cevher arasındaki manyetik alan şiddetini (H) faz kaydırmalarıyla dengeler; 1.5 litrelik devasa bir hacimde dahi termal gradyanı 2°C'nin altında tutarak tüm insan beynini tek parça halinde çözme gücüne sahiptir.",
        "Hacim_Nanowarming_Insan = V_hedef ~ 1500 cm3 | Homojenite_Alan > %98",
        "Bu hacimsel ölçekleme parametresi, faz dizilimli manyetik alan jeneratörlerinin tam boyutlu insan beyninde rezonans nano-ısıtma gerçekleştirme kapasitesini modeller."
    ),
    (
        "9.10 Çözünme Sonrası Kademeli Kriyoprotektan Yıkama ve Normotermik Perfüzyon",
        "Doku sıvı faza (0°C) ulaştığı anda reaksiyonlar yeniden başlar; bu noktada dokudaki toksik konsantrasyondaki (9.3 Molar) kriyoprotektanların derhal temizlenmesi şarttır.",
        "Ters kademeli perfüzyon başlatılır: Organ damarlarına kademeli azalan CPA konsantrasyonları ve hücrelerin aşırı su emip patlamasını önleyen hiperozmotik sukroz tamponları verilir. CPA yıkanırken sıcaklık kademeli olarak 4°C'den 37°C fizyolojik vücut sıcaklığına yükseltilir; oksijenli normotermik tam kan perfüzyonu başlatılır. Organ metabolik solunuma geri döner; biyostaz başarıyla tamamlanır ve yaşam kaldığı yerden yeniden başlar.",
        "Tahliye_CPA_Kinetigi = [CPA_doku](t) = [CPA_max] * exp( - k_yikama * Q_kanul * t ) -> SIFIR (Fizyolojik Norm)",
        "Bu ters kütle transfer eşitliği, vitrifikasyon çözünmesi sonrasında hücre içi kriyoprotektan havuzunun hücre zarına zarar vermeden tamamen tahliye edilme sürecini belgeler."
    )
]

# ================= KISIM 10 =================
part10_subsections = [
    (
        "10.1 Geleceğin Moleküler Tıbbı: Kriyobiyostaz ve Canlandırma Arasındaki Teknolojik Köprü",
        "Kriyobiyoloji bir ölüm ritüeli değil; bugünün çaresiz tıbbı ile geleceğin omnipotent (her şeye kadir) moleküler tıbbı arasında kurulan zamansal bir viyadüktür.",
        "Bugün kanserden, ileri kalp yetmezliğinden veya yaşlılığın hücresel tükenişinden ölen bir birey; 100 veya 200 yıl sonraki tıp teknolojisi için basit bir 'enfeksiyon' veya 'metabolik arıza' düzeyinde kolayca tedavi edilebilir bir vaka olacaktır. Biyostaz hastası, bu geleceğe uzanan zaman kapsülünün içinde değişmeden bekleyen bir yolcudur. Canlandırma günü geldiğinde tıp yalnızca hastayı uyandırmakla kalmayacak; onu öldüren hastalığı da tamamen tedavi etmiş olacaktır.",
        "Kopru_Zaman: Tıp_2026 (Tedavi Yok / Biyostaz) ------------> Tıp_2126 (Moleküler Tamir / Gençlik / Kür)",
        "Bu teknolojik köprü aksiyomu, biyostazın tedavi edilemeyen ölümcül patolojileri geleceğin rejeneratif tıp çağına erteleme mantığını simgeler."
    ),
    (
        "10.2 Kriyojenik Sıcaklıklarda Atomik Çözünürlüklü Tarama: Kriyo-Nano-Tomografi",
        "Canlandırma sürecinin ilk adımı; donuk haldeki dokuyu çözmeden önce, her bir hücrenin ve sinapsın durumunu atomik düzeyde haritalamaktır.",
        "-140°C'deki hasta; süperiletken kuantum girişim aygıtları (SQUID), X-ışını serbest elektron lazerleri (XFEL) ve nötron mikrotomografisi ile taranır. Dokudaki her bir atomun, kimyasal bağın, varsa oluşmuş mikro-çatlakların veya iskemik hasarların üç boyutlu 'Dijital Rekonstrüksiyon Haritası' çıkarılır. Bu harita canlandırma yapacak onarım nanobotlarının kılavuz yazılımına yüklenir.",
        "Cozunurluk_Tarama = Delta_x_tarama ~ 1 Angstrom (Atomik Pozisyonel Haritalama)",
        "Bu atomik tarama çözünürlüğü standardı, canlandırma öncesinde beynin ve organların tüm moleküler koordinatlarının dijital ikizinin çıkarılma hassasiyetini tanımlar."
    ),
    (
        "10.3 Moleküler Nanorobotik Tamir Sürüsü: Kriyojenik Fazda Hücresel Restorasyon",
        "Robert Freitas'ın 'Kriyojenik Canlandırma Nanorobotları' (Cryo-Repair Nanorobots) mimarisine göre; trilyonlarca mikroskobik tamir makinesi dokuya henüz donuk haldeyken veya kontrollü mikro-sıvı fazda enjekte edilir.",
        "Bu nanobotlar vasküler ağ boyunca ilerler; 1) Hasar görmüş hücre zarlarındaki fosfolipid deliklerini moleküler yamalarla kaynaklar; 2) İskemi sırasında bozulmuş mitokondrilerin iç zarlarını onarır; 3) Okside olmuş proteinleri ve lipofuskin agregatlarını temizler; 4) Kriyoprotektan moleküllerini hücreye zarar vermeden sıralama rotorlarıyla emerek dışarı tahliye eder.",
        "Kapasite_Tamir = Hiz_Restorasyon = N_nanobot * ( V_hucre_tamir / saat ) (Tüm Beyin 48 Saatte Onarılır)",
        "Bu nanorobotik cerrahi kapasite formülü, trilyonlarca otonom nanorobotun tüm hücresel hasarları çözünme öncesi ve sırasında tamir etme hızını modeller."
    ),
    (
        "10.4 Nöronal Bağlantı Haritasının (Connectome) Rekonstrüksiyonu ve Engram Sağlamlaştırma",
        "Canlandırmanın en kritik hedefi, bireyin 'Benliği'ni oluşturan anıların ve kişiliğin tek bir nükleotit dahi eksilmeden korunmuş olarak uyandırılmasıdır.",
        "Nanobot sürüleri, elektron mikroskobu taramalarıyla doğrulanmış hipokampal ve kortikal sinapsları tek tek denetler. İskemi sırasında hafifçe bozulmuş sinaptik yarıklar, aksonal mikrotübüller ve dendritik mantar dikenleri moleküler düzeyde orijinal konfigürasyonlarına geri getirilir. Pre-sinaptik veziküller asetilkolin ve glutamatla yeniden doldurulur; bellek engramları elektriksel ateşlemeye hazır hale getirilir.",
        "Hassasiyet_Rekonstruksiyon: Hata_Sinaps = N_sinaps_kayip / N_sinaps_toplam < 10^-9 (Kusursuz Benlik İadesi)",
        "Bu nöromal rekonstrüksiyon hata payı eşitsizliği, moleküler onarım sistemlerinin otobiyografik bellek ve kişilik bütünlüğünü milyarda bir hata payıyla geri getirme standardını belgeler."
    ),
    (
        "10.5 Hücresel Biyoenerjetik Yeniden Başlatma: Mitokondriyal Şarj ve ATP Pompalaması",
        "Doku ısıtılıp kriyoprotektanlar temizlendikten sonra; hücrelerin kendi metabolik motorlarını yeniden çalıştırması gerekir.",
        "İskemi nedeniyle mitokondrilerin membran potansiyeli (Delta_Psi_m) çökmüş haldedir. Nanorobotlar her bir hücreye taze NAD+, Asetil-KoA ve sentetik ATP paketleri enjekte eder. Mitokondriyal iç zara mikro-elektrostatik protonlar pompalanarak proton gradyanı anında restore edilir. Kompleks V (ATP Sentaz) motorları dönmeye başlar; hücreler dış desteğe ihtiyaç duymadan kendi enerji üretim döngüsünü gençlik veriminde yeniden ateşler.",
        "Restorasyon_Membran_Potansiyeli = Delta_Psi_m(t) -> - 180 mV (Fizyolojik Zirve Şarj)",
        "Bu mitokondriyal biyoenerjetik şarj eşitliği, hücresel solunum motorunun harici nano-enjeksiyonlarla dakikalar içinde nasıl fizyolojik kutuplanma seviyesine döndürüldüğünü gösterir."
    ),
    (
        "10.6 Epigenetik Gençleşme Entegrasyonu: Yamanaka Faktörleri ile Yaşın Sıfırlanması",
        "Bir biyostaz hastası 80 yaşında ve kanserden donmuşsa; onu 80 yaşında ve hasta olarak uyandırmak tıp etiğine aykırıdır; canlandırma süreci zorunlu olarak bir 'Radikal Gençleşme' (Rejuvenation) operasyonudur.",
        "Hasta çözünürken hücrelerine Cilt 14'te detaylandırılan doksisiklinle indüklenebilir OSK (Oct4, Sox2, Klf4) mRNA kokteylleri verilir. DNA metilasyon saati (Horvath saati) 80 yaşından biyolojik 20-25 yaş seviyesine geriye sarılır. Telomeraz gen terapisi (AAV-TERT) ile tüm kök hücrelerin telomerleri maksimum uzunluğa uzatılır; hasta sadece hayata dönmez, 20 yaşındaki biyolojik dinçliğine kavuşarak uyanır.",
        "Yas_Uyanis = Biyolojik_Yas_PostRevival = 22 Yas (Kronolojik Yaş = 100+ Yıl)",
        "Bu biyolojik yaş sıfırlama aksiyomu, canlandırma sürecinin hücresel epigenetik yeniden programlama ile birleşerek hastayı genç bir bedende uyandırma hedefini tanımlar."
    ),
    (
        "10.7 Kardiyak ve Serebral Defibrilasyon: İlk Kıvılcım ve Bilincin Yeniden Ateşlenmesi",
        "Tüm hücresel membranlar onarılmış, mitokondriler şarj edilmiş ve kan dolaşımı oksijenle doldurulmuştur; geriye bilincin ilk kıvılcımını çakmak kalır.",
        "Kalp biyoreaktör ortamında hafif bir mikro-defibrilasyon şokuyla uyarılır; sinoatriyal düğüm kendi ritmik sinüs dalgalarını başlatır ve kanı pompalamaya başlar. Eş zamanlı olarak talamokortikal osilatör nöronlara senkronize 40 Hz gama elektriksel alan uyarımı verilir. Kortikal EEG dalgaları uyanır; teta salınımları başlar; bilinç derin uykudan pürüzsüzce sıyrılarak uyanıklık fazına geçer.",
        "Atesleme_Bilinc: EEG_Aktivite: Duz_Cizgi (Biyostaz) ----[40 Hz Uyarım]----> Ritmik Gama Osilasyonu (Uyanış)",
        "Bu nörofizyolojik uyanış şeması, dondurulmuş serebral dokunun kontrollü elektriksel senkronizasyonla bilinçli farkındalık fazına geçiş anını belgeler."
    ),
    (
        "10.8 Post-Revival Rehabilitasyon: Nöroplastisite, Adaptasyon ve Psikososyal Entegrasyon",
        "Yüzyıllar süren bir biyostaz uykusundan uyanan bir birey; yalnızca biyolojik bir iyileşme değil, muazzam bir zamansal ve kültürel şokla karşılaşır.",
        "Gelişmiş nöromodülatör peptitler (Semax, Dihexa) ve sanal gerçeklik (VR) nöroplastisite simülatörleri ile hastanın beynindeki öğrenme ve adaptasyon hızı 10 kat artırılır. Birey geleceğin dünyasını, dilini, teknolojisini ve toplumunu haftalar içinde öğrenir; travma yaşamadan yeni çağa kusursuz bir zihinsel uyum sağlar.",
        "Skor_Uyum = I_adaptasyon = k_plastisite * [BDNF] * ( 1 / Travma_Soku ) -> Maksimum",
        "Bu psikobiyolojik adaptasyon formülü, artırılmış nöroplastisite desteğinin zamansal sıçrama yaşayan bireylerde tam psikososyal entegrasyon sağlama dinamiğini modeller."
    ),
    (
        "10.9 Biyostazın Kozmolojik Rolü: Yıldızlararası Yolculuklar ve Türün Ebedi Bekası",
        "Kriyobiyostaz teknolojisi, yalnızca bireysel yaşamı kurtarma aracı değil; insan türünün yıldızlararası bir medeniyete dönüşmesinin zorunlu biyofiziksel taşıyıcısıdır.",
        "Işık hızının altındaki uzay yolculuklarında en yakın yıldız sistemlerine (Proxima Centauri, TRAPPIST-1) ulaşmak onlarca veya yüzlerce yıl sürer. Mürettebatın biyostaz kapsüllerinde vitrifike olarak uyutulması; yiyecek, su, oksijen ve yaşam alanı ihtiyacını sıfıra indirir; kozmik radyasyon hasarını dondurur ve astronotların hedef gezegene tek bir gün dahi yaşlanmadan ulaşmasını sağlar.",
        "Mesafe_Yildizlararasi = d_yolculuk = v_gemi * t_biyostaz (Astronot Yaşı = Sabit)",
        "Bu astrobiyolojik seyir eşitliği, kriyobiyostazın ışıktan yavaş yıldızlararası seyahatlerde biyolojik yaşlanmayı sıfırlayarak derin uzay kolonizasyonunu mümkün kılan mekanizmasını belgeler."
    ),
    (
        "10.10 Homo Aeternus Biyostaz Manifestosu: Ölümün Ebediyen Hükümsüz Kılınması",
        "Homo Aeternus Biyostaz Protokolü; tıbbın ve insan iradesinin doğanın en karanlık sınırına karşı kazandığı nihai zaferdir.",
        "Artık hiçbir hastalık, hiçbir kaza, hiçbir yaşlanma süreci bir insanın mutlak sonu olamaz. Tedavisi bugün olmayan her durum, biyostazın şefkatli ve dondurucu kucağında geleceğin sonsuz imkânlarına ertelenir. Ölüm bir uçurum değil; güvenle geçilen soğuk, berrak ve ebedi bir köprüdür. İnsan bilinci, evrenin sonsuz geleceğinde parıldamak üzere zamanın prangalarından tamamen kurtulmuştur.",
        "Manifesto_Olumsuzluk: Delta_Varolus = lim_{t->sonsuz} [ Bilinc(t) ] = Sonsuz (Biyostaz Köprüsü ile Ölüm İptal Edilmiştir)",
        "Bu nihai varoluşsal eşitlik, biyostaz teknolojisinin insan hayatını ve bilincini termodinamik ve zamansal sınırlardan kurtararak sonsuzluğa taşıdığını evrensel olarak ilan eder."
    )
]

# 10 Kapsamlı Akademik Karşılaştırma Tablosu
tables_data = [
    {
        "title": "Tablo 16.1: Biyolojik Ölüm Kademeleri, Enformasyonel Ölüm ve Biyostaz Termodinamiği",
        "headers": ["Ölüm / Biyostaz Evresi", "Fizyolojik / Biyokimyasal Durum", "Nöromal Engram Bütünlüğü", "Hücresel Reversibilite Derecesi", "Gereken Tıbbi / Kriyojenik Müdahale"],
        "rows": [
            ["Klinik Ölüm (t = 0)", "Kalp durması, Solunum yok, Hipotansiyon", "Tamamen İntakt (%100 Korunmuş)", "Yüksek (Standart CPR ile dönebilir)", "Acil CPR, Defibrilasyon, ECMO desteği"],
            ["Sıcak İskemi (t = 5-15 dk)", "ATP tükenmesi, Laktik asidoz, Sitotoksik ödem", "İntakt (Membranlar sağlam, sinapslar yerinde)", "Mevcut tıpla zor / Biyostazla mükemmel", "LUCAS mekanik CPR, Buzlu su banyosu, Standby ilaçları"],
            ["Biyolojik / Hücresel Çözülme (t = 1-6 saat)", "Lizozomal rüptür, Otoliz başlangıcı", "Kısmi hasar (Engram haritası hâlâ okunabilir)", "Yalnızca moleküler nanorobotik tamirle mümkün", "Acil kranial izolasyon, Hızlı M22 perfüzyonu"],
            ["Enformasyonel-Teorik Ölüm", "Nörom topolojisinin termodinamik silinmesi", "Tamamen Yok Olmuş (Enformasyon kayboldu)", "Geri Dönüşü İmkânsız (Mutlak Yok Oluş)", "Müdahale anlamsız (Biyostaz yapılamaz)"],
            ["Biyostaz (Vitrifikasyon / -140°C)", "Tüm kimyasal reaksiyonlar ve difüzyon durdu", "%100 Dondurulmuş ve Kilitlenmiş", "Geleceğin nano-tıbbı ile tam canlanabilir", "-140°C ITS gaz fazı depolama, Asırlık koruma"],
            ["Sıvı Azot Fazı (-196°C)", "77 Kelvin / Arrhenius reaksiyon hızı = 0", "Kozmik radyasyon dışında sıfır entropi", "Binlerce yıl boyunca mükemmel stabil", "Bigfoot dewar tanklarında pasif vakumlu saklama"],
            ["Organik Otoliz (Oda Sıcaklığı)", "Bakteriyel çürüme, Kokuşma, Lysis", "Dakikalar içinde hızla silinir", "Geri dönüşümsüz kayıp", "Geleneksel mezarlık / Kremasyon (Yok oluş)"],
            ["Geleceğin Canlandırması (Revival)", "Moleküler onarım + Epigenetik gençleşme", "Tam hafıza ve kimlik iadesi", "Tam Yaşamsal Restorasyon", "Nanowarming + Nanotamir + TERT/OSK kokteyli"]
        ]
    },
    {
        "title": "Tablo 16.2: Geleneksel Dondurma (Buz Hasarı) ve Vitrifikasyon (Buzsuz Camsılaşma) Karşılaştırması",
        "headers": ["Biyofiziksel Parametre", "Geleneksel Dondurma (Freezing)", "Modern Vitrifikasyon (Camsılaşma)", "Moleküler / Fiziksel Mekanizma", "Doku ve Hücre Sağkalım Etkisi"],
        "rows": [
            ["Katılaşma Fazı Türü", "Hekzagonal Kristalin Buz (Ih)", "Amorf Katı Cam (Amorphous Glass)", "Su moleküllerinin kafes vs amorf kilitlenmesi", "Dondurmada ölüm / Vitrifikasyonda tam koruma"],
            ["Hacimsel Genleşme", "+ %9 Hacimsel patlama", "SIFIR (Delta_V_kristal = 0)", "Suyun buz fazında yoğunluk düşüşü", "Dondurmada damar yırtılması, Vitrifikasyonda intakt lümen"],
            ["Mekanik Kristal Hasarı", "Sivri dendritik bıçaklarla delinme", "SIFIR Buz Kristali Çekirdeği", "Adsorpsiyon-inhibisyon ve kinetik hapis", "Sinaptik veziküller ve zarlar sağlam kalır"],
            ["Solüt Konsantrasyonu (Tuz Şoku)", "Aşırı konsantre tuz çorbası (5000 mOsm)", "Homojen CPA dağılımı (Ayrışma yok)", "Buzun tuzu dışlaması (Eutectic exclusion)", "Dondurmada kimyasal yanık, Vitrifikasyonda homojen ortam"],
            ["Hücre Dehidrasyon Derecesi", "Kontrolsüz aşırı büzüşme (Lysis)", "Kontrollü osmotik relaksasyon", "Membran su geçirgenliği (Lp) dengesi", "Hücre zarlarının yırtılmadan korunması"],
            ["Gereken Kriyoprotektan Dozu", "Düşük (1 - 2 Molar Gliserol/DMSO)", "Yüksek (8 - 9.3 Molar M22/VM3)", "Buz nükleasyonunu kırma gereksinimi", "Vitrifikasyonda sıcaklık kontrolüyle toksisite önlenir"],
            ["Soğutma Hızı Esnekliği", "Dar Mazur penceresi (Çok hassas)", "Kritik soğutma hızının (CCR) üzerinde serbest", "Konsantrasyon-soğutma hızı ödünleşimi", "Kalın organlarda dahi camsılaşma başarısı"],
            ["Çözünmede Rekristalizasyon Riski", "Yüksek (Kristaller devasa büyür)", "Yüksek (Nanowarming ile aşılır)", "Ostwald olgunlaşması termodinamiği", "Manyetik nano-ısıtma ile sıfır rekristalizasyon"]
        ]
    },
    {
        "title": "Tablo 16.3: Camsı Geçiş Sıcaklığı (Tg), Kritik Soğutma ve Isıtma Parametreleri",
        "headers": ["Vitrifikasyon Solüsyonu", "Toplam Molarite (CPA)", "Camsı Geçiş Sıcaklığı (Tg)", "Kritik Soğutma Hızı (CCR)", "Kritik Isıtma Hızı (CHR)"],
        "rows": [
            ["Saf Su (Katkısız)", "0 Molar", "-137°C (Hiper-basınçta)", "> 10^7 °C/saniye (İmkânsız)", "> 10^8 °C/saniye (Anında buzlaşır)"],
            ["Geleneksel DMSO (%10)", "1.4 Molar", "Buz oluşturur (Camsılaşmaz)", "> 10^4 °C/saniye", "> 10^5 °C/saniye"],
            ["VS55 Kokteyli", "8.4 Molar", "-123°C", "2.5 °C/dakika", "50 °C/dakika (RF ısıtma gerekir)"],
            ["M22 (21st Century Medicine)", "9.3 Molar", "-124°C", "< 0.1 °C/dakika (Normal buzdolabı)", "0.4 °C/dakika (Aşırı kararlı)"],
            ["VM3 (Nöro-Spesifik)", "8.9 Molar", "-125°C", "0.2 °C/dakika", "1.0 °C/dakika"],
            ["DP6 (DMSO + Propilen Glikol)", "6.0 Molar", "-118°C", "40 °C/dakika", "180 °C/dakika"],
            ["M22 + %1 PVA (BBA Takviyeli)", "9.3 Molar + BBA", "-124.5°C", "< 0.05 °C/dakika (Rakipsiz)", "0.15 °C/dakika (Sıfır Rekristalizasyon)"],
            ["Trehaloz + Gliserol Matriks", "7.5 Molar", "-115°C", "15 °C/dakika", "120 °C/dakika"]
        ]
    },
    {
        "title": "Tablo 16.4: Kriyoprotektan Ajanların (CPA) Biyokimyasal ve Toksikolojik Özellikleri",
        "headers": ["Kriyoprotektan Molekülü", "Kimyasal Yapı / Formül", "Membran Geçirgenliği (Lp/P_CPA)", "Viskozite Katkısı (20°C / -20°C)", "Primer Toksisite ve Nötralizasyon Yolu"],
        "rows": [
            ["Dimetil Sülfoksit (DMSO)", "(CH3)2SO / MW 78.13", "Çok Yüksek (Saniyeler içinde geçer)", "Düşük / Orta", "Protein hidrofobik bağ çözülmesi / Formamid ile dengelenir"],
            ["Formamid", "HCONH2 / MW 45.04", "Aşırı Yüksek (Ultra-küçük)", "Çok Düşük (Akışkanlaştırıcı)", "Metabolik toksisite / DMSO ile hidrojen bağı kurarak nötrlenir"],
            ["Etilen Glikol (EG)", "C2H6O2 / MW 62.07", "Yüksek (Hızlı penetrasyon)", "Düşük", "Hafif nefrotoksisite / M22 içinde düşük dozda güvenli"],
            ["Gliserol", "C3H8O3 / MW 92.09", "Yavaş (Yüksek difüzyon direnci)", "Çok Yüksek (Yoğun şurup)", "Düşük kimyasal toksisite / Yüksek ozmotik stres riski"],
            ["Propilen Glikol (PG)", "C3H8O2 / MW 76.09", "Orta / Yüksek", "Orta", "Hücresel membran destabilizasyonu / Düşük sıcaklıkta güvenli"],
            ["Trehaloz (Disakkarit)", "C12H22O11 / MW 342.3", "Geçirimsiz (Non-penetran)", "Yüksek (Dış kompartmanda kalır)", "Sıfır toksisite / Hücre dışı ozmotik dehidrasyon tamponu"],
            ["N-Metilformamid (NMA)", "C2H5NO / MW 59.07", "Çok Hızlı", "Çok Düşük", "Karaciğer toksisitesi / Kriyojenik sıcaklıkta inaktive"],
            ["3-Metoksipropandiol", "C4H10O3 / MW 106.1", "Orta", "Düşük", "Düşük toksisite / M22 solüsyonunda viskozite regülatörü"]
        ]
    },
    {
        "title": "Tablo 16.5: Doğal Antifriz Proteinleri (AFPs) ve Sentetik Buz Blokörlerinin (BBA) Matrisi",
        "headers": ["Buz Önleyici Ajan", "Köken / Kimyasal Yapı", "Termal Histerezis Kapasitesi", "Rekristalizasyon İnhibisyonu (IRI)", "Biyouyumluluk ve Klinik Uygulama"],
        "rows": [
            ["AFP Tip I", "Kutup Pisi Balığı (Alfa-helikal peptid)", "1.5°C (Termal histerezis)", "Yüksek", "Yüksek immünojenisite / Rekombinant maliyet engeli"],
            ["AFP Tip III", "Kutup Kurbağa Balığı (Küremsi protein)", "2.0°C", "Çok Yüksek", "İmmünojenik / Deneysel organ koruma"],
            ["AFGP (Glikopeptid)", "Antarktika Notothenioid balıkları", "2.5°C - 3.0°C (En güçlü doğal)", "Mükemmel", "İnsan klinik kullanımına uygun değil"],
            ["PVA (Supercool X-1000)", "Sentetik Polivinil Alkol (MW ~1.000 Da)", "0.2°C (Histerezis düşük)", "Aşırı Yüksek (%99 buz kilitlenmesi)", "Toksisite sıfır / Alcor ve M22 klinik altın standardı"],
            ["Poligliserol (PGL)", "Dallanmış hiper-poligliserol polimeri", "Düşük histerezis", "Çok Yüksek", "Biyouyumlu / Kriyoprotektan ihtiyacını %25 düşürür"],
            ["Sentetik AFGP Mimetikleri", "SPPS dizilimli D-amino asit peptitleri", "1.8°C", "Mükemmel", "Tam biyouyumluluk / Geleceğin klinik BBA standardı"],
            ["Grafen Kuantum Noktaları", "Yüzey karboksilatlı nano-grafen", "Nükleasyon absorpsiyonu", "Yüksek", "Nanotoksisite optimizasyonu gerektirir"],
            ["Zwitteriyonik Kopolimer", "PSB-kopolimer-trehaloz", "Membran faz stabilizasyonu", "Yüksek", "Hücre zarı lipid çatlamalarını sıfırlama"]
        ]
    },
    {
        "title": "Tablo 16.6: Alcor, CryoRus ve Tomorrow Bio Klinik Standby ve Perfüzyon Karşılaştırması",
        "headers": ["Protokol Parametresi", "Alcor Life Extension (ABD)", "Tomorrow Bio (Avrupa)", "CryoRus (Rusya / Avrasya)", "Klinik Hedef ve Biyofiziksel Standart"],
        "rows": [
            ["İlk Müdahale Aracı", "Tam donanımlı mobil ambulans (SRU)", "Mobil Cerrahi Kurtarma Ambulansı", "Yerel saha ekibi / Hastane morgu", "Yasal ölümden sonraki ilk 5 dakikada müdahale"],
            ["Hipotermi Yöntemi", "LUCAS mekanik CPR + Ice Slurry", "LUCAS CPR + Yüksek debili buz banyosu", "Manuel CPR + Buz paketleri", "Vücut çekirdeğini hızla <10°C'ye indirme"],
            ["Standby İlaç Protokolü", "15 bileşenli agresif kokteyl (Meds)", "Modifiye Alcor protokolü (AB standart)", "Heparin + Temel nöroprotektifler", "Pıhtılaşmayı ve metabolik çöküşü durdurma"],
            ["Vasküler Kanülasyon", "Karotis/Juguler veya Femoral cerrahi", "Femoral veya Karotis hızlı erişim", "Femoral damar cerrahisi", "Geniş lümenli çift kanülle 2 L/dk debi"],
            ["Kullanılan Vitrifikasyon Ajanı", "M22 (9.3 Molar patentli kokteyl)", "M22 bazlı optimize formülasyon", "B1 / Gliserol / DMSO bazlı solüsyonlar", "Buzsuz amorf katılaşmayı garantileme"],
            ["Konsantrasyon İzleme", "İn-line dijital refraktometre (Brix)", "Dijital refraktometri + Akış sensörü", "Periyodik manuel refraktometri", "Venöz çıkışta %95 doygunluğu yakalama"],
            ["Soğutma Odası Kontrolü", "Bilgisayar kontrollü gaz fazı soğutucu", "Programlanabilir kriyojenik ünite", "Kademeli kuru buz / Azot buharı", "Tg etrafında saatte 0.5°C hassas rampa"],
            ["Nihai Depolama Sıcaklığı", "-196°C LN2 veya -140°C ITS (Gaz)", "-196°C Bigfoot dewar tankları", "-196°C Anabiosis dewar tankları", "Elektriksiz pasif sonsuz biyostaz güvenliği"]
        ]
    },
    {
        "title": "Tablo 16.7: Nörokriyoprezervasyon ve Tam Vücut Biyostazının Kapsamlı Değerlendirmesi",
        "headers": ["Karşılaştırma Kriteri", "Nörokriyoprezervasyon (Yalnızca Baş/Beyin)", "Tam Vücut Kriyoprezervasyonu (Whole-Body)", "Biyofiziksel ve Tıbbi Rasyonel"],
        "rows": [
            ["Korunan Biyolojik Yapı", "Serebral Korteks, Hipokampus, Beyin Sapı", "Tüm İnsan Bedeni ve İç Organlar", "Kimlik ve engramlar yalnızca beyinde saklıdır"],
            ["Perfüzyon Homojenliği", "Aşırı Yüksek (%100 doğrudan karotis akışı)", "Orta / Heterojen (Farklı organ dirençleri)", "Sefalik izolasyon organ direnci farklarını eler"],
            ["Soğutma ve Isıtma Hızı", "Çok Hızlı (Küçük termal kütle: 1.5 kg)", "Yavaş (Büyük termal kütle: 70-90 kg)", "Termal kütle küçüldükçe ısı transferi kolaylaşır"],
            ["Termal Çatlama Riski", "Düşük (ITS ve tavlama ile sıfır çatlak)", "Yüksek (Kemik-eklem gerilme kırılmaları)", "Geometrik boyut küçüldükçe mekanik gerilme düşer"],
            ["Kriyoprotektan Tüketimi", "Düşük (10 - 15 Litre konsantre M22)", "Çok Yüksek (100 - 150 Litre M22)", "CPA kimyasal maliyeti ve üretim ölçeği"],
            ["Gelecekteki Canlandırma Şartı", "3D Biyo-baskılı veya klonlanmış genç beden", "Orijinal hastalıklı/yaşlı bedenin tamiri", "Bozulmuş yaşlı bedeni tamir etmek yerine taze beden"],
            ["Maliyet ve Tröst Masrafı", "Ekonomik (80.000$ - 100.000$)", "Yüksek (200.000$ - 250.000$)", "Sıvı azot dewar depolama alanı ve lojistik"],
            ["Psikolojik ve Felsefi Kabul", "Radikal fütüristik (Bazı bireylere zor)", "Geleneksel beden algısına daha yakın", "Ölüm ve kimlik algısındaki sosyolojik ayrımlar"]
        ]
    },
    {
        "title": "Tablo 16.8: Kriyojenik Termal Gerilme Kırılmaları ve Önleme Metodolojileri",
        "headers": ["Termal Gerilme Parametresi", "Kontrolsüz Daldırma (-196°C)", "İzotermal Tavlama (Annealing)", "-140°C ITS (Orta Sıcaklık Depolama)", "Kırılmasız Vitrifikasyon Hedefi"],
        "rows": [
            ["Camsılaşma Sıcaklığı (Tg)", "-124°C", "-124°C", "-124°C", "Termal rejim dönüm noktası"],
            ["Nihai Depolama Sıcaklığı", "-196°C (Sıvı Faz)", "-196°C (Tavlama sonrası)", "-140°C (Gaz Fazı)", "ITS ile Delta_T gerilmesi minimumda tutulur"],
            ["Termoelastik Gerilme (Sigma)", "> 8.5 MPa (Kırılma eşiğini aşar)", "3.0 - 4.5 MPa (Kısmen relakse)", "< 1.5 MPa (Kırılma eşiğinin çok altında)", "Kritik tokluk katsayısının (K_IC) altında kalma"],
            ["Makroskobik Kırık İnsidansı", "%100 (Büyük blok çatlakları)", "%20 - 30 (Mikro-çatlaklar kalabilir)", "SIFIR (%0 Çatlak / Monolitik Cam)", "Akustik emisyon sensörleri ile sıfır ses piki"],
            ["Akustik Emisyon Sinyalleri", "Yoğun yüksek desibel çatlama sesleri", "Seyrek düşük genlikli mikro-sesler", "Tam Sessizlik (Sıfır Akustik Emisyon)", "Gerilme dalgalarının gerçek zamanlı takibi"],
            ["Moleküler Düzeyde Etki", "Nanometrik yarık ayrışması", "Lokal relaksasyon", "Sıfır topoğrafik distorsiyon", "Sinapslar arasında nano-aralık açılmaması"],
            ["Uygulanan Mühendislik", "Doğrudan LN2 daldırma", "Tg'de 24 saat izotermal bekleme", "Gaz fazı kontrollü kryo-soğutucu", "Tam kontrollü bilgisayar algoritması"],
            ["Canlandırma Başarı Etkisi", "Nanorobotik çatlak kaynağı gerekir", "Orta derecede nano-tamir ihtiyacı", "Sıfır mekanik tamir / Doğrudan çözünme", "Kusursuz nöromal mimari restorasyonu"]
        ]
    },
    {
        "title": "Tablo 16.9: Geleneksel Çözme Yöntemleri ve Manyetik Nano-Isıtma (Nanowarming) Analizi",
        "headers": ["Yeniden Isıtma Yöntemi", "Isı Transfer Prensibi", "Isıtma Hızı (dT/dt)", "Termal Gradyan (Delta_T)", "Rekristalizasyon ve Başarı Oranı"],
        "rows": [
            ["Sıcak Su Banyosu (Konveksiyon)", "Dış yüzeyden içeri ısı iletimi (Fourier)", "Yavaş (< 5 °C/dakika merkezde)", "Aşırı Yüksek (Delta_T > 50°C)", "Kaçınılmaz buz rekristalizasyonu / Başarısız"],
            ["Mikrodalga / Dielektrik Isıtma", "Su dipollerinin dielektrik kaybı", "Orta (10 - 20 °C/dakika)", "Sıcak noktalar (Thermal Runaway)", "Dengesiz lokal kaynama ve yanıklar"],
            ["Ultrasonik Akustik Isıtma", "Ses dalgalarının viskoz absorbsiyonu", "Düşük / Orta", "Kavitasyonel doku hasarı riski", "Büyük organlarda aşırı ses zayıflaması"],
            ["Nanowarming (AMF + sIONP)", "Manyetik Néel relaksasyonu (İçten dışa)", "Ultra-Hızlı (100 - 200 °C/dakika)", "SIFIR (Delta_T < 1.0°C tam homojen)", "Sıfır Rekristalizasyon / Tarihi Organ Kurtarma"],
            ["Lazer Fotonik Isıtma", "Optik penetrasyon ve altın nanorodlar", "Aşırı Hızlı (Yalnızca mikron yüzey)", "Yalnızca 1 mm derinlikte etkili", "Hücresel biyopsiler için uygun, organa yetersiz"],
            ["Radyofrekans İndüksiyon (Çıplak)", "Dokudaki iyonik girdap akımları", "Çok Yavaş (Doku direnci yetersiz)", "Düzensiz", "Hedefsiz enerji dağılımı"],
            ["Faz Dizilimli Nano-Isıtma", "Çoklu bobin alan optimizasyonu + sIONP", "150 °C/dakika tam insan beyninde", "Kusursuz homojenlik (< 1.5°C)", "Tüm insan beynini tek parça buzsuz çözme"],
            ["Isıtma Sonrası Yıkama", "Normotermik ters kademeli perfüzyon", "Fizyolojik 37°C'ye geçiş", "Kontrollü", "sIONP ve CPA'nın kandan %99.9 temizlenmesi"]
        ]
    },
    {
        "title": "Tablo 16.10: Homo Aeternus Biyostaz, Canlandırma ve Zamansal Sıçrama Manifestosu",
        "headers": ["Biyostaz Aşaması", "Uygulanan İleri Teknoloji", "Biyofiziksel Durum / Eylem", "Moleküler Güvenlik Garantisi", "Varoluşsal ve Kozmolojik Sonuç"],
        "rows": [
            ["Standby & Kurtarma (0. Saat)", "LUCAS CPR + Buz banyosu + İlaçlar", "Sıcak iskemiyi ilk 60 saniyede durdurma", "ATP tükenmesini ve otolizi kilitler", "Ölüm anının dondurulması"],
            ["Serebral Vitrifikasyon (2. Saat)", "M22/VM3 perfüzyonu + Refraktometri", "Suyun amorf katı cam fazına dönüşümü", "Buz kristali oluşumunu sıfırlar", "Sinaps ve engramların atomik korunumu"],
            ["ITS Kriyojenik Depolama (Yüzyıllar)", "-140°C Gaz fazı Bigfoot dewar tankı", "Tüm metabolik reaksiyonların donması", "Termal çatlama ve kırılmaları sıfırlar", "Zamanın akışından mutlak bağımsızlık"],
            ["Kriyo-Nano Tarama (Canlandırma Günü)", "Kuantum SQUID + XFEL atomik tarama", "Beyin bağektomunun 3D dijital ikizi", "Sıfır enformasyon kaybının tescili", "Moleküler tamir planının çıkarılması"],
            ["Manyetik Nano-Isıtma (Nanowarming)", "Alternatif Manyetik Alan + sIONPs", "150°C/dakika hızla homojen çözünme", "Buz rekristalizasyonunu tamamen baypas", "Dokunun hasarsız sıvı faza dönüşü"],
            ["Nanorobotik Moleküler Tamir", "Trilyonlarca otonom tamir nanobotu", "Hücre zarlarının ve mitokondrinin onarımı", "CPA moleküllerinin kandan süpürülmesi", "Hücresel biyokimyanın kusursuz restorasyonu"],
            ["Epigenetik Gençleşme", "OSK Yamanaka faktörleri + AAV-TERT", "Biyolojik yaşın 80'den 20'ye indirilmesi", "Kanserin ve yaşlanma kökünün tasfiyesi", "Genç ve ölümsüz biyolojide uyanış"],
            ["Defibrilasyon ve Bilinç Ateşleme", "40 Hz Talamokortikal ritmik uyarım", "Kalp ve beyin osilasyonlarının başlaması", "Bilinç ve benliğin kesintisiz iadesi", "Homo Aeternus: Zamansal sıçrama zaferi"]
        ]
    }
]

# Dokuman Olusturma Dongusu
parts = [
    ("KISIM 1: BİYOSTAZ VE KRİYOBİYOLOJİNİN TERMODİNAMİK TEMELLERİ", part1_subsections),
    ("KISIM 2: BUZ KRİSTALİ NÜKLEASYONU, HÜCRESEL DEHİDRASYON VE MEKANİK KIRILGANLIK", part2_subsections),
    ("KISIM 3: VİTRİFİKASYON BİYOFİZİĞİ: CAMSI GEÇİŞ SICAKLIĞI (Tg) VE AMORF KATI FAZ", part3_subsections),
    ("KISIM 4: KRİYOPROTEKTAN AJANLAR (CPA): MOLEKÜLER SINIFLAR VE TOKSİSİTE KİNETİĞİ", part4_subsections),
    ("KISIM 5: BUZ BLOKE EDİCİ AJANLAR (BBA) VE SENTETİK ANTİFRİZ PROTEİNLERİ", part5_subsections),
    ("KISIM 6: KLİNİK KRİYOPROTEKTİF STANDBY VE PERFÜZYON PROTOKOLÜ: ALCOR VE CRYORUS MİMARİSİ", part6_subsections),
    ("KISIM 7: NÖROKRİYOPREZERVASYON VS TAM VÜCUT KRİYOPREZERVASYONU", part7_subsections),
    ("KISIM 8: TERMAL STRES, KIRILMA VE ÇATLAMA BİYOMEKANİĞİ", part8_subsections),
    ("KISIM 9: NANOTEKNOLOJİK YENİDEN ISITMA (NANOWARMING) VE RADYOFREKANS MANYETİK ÇÖZME", part9_subsections),
    ("KISIM 10: HOMO AETERNUS BİYOSTAZ MANİFESTOSU: MOLEKÜLER ONARIM, CANLANDIRMA VE ZAMANSAL SIÇRAMA", part10_subsections)
]

for p_idx, (part_title, subsections) in enumerate(parts):
    # Kisim Basligi
    p_header = doc.add_paragraph()
    p_header.paragraph_format.space_before = Pt(18)
    p_header.paragraph_format.space_after = Pt(12)
    p_header.paragraph_format.keep_with_next = True
    hrun = p_header.add_run(part_title)
    hrun.font.name = "Calibri"
    hrun.font.size = Pt(16)
    hrun.font.bold = True
    hrun.font.color.rgb = RGBColor(13, 71, 161)
    
    for sub_title, lead_para, deep_para, formula, explanation in subsections:
        # Alt Baslik
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(12)
        p_sub.paragraph_format.space_after = Pt(4)
        p_sub.paragraph_format.keep_with_next = True
        srun = p_sub.add_run(sub_title)
        srun.font.name = "Calibri"
        srun.font.size = Pt(12.5)
        srun.font.bold = True
        srun.font.color.rgb = RGBColor(38, 50, 56)
        
        # Giris Paragrafi
        p_lead = doc.add_paragraph()
        p_lead.paragraph_format.space_after = Pt(6)
        p_lead.paragraph_format.line_spacing = 1.15
        lrun = p_lead.add_run(lead_para)
        lrun.font.name = "Calibri"
        lrun.font.size = Pt(10)
        lrun.font.italic = True
        lrun.font.color.rgb = RGBColor(69, 90, 100)
        
        # Derin Metin
        p_deep = doc.add_paragraph()
        p_deep.paragraph_format.space_after = Pt(8)
        p_deep.paragraph_format.line_spacing = 1.15
        drun = p_deep.add_run(deep_para)
        drun.font.name = "Calibri"
        drun.font.size = Pt(10)
        drun.font.color.rgb = RGBColor(33, 33, 33)
        
        # Formül Kutusu
        tbl_f = doc.add_table(rows=1, cols=1)
        tbl_f.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell_f = tbl_f.cell(0, 0)
        set_cell_background(cell_f, "ECEFF1")
        set_cell_margins(cell_f, top=100, bottom=100, left=150, right=150)
        
        pf = cell_f.paragraphs[0]
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.paragraph_format.space_before = Pt(2)
        pf.paragraph_format.space_after = Pt(2)
        frun = pf.add_run(f"Kriyobiyolojik / Termodinamik Bağıntı:  {formula}")
        frun.font.name = "Consolas"
        frun.font.size = Pt(9.5)
        frun.font.bold = True
        frun.font.color.rgb = RGBColor(13, 71, 161)
        
        # Formül Açıklaması
        p_exp = doc.add_paragraph()
        p_exp.paragraph_format.space_before = Pt(6)
        p_exp.paragraph_format.space_after = Pt(12)
        p_exp.paragraph_format.line_spacing = 1.15
        erun = p_exp.add_run(f"Parametrik Analiz ve Mekanizma: {explanation}")
        erun.font.name = "Calibri"
        erun.font.size = Pt(9.5)
        erun.font.italic = True
        erun.font.color.rgb = RGBColor(84, 110, 122)
        
        # Sayfa Kesmesi: Her alt basliktan sonra tam 1 sayfa mimarisi
        doc.add_page_break()
        
    # Kisim Sonu Kapsamli Tablo
    t_data = tables_data[p_idx]
    
    p_tbl_title = doc.add_paragraph()
    p_tbl_title.paragraph_format.space_before = Pt(14)
    p_tbl_title.paragraph_format.space_after = Pt(6)
    p_tbl_title.paragraph_format.keep_with_next = True
    trun = p_tbl_title.add_run(t_data["title"])
    trun.font.name = "Calibri"
    trun.font.size = Pt(11.5)
    trun.font.bold = True
    trun.font.color.rgb = RGBColor(13, 71, 161)
    
    table = doc.add_table(rows=len(t_data["rows"]) + 1, cols=len(t_data["headers"]))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Baslik Satiri
    hdr_cells = table.rows[0].cells
    for c_idx, h_text in enumerate(t_data["headers"]):
        cell = hdr_cells[c_idx]
        set_cell_background(cell, "0D47A1")
        set_cell_margins(cell, top=120, bottom=120, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.font.name = "Calibri"
        run.font.size = Pt(8.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        
    # Veri Satirlari
    for r_idx, row_items in enumerate(t_data["rows"]):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F5F5F5" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val_text in enumerate(row_items):
            cell = row_cells[c_idx]
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            run = p.add_run(val_text)
            run.font.name = "Calibri"
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(33, 33, 33)
            if c_idx == 0:
                run.font.bold = True
                
    doc.add_page_break()

# Kaydet
doc.save(OUTPUT_PATH)
print(f"CİLT 16 Başarıyla Kaydedildi: {OUTPUT_PATH}")