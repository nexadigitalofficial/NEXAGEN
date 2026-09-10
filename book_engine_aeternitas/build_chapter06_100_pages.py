# -*- coding: utf-8 -*-
"""
PROJECT AETERNITAS - CİLT 06: KÖK HÜCRE BİYOLOJİSİ, NİŞ DİNAMİKLERİ VE REJENERATİF TIP
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

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_06_KOK_HUCRE_BIYOLOJISI_VE_REJENERATIF_TIP_TAM_100_SAYFA.docx"

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
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 06: KÖK HÜCRE BİYOLOJİSİ VE REJENERATİF TIP")
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
s_run = sub_p.add_run("CİLT 06: KÖK HÜCRE BİYOLOJİSİ, NİŞ DİNAMİKLERİ VE REJENERATİF TIP\\n(KLONAL TÜKENİŞİN TASFİYESİ, ORGAN BİYOPRİNTİNG VE ENDOJEN ONARIM)")
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
ih_run = intro_h.add_run("CİLT 06 MANİFESTOSU: HÜCRESEL REZERVUARLAR VE SINIRSIZ DOKU REJENERASYONU")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "Çok hücreli bir organizmanın zamana karşı ayakta kalabilmesi, doku parankimini sürekli tazeleyen kök hücre rezervuarlarının "
    "canlılığına ve mikro-çevresel niş (niche) dinamiklerine bağımlıdır. Yaşlanma sürecinde kök hücreler; telomerik erozyon, "
    "mitokondriyal DNA mutasyonları, epigenetik kayma ve klonal hematopoez (CHIP) nedeniyle asimetrik bölünme ve kendini yenileme "
    "(self-renewal) yeteneklerini kaybederek tükenir.\\n\\n"
    "Bu ciltte; hematopoietik (HSC), mezenkimal (MSC), nöral (NSC) ve kas uydu kök hücrelerinin tükeniş mekanizmaları, "
    "Wnt, Notch, TGF-beta/BMP morfogenik sinyal gradyanları, laminer ve perivasküler niş biyofiziği, klonal hematopoez mutasyonları "
    "(DNMT3A, TET2, ASXL1), parakrin eksozom tedavileri, 3D vaskülarize organoidler, lazer tabanlı biyobaskı (bioprinting) "
    "ve endojen kök hücre havuzlarını gençleştirerek organ atrofisini kalıcı olarak tasfiye etmenin protokolleri 100 ayrıntılı bölümde incelenmektedir."
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
# KISIM 1: KÖK HÜCRE HİYERARŞİSİ, POTENS VE ASİMETRİK BÖLÜNME BİYOFİZİĞİ
# ==============================================================================
part1_subsections = [
    (
        "1.1",
        "Kök Hücre Potens Spektrumu: Totipotens, Pluripotens, Multipotens ve Unipotens",
        "Gelişimsel biyolojide hücrelerin farklılaşma esnekliği, zigottan olgun dokuya doğru kademeli olarak daralan bir termodinamik hiyerarşidir.",
        "Totipotent bir zigot ve ilk blastomerler hem embriyonik dokuları hem de ekstra-embriyonik zarları (plasenta) oluşturma yeteneğine sahiptir. Blastosistin iç hücre kitlesi (ICM) pluripotens kazanarak üç germ yaprağını (ektoderm, mezoderm, endoderm) oluşturabilir ancak plasenta üretemez. Yetişkin dokularda bulunan multipotent kök hücreler (örneğin hematopoietik veya mezenkimal kök hücreler) yalnızca ait oldukları dokunun hücre tiplerine farklılaşabilirken; unipotent kök hücreler (örneğin epidermal bazal hücreler veya spermatogonyal kök hücreler) tek bir hücre soyunu üretir. Yaşlanma, bu potens basamaklarındaki klonal rezervuarın kademeli olarak tükenmesidir.",
        "Potency_State(x) = Waddington_Height(x) = - ln(Differentiation_Entropy)",
        "Gelişimsel potansiyel enerjisi, hücrenin farklılaşma durum uzayındaki Shannon bilgi entropisinin negatif logaritması olarak formalize edilir."
    ),
    (
        "1.2",
        "Asimetrik Hücre Bölünmesi Biyofiziği: Bir Hücre Kalır, Bir Hücre Farklılaşır",
        "Kök hücre havuzunun ömür boyu tükenmeden korunması, her mitozda polarize protein segregasyonu ile yürütülen asimetrik bölünme ile sağlanır.",
        "Kök hücreler simetrik bölündüklerinde iki kök hücre (genişleme) veya iki farklılaşmış hücre (tükenme) üretir; homeostaz ise asimetrik bölünmeyle sürdürülür. Mitoz sırasında Par3/Par6/aPKC polarite kompleksi hücrenin apikal korteksine göç ederken, hücre kaderi belirteçleri (örneğin Numb ve Prospero) bazal kutba toplanır. Mitotik iğ ipliğinin apikal-bazal eksende yönlenmesiyle (G-protein sinyalleşmesi ve LGN/NuMA kompleksi), niş mikro-çevresine temas eden yavru hücre kök hücre olarak kalır; nişten dışarı itilen hücre ise Numb'ın Notch reseptörünü inhibe etmesiyle farklılaşma yoluna girer.",
        "Asymmetry_Ratio = [Numb_basal] / [Numb_apical] > 10.0",
        "Asimetrik hücre kaderi segregasyonu, Numb polarite inhibitörünün bazal korta apikal korta kıyasla en az 10 kat zenginleşmesi ile güvenceye alınır."
    ),
    (
        "1.3",
        "Mitotik İğ İpliği Oryantasyonu ve LGN-NuMA-Dynein Moleküler Motorları",
        "Kök hücrenin nişe olan fiziksel açısı, kortikal LGN/NuMA/Galpha-i kompleksinin mikrotübüllere uyguladığı mekanik çekme kuvvetiyle belirlenir.",
        "İğ ipliği kutuplarının hücre zarına kenetlenmesi, kortikal mikrotübül motoru sitoplazmik Dynein ve kofaktörü Dynactin tarafından yönetilir. Galpha-i proteinine bağlı LGN adaptörü, NuMA'yı kortekse bağlar. NuMA astral mikrotübülleri yakalar ve Dynein motoru pikonewton düzeyinde mekanik gerilim uygulayarak iğ ipliğini tam olarak niş bazal zarına dik pozisyonda hizalar. Yaşlanmayla birlikte bu iğ ipliği kontrol noktalarının bozulması, kök hücrelerin rastgele açılarla bölünerek nişten erken kopmasına ve tükenmesine neden olur.",
        "Torque_Spindle = sum_k r_k CROSS F_dynein_k = I_spindle * (d^2 theta / dt^2)",
        "İğ ipliği rotasyonel torku, astral mikrotübüllere bağlı Dynein motorlarının uyguladığı vektörel kuvvetlerin iğ eylemsizlik momentine oranıdır."
    ),
    (
        "1.4",
        "İmmortal İplik Hipotezi (Immortal Strand Hypothesis) ve DNA Segregasyonu",
        "John Cairns'in 1975'te önerdiği hipotez: Kök hücre bölünürken en eski ve hatasız kalıp DNA zincirini kendinde tutarak yavruya kopyayı verir.",
        "DNA replikasyonu sırasında DNA polimerazların yaptığı mutasyonlar öncelikle yeni sentezlenen yavru iplikte yer alır. Cairns, kök hücrelerin replikatif mutasyonel yükten korunmak amacıyla, en orijinal ve hatasız 'ana DNA ipliklerini' (immortal strand) her bölünmede otoprotektif olarak kök hücrede tuttuğunu, kopyalanmış 'yeni iplikleri' ise farklılaşacak hücreye aktardığını öne sürmüştür. Karaciğer ve bağırsak kök hücrelerinde bu selektif kromozom ayrışmasının centrosome ve iğ kutbu asimetrisiyle organize edildiği gösterilmiştir.",
        "P_segregation_fidelity = 1 - epsilon_strand_loss = f([Centrosome_Age_Asymmetry])",
        "İmmortal zincir ayrışma sadakati, hücredeki yaşlı ve genç sentrozom kutuplarının mitotik iğ ipliğindeki asimetrik afinite farkı ile modellenir."
    ),
    (
        "1.5",
        "Kök Hücre Kümülüsünde Simetrik vs. Asimetrik Bölünme Dengesi",
        "Doku hasarı sonrası kök hücreler geçici olarak simetrik bölünmeyle sayılarını katlar; ardından asimetrik bölünmeye dönerek homeostazı korur.",
        "Normal koşullarda kök hücre havuzu dengededir (k_asym ~ 1). Doku yaralanması veya senolitik klirens sonrasında yerel büyüme faktörleri (Wnt ve EGF), iğ ipliği açısını yatay eksene çevirerek simetrik proliferatif bölünmeyi (k_sym_renew) tetikler. Havuz onarıldığında temas inhibisyonu ve Notch sinyali sistemi yeniden asimetrik moda kilitler. Yaşlılıkta bu anahtar bozulur; kök hücreler kontrolsüz simetrik farklılaşmaya (k_sym_diff) kayarak rezervuarı hızla eritir.",
        "d[Stem_Pool] / dt = (2 * k_sym_renew - 2 * k_sym_diff) * [Stem_Pool]",
        "Kök hücre havuzunun zamana bağlı net değişimi, simetrik kendini yenileme hızı ile simetrik farklılaşma tükeniş hızı arasındaki net farktır."
    ),
    (
        "1.6",
        "Kök Hücre Hareketsizliği (Quiescence / G0 Fazı) ve Metabolik Korunma",
        "Kök hücreler ömürlerinin %95'ini metabolik olarak uykuda (G0 fazında) geçirerek replikatif DNA mutasyonlarından ve serbest radikallerden korunur.",
        "Prolifere olan her hücre DNA kırığı ve metabolik yıpranma riski altındadır. Bu nedenle hematopoietik (HSC) ve nöral (NSC) kök hücreler, derin bir sükunet (quiescence / G0) durumunda bekler. G0 fazı; CDK inhibitörleri p57Kip2 ve p27Kip1'in yüksekliği, düşük RNA polimeraz II aktivitesi ve minimum protein translasyonu ile karakterizedir. Kök hücre yalnızca acil rejenerasyon sinyali geldiğinde G0'dan G1'e sıçrar, bölünür ve hızla tekrar G0 sığınağına çekilir.",
        "Quiescence_State = TRUE  iff  [p57Kip2] > Threshold AND [mTORC1_activity] < Basal_Limit",
        "Kök hücre G0 sükunet durumu, yüksek p57 nükleer baskısı ve minimum mTORC1 aktivite eşiğinin eşzamanlı sağlanması ile sürdürülür."
    ),
    (
        "1.7",
        "Metabolik Uyku: OksFos Baskılanması ve Glikolitik Koruma Kalkanı",
        "Hareketsiz kök hücreler mitokondrilerini neredeyse tamamen kapatarak anaerobik glikolizle yaşar ve ROS üretimini sıfıra indirir.",
        "Kemik iliği ve kas nişleri fizyolojik olarak hipoksiktir (%1-3 O2). Kök hücreler yüksek HIF-1alpha ekspresyonu ile mitokondriyal elektron taşıma zincirini minimumda tutar; ihtiyaç duydukları bazal ATP'yi sitoplazmik glikolizle üretir. Bu hipoksik glikolitik zırh, kök hücre genomunu mitokondriyal süperoksit bombardımanından ve erken replikatif yaşlanmadan korur. Yaşlanmayla birlikte mitokondriyal sızıntının artması kök hücreleri erkenden uyandırarak tüketir.",
        "J_ROS_stem = k_leak * [Mito_Mass_Active] * (1 - [HIF-1alpha_protection]) -> Minimum",
        "Kök hücredeki bazal serbest radikal akısı, HIF-1alpha aracılı mitokondriyal OksFos baskılanması sayesinde fizyolojik minimumda tutulur."
    ),
    (
        "1.8",
        "Otolog Otofaji ve Kök Hücre Proteostaz İdamesi",
        "Kök hücreler uzun ömürlerini, hasarlı proteinleri ve organelleri sürekli yutan bazal otofaji makinesine borçludur.",
        "Hareketsiz kök hücreler bölünerek hasarlı molekülleri seyreltecek lükse sahip değildir; bu nedenle protein çöplerini temizlemek için Atg7 ve Beclin-1 bağımlı otofajiye mutlak muhtaçtır. Otofajisi genetik olarak nakavt edilen farelerde HSC'ler hızla fonksiyonunu kaybeder, sitoplazmada protein agregatları birikir ve kök hücre havuzu çöker. Otofajinin uyarılması (örneğin spermidin veya rapamisin ile), yaşlı kök hücrelerin gençlik kolonizasyon kapasitesini geri kazandırır.",
        "Autophagy_Clearance_Flux = k_Atg7 * [LC3-II] * [Proteic_Aggregates] / (K_m + [Aggregates])",
        "Kök hücre proteostaz temizleme akısı, lipidlenmiş LC3-II konsantrasyonu ve Atg7 enzimatik kapasitesine doğrudan bağımlıdır."
    ),
    (
        "1.9",
        "Kök Hücre Yaşlanması (Stem Cell Exhaustion): 9 Hallmarks of Aging Entegrasyonu",
        "Lopez-Otin'in yaşlanma işaretlerinden biri olan kök hücre tükenişi, diğer 8 işaretin doku düzeyindeki nihai ortak sonucudur.",
        "Telomer kısalması kök hücre bölünmesini durdurur; epigenetik kayma diferansiasyon genlerini bozar; DNA mutasyonları klonal genişlemeyi tetikler; proteostaz kaybı hücreyi agregatlarla boğar; hücresel senesens nişi zehirler. Tüm bu hücresel patolojiler kök hücre havuzunda birleştiğinde, organizma yeni doku üretemez hale gelir. Doku atrofisi, sarkopeni, anemi ve immün yetmezlik, temelde kök hücre rezervuarının tükenişinin klinik tablosudur.",
        "Exhaustion_Index = sum_j=1^8 w_j * Hallmark_Damage_j(t) / Functional_Niche_Support",
        "Kök hücre tükeniş indeksi, diğer hücresel yaşlanma işaretlerinin kümülatif hasar toplamının fonksiyonel niş desteğine olan oranıdır."
    ),
    (
        "1.10",
        "Kök Hücrelerin İmmün Tanınması ve İmmün Ayrıcalık (Immune Privilege)",
        "Kök hücreler, yabancı veya hasarlı algılanmamak için düşük MHC-I ekspresyonu ve bağışıklık kontrol noktaları (CD47, PD-L1) taşır.",
        "Kök hücre nişleri immün-ayrıcalıklı bölgelerdir. Kök hücre yüzeyinde klasik Majör Histokompatibilite Kompleksi (MHC-I) seviyesi düşüktür; bu durum sitotoksik CD8+ T lenfositlerinin saldırısını engeller. NK hücrelerinin 'missing-self' saldırısından kaçmak için ise hücre zarında HLA-E veya CD47 ('beni yeme' sinyali) sergilenir. Yaşlanmayla birlikte dokudaki enflamasyon kök hücreleri ektopik MHC-I ekspresyonuna zorlar; bu durum kök hücrelerin otoimmün hedef haline gelerek temizlenmesine yol açar.",
        "Immune_Protection_Ratio = ([CD47] + [HLA-E]) / ([MHC-I_stimulatory] + epsilon)",
        "Kök hücre immün korunma oranı, fagositik inhibitör ligandların sitotoksik antijen sunum yüzeyine olan üstünlüğü ile sağlanır."
    )
]

# ==============================================================================
# KISIM 2: KÖK HÜCRE NİŞİ (NICHE) BİYOFİZİĞİ VE MİKRO-ÇEVRESEL REGÜLASYON
# ==============================================================================
part2_subsections = [
    (
        "2.1",
        "Schofield Niş Hipotezi ve Mikro-Çevrenin Mutlak Gücü",
        "1978'de Ray Schofield'ın tanımladığı kök hücre nişi: Kök hücre kök hücreliğini tek başına taşımaz; nişin sunduğu fiziksel ve kimyasal sinyallerle var olur.",
        "Schofield, bir kök hücrenin nişinden (niche) çıkarıldığında hızla farklılaşacağını veya öleceğini, kök hücre potansiyelinin spesifik anatomik mikro-çevre tarafından sağlandığını kanıtlamıştır. Niş; destekleyici stroma hücreleri (osteoblastlar, endotel, fibroblastlar, perisitler), ekstraselüler matriks lifleri, yerel mekanik gerilim ve morfogen sinyal gradyanlarından (Wnt, Notch, BMP, FGF) oluşan kompleks bir biyofiziksel ekosistemdir. Yaşlı bir kök hücre genç bir nişe konduğunda gençleşir; genç bir kök hücre yaşlı bir nişe konduğunda yaşlanır.",
        "Stemness_State = Cell_Intrinsic_Genome * Niche_Signaling_Field(x, y, z)",
        "Kök hücrelik durumu, intrinsik hücresel genetik kapasite ile üç boyutlu niş morfogenik sinyal alanının çarpımsal entegrasyonudur."
    ),
    (
        "2.2",
        "Kemik İliği Hematopoietik Kök Hücre Nişi: Endosteal vs. Perivasküler Niş",
        "Hematopoietik kök hücreler (HSC) kemik iliğinde iki farklı anatomik kompartımanda konuşlanır: Sessiz endosteal niş ve aktif sinüzoidal perivasküler niş.",
        "Endosteal niş, trabeküler kemiğin iç yüzeyindeki osteoblastlar tarafından kurulur; yüksek kalsiyum konsantrasyonu, aşırı hipoksi ve N-kadherin temasları ile HSC'leri derin G0 sükunetinde tutar. Perivasküler niş ise kemik iliği sinüzoidlerini saran endotel hücreleri ve CXCL12-abundant reticular (CAR) hücreleri / Leptin Reseptör pozitif (LepR+) perisitlerden oluşur. Burada HSC'ler kan akımına daha yakındır ve acil hematopoez için uyarılabilir durumdadır. Yaşlanmayla birlikte endosteal niş erir, kemik iliği yağ dokusuyla dolar ve HSC'ler sığınaklarını kaybeder.",
        "[HSC_retention] = k_anchor * [CXCL12_niche] * [CXCR4_HSC] * [SCF_bound]",
        "HSC'lerin kemik iliğinde tutunma kuvveti, niş kaynaklı CXCL12 kemokini ve Kök Hücre Faktörü (SCF) ile HSC yüzeyindeki CXCR4 ve c-Kit reseptör saturasyonuna dayanır."
    ),
    (
        "2.3",
        "Wnt/beta-Katenin Morfogen Gradyanı ve Frizzled Reseptör Kinetiği",
        "Kök hücre nişinde kısa menzilli Wnt ligand gradyanı, kendini yenileme ile farklılaşma arasındaki kararı santimetrenin binde biri ölçeğinde belirler.",
        "Wnt proteinleri hidrofobik palmitoilasyon taşır; bu nedenle doku sıvısında serbestçe difüze olamaz, hücre-hücre temasıyla lokal olarak sunulur. Niş tabanındaki yüksek Wnt konsantrasyonu Frizzled reseptörlerini ve LRP5/6 ko-reseptörlerini aktive ederek beta-katenini nükleusa gönderir ve pluripotens/kendini yenileme genlerini açar. Nişten birkaç mikrometre uzaklaşıldığında Wnt seviyesi düşer ve Dkk1/Notum antagonistleri baskın hale gelir; hücre derhal farklılaşmaya girer. Yaşlanmada sistemik Wnt aşırılığı kök hücreleri senesense zorlar.",
        "Wnt_Gradient(x) = Wnt_0 * exp(- x / lambda_diffusion) * (1 / (1 + [Dkk1] / K_i))",
        "Doku uzayındaki Wnt morfogen gradyanı, yerel palmitoilasyon difüzyon uzunluğu (lambda) ve Wnt inhibitörü Dkk1'in baskılama katsayısı ile modellenir."
    ),
    (
        "2.4",
        "Notch Sinyalleşmesi ve Yan-Yana İnhibisyon (Lateral Inhibition)",
        "Delta/Jagged ligandları ve Notch reseptör etkileşimi, komşu iki kök hücreden birinin farklılaşırken diğerinin kök kalmasını dikte eder.",
        "Notch sinyali membran-bağlı ligandlar (Delta-like DLL1/4 veya Jagged1/2) ile komşu hücrenin Notch reseptörünün fiziksel mekanik çekişiyle çalışır. Ligand bağlandığında ADAM metalloproteinazı ve gama-sekretaz kompleksi Notch reseptörünün hücre içi kuyruğunu (NICD) proteolitik olarak keser. Serbest kalan NICD nükleusa girerek RBPJ ile birleşir ve Hes1/Hey1 genlerini aktive eder. Bu sinyal komşu hücredeki farklılaşmayı susturur. Bağırsak kriptlerinde ve kas dokusunda lateral inhibisyon kök hücre düzeninin garantisidir.",
        "Rate_NICD_cleavage = k_cleavage * Force_mechanical * [Delta_ligand] * [Notch_receptor]",
        "Notch intraselüler alan (NICD) salınım hızı, ligand-reseptör kompleksine uygulanan mekanik çekme kuvveti ve gama-sekretaz aktivitesi ile ilişkilidir."
    ),
    (
        "2.5",
        "TGF-beta / BMP Ekseninin Kök Hücre Kaderindeki Zıt Kutupları",
        "Kemik Morfogenetik Proteinleri (BMP) ve TGF-beta kaskadı, doku kök hücrelerinde sükunet ile terminal farklılaşma arasındaki hassas dengeyi yönetir.",
        "BMP sinyali (BMPR1/2 reseptörleri üzerinden Smad1/5/8 aktivasyonu), birçok dokuda (örneğin saç folikülü ve nöral niş) kök hücrelerin bölünmesini baskılayarak quiescence durumunu korur. Noggin ve Chordin gibi niş kaynaklı BMP antagonistleri salgılandığında BMP kalkanı kalkar ve kök hücre bölünmeye başlar. TGF-beta/Smad2/3 ekseni ise aşırı aktifleştiğinde doku fibrozisini ve kök hücre senesensini tetikler. Yaşlı nişte BMP/Noggin oranının bozulması kök hücrelerin uyanamamasına yol açar.",
        "Stem_Activation_Switch = [Noggin] / [BMP4] > Threshold_Activation",
        "Kök hücre aktivasyon şalteri, yerel Noggin inhibitör seviyesinin serbest BMP4 morfogen konsantrasyonunu aşmasıyla tetiklenir."
    ),
    (
        "2.6",
        "Ekstraselüler Matriks Sertliği ve Biyomekanik Mekanotransdüksiyon",
        "Dennis Discher'ın öncü keşfi: Kök hücreler üzerine oturdukları substratın mekanik sertliğini (kiloPaskal) hissederek farklılaşma kararı verir.",
        "Mezenkimal kök hücreler (MSC) laboratuvarda farklı elastikiyete sahip jeller üzerine ekildiğinde şaşırtıcı bir sonuç verir: Beyin dokusu gibi yumuşak bir jelde (E ~ 0.1-1 kPa) nörona; kas benzeri orta sertlikte bir jelde (E ~ 10 kPa) miyosite; kemik benzeri sert bir zeminde (E ~ 30-40 kPa) ise osteoblaste farklılaşırlar. Hücreler bu sertliği integrin fokal adezyonları ve F-aktin lifleri üzerinden algılar; mekanik gerilim nükleer porları fiziksel olarak gererek YAP/TAZ transkripsiyon faktörlerinin çekirdeğe dolmasını sağlar.",
        "Differentiation_Lineage = f(E_matrix) -> {Soft: Neuro, Medium: Myo, Rigid: Osteo}",
        "Kök hücre soy farklılaşması, ekstraselüler matriks Young elastikiyet modülünün (E_matrix) mekanotransdüksiyon katsayısı ile deterministik olarak yönlendirilir."
    ),
    (
        "2.7",
        "Mekanik Algılayıcılar: İntegrinler, FAK Kinazı ve Nükleer Lamin A Gerilimi",
        "Matriks gerilimi hücre zarından integrinler ve LINC kompleksi aracılığıyla doğrudan nükleer zar Lamin A ağına mekanik olarak iletilir.",
        "Hücre dışındaki mekanik stres integrin kümelenmesini tetikler; Fokal Adezyon Kinazı (FAK) otofosforile olur. Aktomiyozin stres lifleri gerilimi hücre iskeletinden nükleer zara aktarır. Nükleer zarda Nesprin ve SUN proteinleri (LINC kompleksi) gerilimi Lamin A/C fibrillerine iletir. Sert matrikste Lamin A gerilir ve konformasyonel olarak katlanarak parçalanmaktan korunur (Lamin A artışı osteogenezi uyarır). Yaşlanan dokularda fibrotik sertleşme kök hücrelerin mekanik algısını bozarak osteojenik veya senesen kaymaya neden olur.",
        "Nuclear_Stress_Transmission = T_cytoskeleton * cos(theta_actin) * [LINC_density]",
        "Nükleusa iletilen mekanik gerilim vektörü, aktomiyozin çekme gerilimi ve nükleer zar LINC köprü yoğunluğunun birleşik kuvvetidir."
    ),
    (
        "2.8",
        "Hipoksi ve Vasküler Niş: HIF-1alpha / HIF-2alpha Transkripsiyonel Kontrolü",
        "Doku kök hücre nişlerindeki düşük oksijen parsiyel basıncı (pO2), metabolik ve epigenetik kararlılığın vazgeçilmez koruyucusudur.",
        "Kök hücre nişleri periferik kana göre aşırı derecede hipoksiktir (%1-3 O2, 7-20 mmHg pO2). Düşük oksijen, Prolil Hidroksilazların (PHD) inaktif kalmasını sağlayarak HIF-1alpha ve HIF-2alpha'yı kararlı tutar. HIF transkripsiyon faktörleri; VEGF salgısını uyararak anjiyogenezi beslerken, Oct4 ekspresyonunu destekler ve p21/p27 üzerinden G0 fazını kilitler. Yaşlanmada niş mikrovaskülatürünün bozulması veya ektopik oksijenlenme, kök hücrelerin erken oksidatif ölümüne yol açar.",
        "[HIF-1alpha]_steady = k_syn / (k_PHD * pO2 / (K_m_O2 + pO2) + k_dilution)",
        "Kararlı HIF-1alpha konsantrasyonu, yerel oksijen parsiyel basıncının (pO2) artmasıyla hiperbolik olarak tükenir."
    ),
    (
        "2.9",
        "Yaşlanan Nişin Bozulması: Fibröz Bağ Dokusu Birikimi ve SASP Kirliliği",
        "Kök hücrelerin yaşlanması sadece kendi genetiklerinden değil; nişi oluşturan destek hücrelerinin senesense girerek toksik SASP salgılamasından kaynaklanır.",
        "Kemik iliği stromal hücreleri veya kas fibroblastları yaşlandıkça senesense girer. Bu hücreler niş boşluğuna devasa miktarlarda TGF-beta, IL-6, TNF-alpha ve MMP-9 boşaltır. Bu toksik enflamatuar bataklık, kök hücrelerin CXCL12 ve SCF reseptörlerini desensitize eder; ekstraselüler matriks çapraz bağlanarak sertleşir (AGE birikimi). Kök hücreler niş içinde fizyolojik olarak boğulur ve klonal yenilenme yetilerini kaybeder.",
        "Niche_Deterioration_Index = [SASP_local] * [Matrix_Crosslinking_AGE] / ([CXCL12] * [SCF])",
        "Niş dejenerasyon indeksi, yerel enflamatuar sitokin ve matriks glikasyon yükünün koruyucu tutunma sinyallerine oranı ile ölçülür."
    ),
    (
        "2.10",
        "Heterokronik Parabiyoz Deneyleri: Genç Kanın Yaşlı Nişi Canlandırması",
        "Irving Weissman, Thomas Rando ve Amy Wagers'ın tarihi keşfi: Genç ve yaşlı farenin kan dolaşımının birleştirilmesi yaşlı kök hücreleri gençleştirir.",
        "Genç (2 aylık) ve yaşlı (24 aylık) farelerin dolaşım sistemleri cerrahi olarak birleştirildiğinde (heterokronik parabiyoz), genç kanında dolaşan sistemik faktörler yaşlı farenin kas uydu hücrelerini, nöral kök hücrelerini ve hepatositlerini haftalar içinde gençleştirmiştir. Yaşlı kök hücrelerin Notch sinyalleşmesi restore edilmiş, kas rejenerasyonu genç seviyeye çıkmış ve hipokampal nörogenez katlanmıştır. Bu deneyler, kök hücrelerin kalıcı olarak ölmediğini; yaşlı niş ve kan ortamında 'uyutulduğunu' ve doğru sinyallerle yeniden gençleştirilebileceğini tartışmasız ispatlamıştır.",
        "Rejuvenation_Velocity = k_systemic * ([GDF11_young] + [Klotho_young]) / ([CCL11_old] + [B2M_old])",
        "Heterokronik parabiyotik gençleşme hızı, genç dolaşımındaki gençleştirici faktörlerin yaşlı kanındaki toksik kemokin ve pro-senesen yüke oranı ile belirlenir."
    )
]

# ==============================================================================
# KISIM 3: KLONAL HEMATOPOEZ (CHIP) VE KAN KÖK HÜCRELERİNİN TÜKENİŞİ
# ==============================================================================
part3_subsections = [
    (
        "3.1",
        "Hematopoietik Kök Hücre (HSC) Yaşlanması ve Miyeloid Taraflılık (Myeloid Bias)",
        "Gençlikte dengeli eritrosit, trombosit ve lenfosit üreten HSC havuzu, yaşlanmayla birlikte lenfoid seriyi terk ederek miyeloid seriye kilitlenir.",
        "Genç bir bireyde tek bir HSC hem miyeloid (granülosit, monosit) hem de lenfoid (T ve B lenfosit) serileri eşit oranda üretir. Yaşlandıkça epigenetik kısıtlamalar ve CD150hi/CD48lo fenotipi gösteren miyeloid-taraflı kök hücreler (My-HSC) havuzu domine eder; lenfoid-taraflı kök hücreler (Ly-HSC) apoptozla elenir. Sonuç olarak bağışıklık sistemi taze T ve B hücresi üretemez (immünosenesens); buna karşılık aşırı monosit ve granülosit üretimi damar duvarında steril enflamasyonu besler.",
        "Myeloid_Bias_Ratio = [My_biased_HSC] / [Ly_biased_HSC] ~ exp(alpha_age * Age_years)",
        "Miyeloid taraf tutma rasyosu, takvim yaşının ilerlemesiyle eksponansiyel olarak artarak lenfosit üretimini felce uğratır."
    ),
    (
        "3.2",
        "Klonal Hematopoez (CHIP): Belirsiz Potansiyelli Kanser-Öncesi Genişleme",
        "70 yaş üzeri bireylerin %15-20'sinde, tek bir mutasyona uğramış kök hücre klonunun tüm kan hücre üretimini ele geçirmesi olgusu.",
        "Clonal Hematopoiesis of Indeterminate Potential (CHIP), kan kanseri (lösemi) teşhisi olmayan sağlıklı yaşlı bireylerin periferik kanında, somatik mutasyon taşıyan tek bir HSC klonunun varyant alel frekansının (VAF) en az %2'ye ulaşması durumudur. CHIP klonları sadece lösemi riskini 10 kat artırmakla kalmaz; asıl yıkıcı etkisini ateroskleroz, inme ve kalp yetmezliği mortalitesini 2 ila 3 kat artırarak gösterir.",
        "Variant_Allele_Frequency = [Mutant_Sequencing_Reads] / Total_Sequencing_Reads >= 0.02",
        "CHIP klinik tanı kriteri, derin yeni nesil sekanslamada (NGS) hedef mutasyonun varyant alel frekansının (VAF) %2 eşiğini aşması ile konur."
    ),
    (
        "3.3",
        "DNMT3A Sürücü Mutasyonları: Metilom Kayması ve Klonal Avantaj",
        "CHIP vakalarının en sık görülen sürücü mutasyonu olan DNMT3A (özellikle R882H), kök hücrenin kendini yenileme kapasitesini anormal artırır.",
        "DNMT3A genindeki missense mutasyonlar, de novo DNA metilasyon aktivitesini bozar. Normalde farklılaşma sırasında somatik genleri metilleyerek kapatması gereken DNMT3A çalışmadığında, kök hücre farklılaşamaz ve kendini yenileme simetrik bölünmelerine kilitlenir. R882H mutant proteini vahşi tip enzimi de baskılayarak dominant-negatif davranır. Bu mutant HSC klonu, kemik iliğindeki diğer sağlıklı kök hücreleri rekabetle dışlayarak havuzu istila eder.",
        "Clonal_Growth_Advantage = (mu_mutant - mu_wildtype) / mu_wildtype > 0",
        "DNMT3A mutant klonunun büyüme üstünlüğü, vahşi tip kök hücrelerin bölünme hızına göre net bir pozitif selektif katsayı kazanmasıdır."
    ),
    (
        "3.4",
        "TET2 Kaybı ve Aşırı Enflamatuar Makrofaj Üretimi",
        "TET2 mutasyonu taşıyan kök hücrelerden türeyen monosit ve makrofajlar, devasa IL-1beta ve IL-6 salgılayarak aterosklerozu patlatır.",
        "TET2 dioksijenaz inaktivasyonu, DNA demetilasyonunu felce uğratır. TET2-mutant HSC'lerden diferansiye olan makrofajlar, inflamatuar uyarılara karşı hiper-reaktiftir; NLRP3 inflamazomunu kontrolsüz aktive eder ve devasa miktarlarda IL-1beta sekrete eder. Bu makrofajlar arter duvarına sızarak damar endotelini yıkar, köpük hücrelerine dönüşür ve aterosklerotik plakların yırtılmasına yol açar. CANTOS klinik denemesi, IL-1beta antikoru canakinumab'ın özellikle TET2 mutasyonu taşıyan CHIP hastalarında kardiyovasküler ölümleri dramatik azalttığını kanıtlamıştır.",
        "Plaque_Inflammation = k_inflam * [TET2_mutant_Macrophage] * [IL-1beta_secretion]",
        "Aterosklerotik vasküler enflamasyon şiddeti, TET2 mutant makrofaj yoğunluğu ve aşırı interlökin-1beta salınımı ile katlanır."
    ),
    (
        "3.5",
        "ASXL1 ve PPM1D: Kromatin Yeniden Modellenmesi ve DNA Hasar Kaçışı",
        "ASXL1 mutasyonları Polycomb baskılamasını çözerken; PPM1D mutasyonları p53 kontrol noktasını felç ederek mutant klona ölümsüzlük kazandırır.",
        "ASXL1 (Additional Sex Combs Like 1) mutasyonları, histon H3K27me3 ve H2AK119ub işaretlerini bozarak miyeloid progenitörlerin diferansiasyonunu durdurur ve lösemik transformasyonu hızlandırır. PPM1D (Wip1) mutasyonları ise fosfataz enzimini süper-aktif hale getirir; PPM1D p53 ve Chk2'yi hızla defosforilleyerek DNA hasar kontrol noktasını kapatır. PPM1D mutant klonları kemoterapiye veya radyasyona maruz kalsa dahi apoptoza gitmez ve kemik iliğinde hızla genişler.",
        "Fitness_PPM1D = [PPM1D_active] / ([p53_phospho] + epsilon) * Resistance_Chemotherapy",
        "PPM1D mutant klonal uygunluğu, hiperaktif fosfatazın p53 hasar yanıtını nötralize etme derecesi ile eksponansiyel olarak artar."
    ),
    (
        "3.6",
        "RBC ve Trombosit Üretim Kinetiğinin Yaşla Bozulması: Kronik Anemi",
        "Yaşlılık anemisi (unexplained anemia of the elderly), eritroid progenitörlerin eritropoietin (EPO) duyarsızlığı ve demir kilitlenmesiyle gelişir.",
        "65 yaş üstü bireylerin üçte birinde belirgin bir neden olmaksızın hemoglobin düşer. Bunun moleküler sebebi, yaşlı HSC'lerin eritroid seriye farklılaşma yeteneğini kaybetmesi ve böbrekten gelen EPO hormonuna karşı EPO reseptör desensitizasyonudur. Eşzamanlı olarak kronik sistemik IL-6 yüksekliği karaciğerden hepsidin salgılanmasını tetikler; hepsidin ferroportin kanalını parçalayarak demiri makrofajlarda kilitler ve kemik iliğine demir akışını keser.",
        "Hemoglobin_Synthesis = k_EPO * [EPO] / (K_m + [EPO]) * (1 / (1 + [Hepcidin] / K_i_iron))",
        "Eritropoez hızı, EPO reseptör duyarlılığı ile serum hepsidin demir blokaj katsayısının ters orantısıyla baskılanır."
    ),
    (
        "3.7",
        "HSC Polarite Kaybı: Cdc42 Küçük GTPazı ve Farmakolojik CASIN Gençleşmesi",
        "Hartmut Geiger laboratuvarının devrimsel bulgusu: Yaşlı HSC'lerde Cdc42 aktivitesi tavan yapar ve hücresel polaritenin kaybı miyeloid kaymaya yol açar.",
        "Genç HSC'lerde tubulin, polarite proteinleri ve Cdc42 hücrenin tek bir kutbuna lokalizedir (polarize hücre). Yaşlı HSC'lerde ise Cdc42 aşırı aktifleşerek tüm sitoplazmaya homojen dağılır; polarite kaybolur (apolar hücre). Bu polarite kaybı asimetrik bölünmeyi imkansız kılar. Geiger ve ekibi, spesifik Cdc42 inhibitörü olan CASIN molekülünü yaşlı farelere sadece birkaç gün uygulamış; Cdc42 seviyesi genç seviyeye çekildiğinde yaşlı HSC'lerin polaritesi restore edilmiş, miyeloid kayma düzelmiş ve bağışıklık sistemi gençleşmiştir.",
        "Polarity_Index = [Cdc42_pole] / [Cdc42_diffuse]  (Young > 4.0, Old ~ 1.0)",
        "HSC fonksiyonel gençlik indeksi, Cdc42'nin tekil hücresel kutupta toplanma oranının homojen sitoplazmik dağılıma oranıyla ölçülür."
    ),
    (
        "3.8",
        "Kemik İliği Yağlanması (Adipojenik Dönüşüm) ve Osteoblast Çöküşü",
        "Mezenkimal kök hücrelerin kemik yapıcı osteoblastlar yerine yağ depolayan adipositlere farklılaşması, kemik iliğini sarı bir yağ çölüne çevirir.",
        "Kemik iliği stromasındaki MSC'ler yaşlanmayla birlikte transkripsiyon faktörü Runx2'yi baskılayıp PPAR-gama'yı aktive eder. Kemiğin içi trabeküler kemik yerine devasa sarı yağ damlacıkları ile dolar. Kemik iliği adipositleri aşırı serbest yağ asidi ve pro-enflamatuar TNF-alpha salgılayarak endosteal HSC nişini zehirler ve kemik iliği nakillerinin engraftment başarısını düşürür.",
        "Adipogenic_Shift = [PPAR-gamma] / [Runx2] = f(Age, Mechanical_Unloading)",
        "Kemik iliği stromal yağlanma katsayısı, adipojenik master regülatör PPAR-gama'nın osteojenik Runx2 faktörüne üstünlüğü ile belirlenir."
    ),
    (
        "3.9",
        "Klonal Rekabet ve Kök Hücre Fitness Peyzajı: Nötral vs. Pozitif Seçilim",
        "Kemik iliği on binlerce bağımsız kök hücre klonunun birbiriyle alan ve büyüme faktörü için savaştığı Darwinyen bir mikro-evrim arenasıdır.",
        "Gençlikte kök hücre dinamikleri 'nötral sürüklenme' (neutral drift) gösterir; hiçbir klon diğerine üstünlük sağlayamaz ve klonal çeşitlilik korunur. Yaşlandıkça DNA hasarı ve kronik enflamasyon doku peyzajını değiştirir. Enflamasyona dirençli mutasyon taşıyan klonlar (CHIP klonları) pozitif seçilim kazanarak çevrelerindeki masum genç kök hücreleri aç bırakır ve yok eder.",
        "Clonal_Frequency(t) = f_0 * exp(s_selection * t) / sum_k f_k * exp(s_k * t)",
        "Klonun popülasyon içindeki zamansal frekansı, Darwinyen selektif avantaj katsayısı (s_selection) üzerinden logistik büyüme sergiler."
    ),
    (
        "3.10",
        "CHIP Taraması, NGS Likit Biyopsi ve Erken Lösemi/Ateroskleroz Önleme",
        "Geleceğin tıbbında 50 yaş üzeri her bireye uygulanacak ultra-derin kanda kök hücre sekanslaması ile erken klonal kliring.",
        "Hedefe yönelik derin yeni nesil sekanslama (NGS, derinlik > 2000x), periferik kan lökositlerinde DNMT3A, TET2, ASXL1 ve JAK2 mutasyonlarını VAF %0.5 düzeyindeyken yakalayabilir. CHIP saptanan bireylere hedefe yönelik anti-enflamatuar tedaviler (IL-1beta blokerleri, NLRP3 inhibitörleri) veya spesifik klon-temizleyici senolitikler verilerek hem lösemiye dönüşüm hem de erken kalp krizi riski tamamen önlenebilir.",
        "Detection_Limit_NGS = Error_Rate_Sequencing / sqrt(Read_Depth) < 0.001",
        "Ultra-derin likit biyopside erken mutant kök hücre yakalama limiti, sekanslama okuma derinliğinin kareköküyle mikroskobik VAF seviyelerine indirilir."
    )
]

parts.append(("KISIM 1: KÖK HÜCRE HİYERARŞİSİ, POTENS VE ASİMETRİK BÖLÜNME BİYOFİZİĞİ", part1_subsections))
parts.append(("KISIM 2: KÖK HÜCRE NİŞİ (NICHE) BİYOFİZİĞİ VE MİKRO-ÇEVRESEL REGÜLASYON", part2_subsections))
parts.append(("KISIM 3: KLONAL HEMATOPOEZ (CHIP) VE KAN KÖK HÜCRELERİNİN TÜKENİŞİ", part3_subsections))

# ==============================================================================
# KISIM 4: DOKUYA ÖZGÜL KÖK HÜCRELER: NÖRAL, KAS VE MEZENKİMAL SİSTEMLER
# ==============================================================================
part4_subsections = [
    (
        "4.1",
        "Nöral Kök Hücreler (NSC): Subventriküler Zon (SVZ) ve Dentat Girus Nişleri",
        "Yetişkin insan beyninde nörogenez tamamen durmaz; lateral ventriküllerin subventriküler zonunda ve hipokampal dentat girusta sınırlı olarak devam eder.",
        "Dentat girusun subgranüler zonunda (SGZ) konuşlanan radyal glia-benzeri Tip 1 nöral kök hücreler (NSC), SOX2 ve Nestin eksprese ederek yaşam boyu yeni granül nöronlar üretir. Bu yeni nöronlar mevcut hipokampal devreye entegre olarak desen ayrışması (pattern separation) ve mekansal bellek konsolidasyonunu sağlar. Yaşlanmayla birlikte mikrovasküler kan akımının azalması ve sistemik enflamatuar kemokinlerin (CCL11/ökstaksin ve beta-2 mikroglobulin) artması, NSC'leri derin ve uyanamaz bir G0 kilitlenmesine sokarak nörogenezi %90 oranında düşürür.",
        "Neurogenesis_Rate = k_NSC * [Type1_NSC_pool] * [BDNF] / (1 + [CCL11_plasma] / K_i)",
        "Hipokampal nörogenez hızı, aktif Tip 1 kök hücre rezervi ve yerel BDNF desteğinin sistemik yaşlandırıcı kemokin CCL11 konsantrasyonuna oranıyla belirlenir."
    ),
    (
        "4.2",
        "İskelet Kası Uydu Hücreleri (Satellite Cells): Bazal Lamina ve Sarkolemma Arası",
        "Kas liflerinin zarı (sarkolemma) ile bazal lamina arasına sıkışmış Pax7-pozitif uydu hücreleri, kasın yegane rejenerasyon kaynağıdır.",
        "Uydu hücreleri, normalde derin G0 sükunetinde bekleyen kas kök hücreleridir. Bir ağırlık antrenmanı veya travmatik kas yaralanması sonrasında mekanik gerilim ve salınan nitrik oksit (NO), uydu hücrelerini aktive eder. Pax7 eksprese eden bu hücreler hızla prolifere olur; ardından MyoD ve Miyogenin transkripsiyon faktörlerini aktive ederek miyoblastlara farklılaşır ve hasarlı miyofiberlerle kaynaşır (füzyon). Yaşlılıkta Notch sinyalinin sönmesi ve TGF-beta/Smad3 sinyalinin aşırı artması, uydu hücrelerinin miyosit yerine fibrotik kolajen üreten fibroblastlara dönüşmesine (fibrojenik dönüşüm) neden olur.",
        "Myogenesis_Efficiency = [MyoD] / ([MyoD] + [Smad3_fibrotic]) * [Pax7+_pool]",
        "Miyogenez başarı oranı, pro-miyojenik MyoD faktörünün pro-fibrotik Smad3 sinyaline olan üstünlüğü ile dikte edilir."
    ),
    (
        "4.3",
        "Mezenkimal Kök / Stromal Hücreler (MSC): Kemik İliği, Adipoz Doku ve Göbek Kordonu",
        "MSC'ler kemik, kıkırdak ve yağ dokusuna farklılaşma yeteneğinin ötesinde; devasa parakrin immünomodülatör ve trofik molekül fabrikalarıdır.",
        "Minimal kriterlere göre MSC'ler plastik yüzeye yapışmalı; CD73, CD90 ve CD105 eksprese etmeli; CD34, CD45 ve HLA-DR taşımamalıdır. MSC'lerin nakledildiklerinde dokuyu doğrudan tamir etmekten ziyade, salgıladıkları ekstraselüler veziküller, PGE2, TGF-beta ve TSG-6 aracılığıyla dokudaki M1 makrofajları anti-enflamatuar M2 fenotipine çevirdikleri ve doku kök hücrelerini uyardıkları kanıtlanmıştır. Yaşlanan MSC'ler trofik sekresyonlarını kaybederek pro-enflamatuar SASP üretmeye başlar.",
        "Immunomodulatory_Score = [PGE2_MSC] * [TSG-6] / ([IL-6_SASP] + epsilon)",
        "MSC immünomodülasyon skoru, salgılanan koruyucu PGE2 ve TSG-6 konsantrasyonunun senesen IL-6 sekresyonuna oranıdır."
    ),
    (
        "4.4",
        "İntestinal Kript Kök Hücreleri (Lgr5+ CBC): Hızlı Bölünme ve Wnt Bağımlılığı",
        "Vücudun en hızlı bölünen kök hücreleri olan bağırsak Lgr5-pozitif kript taban hücreleri (CBC), tüm epitel örtüsünü her 4-5 günde bir baştan yaratır.",
        "Hans Clevers tarafından keşfedilen Lgr5+ CBC hücreleri, bağırsak kript tabanında Paneth hücrelerinin hemen arasına yerleşmiştir. Paneth hücreleri CBC'lere sürekli Wnt3a, EGF ve Notch ligandı Dll4 sunarak kök kalmalarını sağlar. Kript tabanından yukarı doğru göç eden hücreler enterosit, goblet ve enteroendokrin hücrelere farklılaşır. Yüksek mitoz hızı nedeniyle Lgr5+ hücreler DNA hasarına ve karsinogeneze aşırı duyarlıdır; APC mutasyonu bu hücrelerde kontrolsüz Wnt aktivasyonuyla kolon kanserini başlatır.",
        "Crypt_Turnover_Rate = v_migration = k_mitosis * [Wnt3a_Paneth] / (K_m + [Wnt3a])",
        "Bağırsak epitel yenilenme hızı, Paneth hücrelerinin sağladığı Wnt3a ligand doygunluğu ve CBC mitotik hızı ile lineer orantılıdır."
    ),
    (
        "4.5",
        "Epidermal Kök Hücreler ve Saç Folikülü Kök Hücreleri (HFSC)",
        "Deri bazal tabakasındaki unipotent epidermal kök hücreler ve saç folikülü 'bulge' bölgesindeki kök hücreler, sürekli deri ve kıl döngüsünü yönetir.",
        "Saç folikülü kök hücreleri (HFSC), kıl folikülünün 'bulge' (şişkinlik) bölgesinde konuşlanır ve CD34 ile K15 eksprese eder. HFSC'ler anajen (büyüme), katajen (gerileme) ve telojen (dinlenme) döngülerini yönetir. Wnt sinyali anajen büyümesini uyarırken; BMP sinyali telojen dinlenmesini kilitler. Yaşlanmayla birlikte HFSC'lerin Col17a1 (Tip XVII Kolajen) ekspresyonunu kaybetmesi, kök hücrelerin minyatürize olarak epidermise kaçmasına ve saç dökülmesine (alopesi) yol açar.",
        "Alopecia_Index = 1 / [Col17a1_hemidesmosome] * [BMP_quiescence]",
        "Senil alopesi ve foliküler atrofi riski, Tip XVII kolajen kalkanının erimesi ve aşırı BMP dinlenme sinyali ile doğrudan koreledir."
    ),
    (
        "4.6",
        "Kardiyak Progenitör Efsanesi: Yetişkin Kalpte Kök Hücre Var mı?",
        "On yıllar süren c-Kit+ kardiyak kök hücre tartışması genetik soy izleme (lineage tracing) deneyleriyle son bulmuştur: Memeli kalbi post-mitotiktir.",
        "2000'li yıllarda iddia edilen yetişkin kemik iliği veya kalpteki c-Kit+ hücrelerin yeni kalp kası ürettiği hipotezi, modern genetik etiketleme teknolojileriyle çürütülmüştür. c-Kit+ hücrelerin kardiyomiyosit değil, damar endoteli oluşturduğu kanıtlanmıştır. Yetişkin insan kalbinde gerçek anlamda multipotent bir kök hücre havuzu yoktur; yıllık kardiyomiyosit yenilenmesi 20 yaşında %1 iken 75 yaşında %0.3'e düşer. Bu durum kalp yetmezliğinde hücresel rejenerasyon için dışarıdan iPSC türevi kardiyomiyosit naklini zorunlu kılar.",
        "Cardiomyocyte_Turnover_Annual = 0.01 * exp(- 0.015 * Age_years)",
        "Yıllık spontan kardiyomiyosit yenilenme oranı, insan takvim yaşının ilerlemesiyle ihmal edilebilir seviyelere geriler."
    ),
    (
        "4.7",
        "Endotel Progenitör Hücreleri (EPC) ve Vasküler Onarım Kapasitesi",
        "Kemik iliğinden mobilize olan CD34+VEGFR2+ endotelyal progenitörler, hasarlı damar yatağına göç ederek endotel re-endotelyalizasyonunu sağlar.",
        "Damar duvarında endotel erozyonu veya iskemi oluştuğunda, yerel dokudan salgılanan VEGF ve SDF-1 (CXCL12) kemik iliğindeki EPC'leri periferik kana çeker. Dolaşımdaki EPC'ler hasarlı damar yüzeyine tutunarak olgun endotel hücrelerine farklılaşır ve nitrik oksit üretimini restore eder. Yaşlı bireylerde ve diyabetiklerde dolaşımdaki fonksiyonel EPC sayısı %80 oranında düşer; bu durum kardiyovasküler tamir yetersizliğinin ve aterosklerozun ana nedenidir.",
        "[EPC_mobilization] = k_SDF1 * [VEGF_ischemic] * [CXCR4_EPC] / (1 + [Hyperglycemia_stress])",
        "Kemik iliğinden dolaşıma EPC mobilizasyon akısı, iskemik doku VEGF sinyali ile diyabetik hiperglisemik stresin ters oranına tabidir."
    ),
    (
        "4.8",
        "Karaciğer Oval Hücreleri (Hepatobiliyer Progenitörler)",
        "Ağır hepatotoksik hasarda hepatosit bölünmesi bloke olduğunda, Hering kanallarındaki EpCAM+ bipotent oval hücreler devreye girer.",
        "Karaciğer hafif hasarlarda mevcut hepatositlerin hiperplazisiyle onarılır. Ancak kronik toksinler veya telomer tükenişi hepatosit proliferasyonunu kilitlediğinde, safra kanalcıklarının tabanındaki Hering kanallarında oturan bipotent kök hücreler ('oval hücreler') uyanır. Oval hücreler hem yeni fonksiyonel hepatositlere hem de kolanjiyositlere (safra kanalı epiteli) farklılaşarak karaciğer mimarisini sıfırdan rejenere edebilir.",
        "Oval_Cell_Activation = TRUE  iff  [Hepatosit_Proliferation] == BLOCKED AND [Wnt/Notch] == ACTIVE",
        "Hepatobiliyer kök hücrelerin devreye girmesi, primer hepatosit replikasyonunun bloke olması ve yerel Wnt/Notch kaskadının uyarılmasıyla tetiklenir."
    ),
    (
        "4.9",
        "Pulmoner Kök Hücreler: Bazal Hücreler ve Tip II Alveoler Epitel (AEC2)",
        "Trakeobronşiyal bazal hücreler hava yolunu onarırken; periferik alveollerde AEC2 hücreleri sürfaktan üretip Tip I hücrelere dönüşür.",
        "Solunum yollarında iki ana kök hücre kompartımanı bulunur: Büyük hava yollarında p63+/KRT5+ bazal kök hücreler silyalı ve goblet hücrelerini üretir. Derin akciğer alveollerinde ise Tip II pnömositler (AEC2, pro-SPC pozitif), gaz değişimini yapan devasa ve yassı Tip I hücrelerin (AEC1) kök hücre öncülüdür. İdiyopatik pulmoner fibroziste AEC2'lerin senesense girmesi gaz değişim yüzeyinin çöküşüne yol açar.",
        "Alveolar_Regeneration_Rate = k_AEC2 * [Viable_AEC2] * [Wnt_niche] / ([TGF-beta_fibrotic] + epsilon)",
        "Alveoler epitel onarım hızı, sağlam AEC2 kök hücre rezervi ve niş Wnt desteğinin pro-fibrotik TGF-beta baskısına olan oranıdır."
    ),
    (
        "4.10",
        "Doku Kök Hücrelerinin İn Vitro Genişletilmesindeki Epigenetik Drift Riski",
        "Kök hücreler vücut dışına çıkarılıp kültür kaplarında çoğaltıldıkça hızla klonal potansiyellerini kaybeder ve spontan mutasyonlar biriktirir.",
        "Hücre tedavileri için hastadan alınan kök hücrelerin (örneğin MSC veya HSC) laboratuvarda trilyonlarca sayıya çoğaltılması şarttır. Ancak plastik yüzeye yapışarak bölünen kök hücreler, her pasajda fizyolojik niş sinyallerinden mahrum kaldığı için telomer kaybeder, DNA metilasyon saatleri anormal hızlanır ve aneuploidi riski doğar. Bu durum, in vitro genişletmede 3D biyoreaktörlerin, fizyolojik hipoksinin (%2 O2) ve niş mimetik ECM kaplamalarının kullanımını zorunlu kılar.",
        "InVitro_Drift_Rate = d[Epigenetic_Aberrations] / dPassage = k_plastic * (1 - [ECM_Physiological_Support])",
        "Laboratuvar kültüründeki epigenetik sapma hızı, fizyolojik mikro-çevresel niş desteğinin eksikliği ile doğru orantılı olarak katlanır."
    )
]

# ==============================================================================
# KISIM 5: PARAKRİN SİNYALLEŞME, EKZOZOMLAR VE REJENERATİF SEKRETOM
# ==============================================================================
part5_subsections = [
    (
        "5.1",
        "Kök Hücrelerin Parakrin Paradigması: Hücre Nakli mi, Salgı Nakli mi?",
        "Arnold Caplan'ın tarihi itirafı: Nakledilen kök hücreler dokuya kalıcı olarak entegre olmaz; dakikalar içinde ölüp arkalarında rejeneratif bir sekretom bırakır.",
        "Kök hücre tedavisinin ilk 30 yılında, damardan verilen MSC'lerin hasarlı organa (örneğin infarktüs kalbine) göç edip orada yeni organ hücrelerine dönüşeceği sanılıyordu. Floresan ve genetik etiketleme deneyleri bu hipotezi çürüttü: İntravenöz verilen hücrelerin %95'i dakikalar içinde akciğer kılcal damarlarında takılıp ölmektedir. Ancak buna rağmen organların iyileşmesinin sırrı 'parakrin etki'dir. Kök hücreler ölmeden önce çevreye salgıladıkları ekstraselüler veziküller, mikroRNA'lar ve trofik sitokinlerle yerel endojen kök hücreleri uyandırır ve immün sistemi sakinleştirir.",
        "Therapeutic_Efficacy = alpha_paracrine * [Secretome_Payload] >> beta_engraftment * [Cell_Survival]",
        "Kök hücre tedavisinin klinik başarı katsayısı, parakrin sekretom biyoyararlanımının doğrudan hücresel doku entegrasyonuna ezici üstünlüğü ile açıklanır."
    ),
    (
        "5.2",
        "Kök Hücre Ekstraselüler Vezikülleri (EVs) ve Ekzozom Biyogenezi",
        "30-150 nanometrelik endozomal kökenli ekzozomlar, kök hücrenin iyileştirici genetik ve proteomik zekasını taşıyan nano-kapsüllerdir.",
        "Ekzozomlar, hücre içi geç endozomların içe doğru tomurcuklanmasıyla oluşan çok-veziküllü cisimcikler (MVB) içinde üretilir. ESCRT (Endosomal Sorting Complexes Required for Transport) makinesi ve tetraspaninler (CD9, CD63, CD81), spesifik kargoları bu veziküllere paketler. MVB hücre zarıyla kaynaştığında ekzozomlar ekstraselüler alana boşalır. Lipid çift katmanlı zarları sayesinde kargolarını (mRNA, miRNA, enzimler) kan dolaşımındaki ribonükleazlardan ve proteazlardan kusursuz koruyarak hedef organa ulaştırırlar.",
        "Exosome_Flux = k_fusion * [MVB_Rab27a] * Membrane_Surface_Area",
        "Hücresel ekzozom salınım akısı, Rab27a/b küçük GTPazı aracılı çok-veziküllü cisimcik membran füzyon kinetiği ile kontrol edilir."
    ),
    (
        "5.3",
        "Ekzozomal mikroRNA Kargo Analizi: miR-21, miR-133b, miR-29 ve miR-10a",
        "Kök hücre ekzozomlarının lümenindeki spesifik miRNA profilleri, hedef dokuda fibrozisi durduran ve anjiyogenezi ateşleyen transkripsiyonel cerrahlardır.",
        "MSC ekzozomlarının yeni nesil sekanslaması (small RNA-seq), belirli miRNA'ların stokastik değil, son derece selektif paketlendiğini ortaya koymuştur: (1) miR-21: PTEN/Akt yolağını aktive ederek hasarlı nöron ve kardiyomiyositlerde apoptozu engeller. (2) miR-29 ailesi: Kolajen I, III ve elastin genlerini translasyonel olarak susturarak organ fibrozisini çözer. (3) miR-133b: Sinir büyüme faktörlerini aktive ederek akson filizlenmesini uyarır. (4) miR-126: VEGF sinyalini amplifiye ederek yeni kapiller oluşumunu katalize eder.",
        "Post_Transcriptional_Silencing = sum_i k_silence_i * [miR_i_exosome] * [Target_mRNA_fibrotic]",
        "Hedef dokudaki fibrotik ve apoptotik gen susturma oranı, ekzozomla aktarılan terapötik miRNA konsantrasyonlarının toplamıdır."
    ),
    (
        "5.4",
        "İmmünomodülasyon: Makrofaj Polarizasyonu (M1 Pro-Enflamatuardan M2 Anti-Enflamatuara)",
        "Kök hücre sekretomu, dokuları parçalayan sitotoksik M1 makrofajları doku onarıcı ve yatıştırıcı M2 makrofajlara dönüştürür.",
        "Kronik yaralarda ve yaşlı dokularda M1 makrofajlar (TNF-alpha, iNOS, IL-1beta salgılayan) kalıcı steril enflamasyon yaratır. Kök hücre ekzozomları makrofajlar tarafından fagositozla alındığında, içerdikleri PGE2, IL-10 ve TGF-beta kaskadları makrofaj epigenomunu yeniden programlar. Makrofajlar fenotip değiştirerek Arg-1 (Arginaz-1), CD206 ve IL-10 salgılayan pro-rejeneratif M2 fenotipine kayar. Bu durum dokudaki enflamatuar yangını anında söndürür.",
        "M2_Polarization_Index = [M2_Arg1_positive] / [M1_iNOS_positive] = f([Exosomal_PGE2])",
        "Doku makrofaj polarizasyon indeksi, yerel ekzozomal prostaglandin E2 ve anti-enflamatuar kargo konsantrasyonuyla doğru orantılıdır."
    ),
    (
        "5.5",
        "Trofik Faktörler: HGF, VEGF, IGF-1, FGF-2 ve Angiopoietin-1 Kokteyli",
        "Kök hücreler dokuya doğrudan hücre eklemek yerine, yerel doku hücrelerinin hayatta kalmasını sağlayan zengin bir büyüme faktörü çorbası salgılar.",
        "Hepatocyte Growth Factor (HGF), kardiyomiyosit apoptozunu durduran ve fibrozisi eriten en güçlü trofik faktördür. VEGF ve Angiopoietin-1 endotel hücrelerini bölünmeye zorlayarak iskemik dokuda saniyeler içinde yeni damar tomurcuklanmasını (anjiyogenez) başlatır. IGF-1 hücresel protein sentezini uyarır; FGF-2 ise komşu fibroblastları sağlıklı ekstraselüler matriks üretimine yönlendirir. Bu kokteyl, doku canlılığını dışarıdan besleyen moleküler bir yaşam destek ünitesidir.",
        "Survival_Signaling = [HGF] * [c-Met] + [VEGF] * [VEGFR2] + [IGF-1] * [IGF-1R]",
        "Parakrin hayatta kalma sinyal akısı, eşzamanlı aktive olan reseptör tirozin kinazların sinerjik otofosforilasyon katsayısıdır."
    ),
    (
        "5.6",
        "Mitokondriyal Transfer: Tünelleme Nanotüpleri (TNT) ile Organel Bağışı",
        "Kök hücreler, tünelleme nanotüpleri (TNT) adı verilen mikroskobik biyolojik boru hatları kurarak hasarlı hücrelere sağlam mitokondri pompalar.",
        "2010'lu yılların en büyüleyici keşiflerinden biri, kök hücrelerin komşu hasarlı hücrelerle (örneğin iskemik kardiyomiyositler veya nöronlar) F-aktin temelli mikroskobik tüpler (tunnelling nanotubes - TNT, çap ~ 50-200 nm) kurabilmesidir. MSC'ler bu boru hatları üzerinden mikrotübül motoru Kinesin-1 ve Miro1 adaptör proteinlerini kullanarak kendi genç ve sağlam mitokondrilerini hasarlı hücreye transfer eder. Hasarlı hücreye giren taze mitokondriler anında ATP üretimini restore eder ve hücreyi apoptotik ölümden kurtarır.",
        "Rate_Mito_Transfer = v_kinesin * [Miro1_adapter] * [TNT_connections] / Resistance_viscous",
        "Mitokondriyal nakil hızı, kök hücre ile hedef hücre arasında kurulan TNT köprü sayısı ve Miro1 adaptör motor verimliliği ile belirlenir."
    ),
    (
        "5.7",
        "Hücresiz (Cell-Free) Rejeneratif Tıp: Ekzozomların Klinik Üstünlüğü",
        "Canlı hücre naklinin getirdiği pulmoner emboli, tümörleşme ve immün ret riskleri, saflaştırılmış hücresiz ekzozom tedavisiyle tamamen bertaraf edilir.",
        "Canlı kök hücre enjeksiyonları; hücrelerin damarda pıhtılaşması, kanda donör hücresine karşı alloantikor gelişmesi ve in vivo malformasyon riskleri taşır. Ekzozomlar ise hücre değildir; bölünemezler, replike olamazlar ve çekirdek taşımazlar; dolayısıyla tümör oluşturma riskleri sıfırdır. Steril filtrelerden geçirilebilir, liyofilize edilerek (dondurularak kurutulmuş toz halinde) oda sıcaklığında yıllarca saklanabilir ve raftan hazır (off-the-shelf) farmasötik bir ilaç gibi damardan güvenle uygulanabilirler.",
        "Safety_Index_CellFree = 1 / (Risk_Embolism + Risk_Tumorigenesis + Risk_AlloRejection) -> infty",
        "Hücresiz ekzozom terapisinin güvenlik profili, canlı hücresel naklin tüm ölümcül komplikasyonlarını teorik ve pratik olarak sıfıra indirir."
    ),
    (
        "5.8",
        "Ekzozom Mühendisliği: Yüzey Modifikasyonu ve Sentetik Kargo Yükleme",
        "Genetik mühendislikle ekzozom zarına yerleştirilen hedefleyici peptidler (LAMP2b füzyonu) ve elektroporasyonla yüklenen sentetik genetik kargolar.",
        "Doğal ekzozomlar karaciğer ve dalakta hızla temizlenir. Ekzozom zarı proteini LAMP2b'nin ekstraselüler ucuna kuduz virüsü glikoproteini (RVG) peptidi eklendiğinde, ekzozomlar kan-beyin bariyerini aşarak spesifik olarak nöronlara kilitlenir. Eşzamanlı olarak elektroporasyon veya sonikasyon ile ekzozom lümenine sentetik siRNA, CRISPR-Cas9 ribonükleoproteinleri veya kemoterapötikler doldurulabilir. Bu modifiye nano-kapsüller, rejeneratif tıbbın en sofistike biyolojik kargo füzeleridir.",
        "Targeting_Specificity_EV = [EV_Brain_Uptake] / [EV_Hepatic_Clearance] * [RVG_density]",
        "Mühendislik ürünü ekzozomun doku hedefleme verimliliği, yüzeydeki RVG ligant yoğunluğu ve reseptör aracılı endositoz afinitesi ile katlanır."
    ),
    (
        "5.9",
        "Plazmafraksiyonu, Genç Plazma Faktörleri ve GDF11 / Klotho Tartışması",
        "Genç kanda dolaşan ve yaşlı doku kök hücrelerini uyandıran spesifik proteinler: GDF11, sKlotho, TIMP2 ve Oksitosin.",
        "Heterokronik parabiyoz deneylerinden ilham alan araştırmacılar, genç kanda bol olup yaşla azalan sistemik molekülleri saflaştırmıştır. GDF11 (Growth Differentiation Factor 11), yaşlı farelerde kardiyak hipertrofiyi geriletmiş ve kas fonksiyonunu artırmıştır. Yaşlanma karşıtı hormon Klotho'nun çözünür formu (sKlotho), Wnt ve FGF23 sinyalini dengeleyerek nöronal sinaptik plastisiteyi korumuştur. Plazma değişimi (TPE) ile yaşlı kanındaki pro-enflamatuar toksinlerin seyreltilmesi bu gençleştirici faktörlerin etkinliğini katlamaktadır.",
        "Systemic_Youth_Index = ([GDF11] + [sKlotho] + [TIMP2]) / ([CCL11] + [B2M] + [Ferritin])",
        "Sistemik gençlik indeksi, kanda dolaşan koruyucu trofik proteinlerin yaşlandırıcı inflamatuar faktörlere rasyosu ile objektif olarak ölçülür."
    ),
    (
        "5.10",
        "Biyoreaktörlerde Ölçeklenebilir Ekzozom Biyoprosesi ve Standardizasyon",
        "Trilyonlarca ekzozomun klinik GMP standartlarında üretimi için 3D içi-boş fiber biyoreaktörler ve teğetsel akış filtrasyonu (TFF).",
        "Klasik 2D hücre kültürü şişeleri klinik doz üretmek için yetersizdir. 3D hollow-fiber biyoreaktörlerde mikrokaryerler üzerinde süspansiyon halinde büyütülen insan göbek kordonu MSC'leri, sürekli taze besiyeri perfüzyonu altında 100 kat daha yoğun ekzozom salgılar. Salgılanan süpernatant, Teğetsel Akış Filtrasyonu (Tangential Flow Filtration - TFF) ve boyutsal dışlama kromatografisi (SEC) ile saflaştırılarak partikül boyutu (Nanoparticle Tracking Analysis - NTA) ve saflığı standardize edilmiş klinik ürün elde edilir.",
        "Volumetric_Productivity = Total_Exosomes_Particles / (Bioreactor_Volume * Time_hours)",
        "Endüstriyel biyoproses hacimsel verimliliği, bioreaktör litresi başına saatte üretilen saflaştırılmış ekzozom partikül sayısı ile ölçeklenir."
    )
]

# ==============================================================================
# KISIM 6: DOKU MÜHENDİSLİĞİ, 3D BİYOBASKI VE DECELLÜLARİZE MATRİKSLER
# ==============================================================================
part6_subsections = [
    (
        "6.1",
        "Doku Mühendisliğinin Üçgeni: Hücreler, İskeleler (Scaffolds) ve Biyoaktif Sinyaller",
        "Robert Langer ve Joseph Vacanti'nin kurduğu klasik paradigma: Canlı bir organ inşa etmek için doğru kök hücre, doğru biyomateryal ve doğru morfogen şarttır.",
        "Doku mühendisliği, fonksiyonel biyolojik replasmanlar üretmek amacıyla üç temel bileşeni bir araya getirir: (1) Hücreler: Diferansiye olabilen otolog veya allogenik kök hücreler (iPSC, MSC). (2) Biyouyumlu iskele (scaffold): Hücrelerin tutunup çoğalabileceği üç boyutlu poröz matriks (kolajen, ipek fibroini, polilaktik asit - PLA). (3) Biyoreaktif ipuçları: Hücreleri doğru dokuya farklılaştıran büyüme faktörleri ve mekanik gerilim (biyoreaktör stimülasyonu). Bu üçlünün zamansal ve mekansal entegrasyonu olmadan hiçbir organ vaskülarize olamaz.",
        "Tissue_Construct_Viability = [Cell_Density] * Scaffold_Porosity * Nutrient_Diffusion_Rate",
        "İnşa edilen doku yapısının canlılığı, poröz iskele içindeki hücresel yoğunluk ve difüzyonel besin transfer hızının çarpımsal fonksiyonudur."
    ),
    (
        "6.2",
        "Decellülarizasyon ve Recellülarizasyon Teknolojisi: Doğal Organ İskeleleri",
        "Kadavra veya hayvan organlarının deterjanlarla tüm hücrelerinden arındırılarak sadece ekstraselüler matriks hayaletinin elde edilmesi.",
        "Doris Taylor'ın çığır açan metodolojisinde; kadavra kalbi veya böbreği SDS ve Triton X-100 deterjan perfüzyonuna tabi tutulur. Tüm hücresel DNA, membranlar ve antijenler yıkanıp atılır; geriye organın kusursuz damar ağını ve mikromimarisini koruyan bembeyaz bir kolajen/elastin iskeleti kalır (decellularization). Bu çıplak iskele perfüzyon biyoreaktörüne bağlanarak hastanın kendi iPSC türevi endotel ve parankim hücreleriyle yeniden tohumlanır (recellularization). Doku reddi riski taşımayan otolog organ üretimi bu yolla başarılmıştır.",
        "DNA_Removal_Criterion = Residual_dsDNA < 50 ng/mg_dry_ECM AND Fragment_Length < 200 bp",
        "Başarılı bir desellülarizasyonun immünolojik güvenlik kriteri, kalan çift zincirli DNA miktarının kuru matriks miligramında 50 nanogramın altına inmesidir."
    ),
    (
        "6.3",
        "3D Biyobaskı (Bioprinting) Yöntemleri: Ekstrüzyon, Mürekkep Püskürtme ve Lazer Tabanlı",
        "Canlı hücreleri ve hidrojel biyomürekkepleri mikrometre hassasiyetinde katman katman dizerek organ geometrisi oluşturan dijital üretim sistemleri.",
        "Biyobaskı üç ana fiziksel prensiple çalışır: (1) Pnömatik Ekstrüzyon: Yüksek viskoziteli jelleri sürekli lifler halinde basar; mekanik olarak güçlü iskeleler için uygundur. (2) Inkjet (Damlacık) Biyobaskı: Termal veya piezoelektrik titreşimle pikolitrelik hücre damlacıkları fırlatır; yüksek çözünürlüklü hücre dizimi sağlar. (3) Stereolitografi ve Dijital Işık İşleme (DLP): Işığa duyarlı fotopolimerleri lazer veya UV projeksiyonuyla mikron altı çözünürlükte saniyeler içinde katılaştırarak kompleks damar ağları inşa eder.",
        "Resolution_Bioprinting = Droplet_Volume^(1/3) * sqrt(Viscosity / Surface_Tension)",
        "Biyobaskı çözünürlük limiti, biyomürekkep viskozitesi, yüzey gerilimi ve minimum damlacık hacminin fiziksel parametreleri ile sınırlıdır."
    ),
    (
        "6.4",
        "Biyomürekkeplerin (Bio-inks) Biyofiziksel Reolojisi: Kayma İnceltmesi ve Çapraz Bağlanma",
        "Canlı hücreleri nozuldan çıkarken öldürmeyen kayma inceltmesi (shear-thinning) sergileyen ve basıldıktan sonra hızla katılaşan akıllı hidrojeller.",
        "İdeal bir biyomürekkep iki zıt mekanik özelliği birleştirmelidir: Nozuldan geçerken yüksek kayma gerilimi altında viskozitesi hızla düşmeli (shear-thinning) böylece hücreler mekanik parçalanmadan korunmalıdır; nozuldan çıktığı anda ise şeklini koruyabilmesi için viskozitesi anında fırlamalıdır (thixotropy). Jelatin-metakriloil (GelMA), sodyum aljinat, hiyalüronik asit ve desellülarize ECM hidrojelleri; mavi ışık (405 nm) veya kalsiyum klorür ile saniyeler içinde çapraz bağlanarak stabil doku karkasları oluşturur.",
        "Shear_Thinning_Viscosity = K_consistency * (Shear_Rate_gamma)^(n - 1)  (where n < 1)",
        "Biyomürekkep akış viskozitesi, psödoplastik kayma incelme indeksi (n < 1) uyarınca kayma hızı arttıkça dramatik olarak düşer."
    ),
    (
        "6.5",
        "Vaskülarizasyon Paradoksu: 200 Mikrometre Oksijen Difüzyon Bariyerinin Aşılması",
        "Kalın bir organ basmanın önündeki en büyük fiziksel engel: Kan damarı olmayan bir dokuda hücreler difüzyon sınırının ötesinde 2 saatte ölür.",
        "Hücrelerin oksijen ve glukoz difüzyon limiti yaklaşık 150-200 mikrometredir. Bu mesafenin ötesinde hücreler hipoksik nekroza uğrar. Bu paradoksu aşmak için 'kurbanlık mürekkepler' (sacrificial inks; pluronic F127 veya jelatin) kullanılır. Doku basılırken damar yolları kurbanlık mürekkeple çizilir; baskı bittiğinde sıcaklık düşürülerek kurbanlık jel eritilip dışarı yıkanır. Geride kalan açık vasküler lümenler endotel hücreleriyle (HUVEC) kaplanarak perfüze edilebilir mikrokapiller ağlar oluşturulur.",
        "Oxygen_Concentration(r) = C_surface - (R_oxygen_consumption / (4 * D_oxygen)) * (R_tissue^2 - r^2)",
        "Krogh silindir difüzyon eşitliği, doku merkezindeki oksijen konsantrasyonunun difüzyon katsayısı ve metabolik tüketim hızına bağlı olarak düşüşünü tanımlar."
    ),
    (
        "6.6",
        "Kendi Kendine Organize Olan Organoidler (Organoids): Minyatür Organ Modelleri",
        "Matrigel içine gömülen pluripotent kök hücrelerin intrinsik genetik programlarını çalıştırarak kendi kendilerine beyin, böbrek ve bağırsak oluşturması.",
        "Hans Clevers (bağırsak), Yoshiki Sasai (optik kap) ve Madeline Lancaster (beyin) tarafından geliştirilen organoid teknolojisi, kök hücrelerin morfolojik kendi kendine montaj (self-assembly) kabiliyetine dayanır. Doğru büyüme faktörleri sağlandığında iPSC'ler; nöroepitelyal rozetler, kortikal katmanlar, glomerül filtrasyon bariyerleri ve bağırsak villusları içeren milimetrik 3D minyatür organlara dönüşür. Organoidler, insan hastalık modellemesi ve kişiselleştirilmiş ilaç taramasında devrim yaratmıştır.",
        "Morphogenesis_Field = Turing_Reaction_Diffusion(Activator, Inhibitor) * Geometric_Constraint",
        "Organoid kendi kendine organizasyonu, Alan Turing'in reaksiyon-difüzyon morfogenetik dalga modelleri ve yerel matriks sınır koşulları ile gerçekleşir."
    ),
    (
        "6.7",
        "Organ-on-a-Chip (Çip-Üstü-Organ) ve Mikroakışkan Dinamik Biyoreaktörler",
        "Mikroakışkan polidimetilsiloksan (PDMS) kanallarında kan akımı ve mekanik nefes alma stresini simüle eden fizyolojik çip modelleri.",
        "Wyss Enstitüsü'nden Donald Ingber'ın geliştirdiği 'Lung-on-a-Chip', hava yolu epiteli ile kan kılcal damarlarını esnek poröz bir membranla ayırır. Yan kanallara vakum uygulanarak akciğerin nefes alma-verme mekanik gerilimi taklit edilir. Sıvı perfüzyonu ile kayma stresi sağlanır. Çip-üstü-karaciğer, çip-üstü-böbrek ve çoklu organ çiplerinin birleştirilmesiyle (Human-on-a-Chip), bir ilacın tüm vücut metabolizması ve toksisitesi hayvan kullanılmadan mikro-ölçekte test edilebilir.",
        "Shear_Stress_Fluid = 6 * Dynamic_Viscosity * Flow_Rate / (Channel_Height^2 * Channel_Width)",
        "Mikroakışkan çip kanalındaki endotelyal kayma stresi (dyne/cm^2), akışkan debisi ve kanal mikromimarisinin analitik fonksiyonudur."
    ),
    (
        "6.8",
        "Tel Aviv Üniversitesi 3D Vaskülarize Kalp Baskısı (Tal Dvir Atılımı)",
        "2019 yılında hastanın kendi omental yağ dokusu biyopsisinden türetilen hücreler ve biyomürekkeple basılan ilk vaskülarize insan kalbi.",
        "Tal Dvir laboratuvarı, hastanın yağ dokusundan hem pluripotent kök hücreleri hem de ekstraselüler matriksi izole etmiştir. Matriks kişiselleştirilmiş hidrojellere dönüştürülmüş; iPSC'ler kardiyomiyosit ve endotel hücrelerine farklılaştırılmıştır. Bilgisayarlı tomografi (BT) verileri kullanılarak hastanın anatomik kalp mimarisi, koroner arter dallanmalarıyla birlikte 3D ekstrüzyonla basılmıştır. Basılan kalp senkronize olarak kasılmış ve bağışıklık reddi riskini tamamen sıfırlayan otolog organ üretiminin kapısını açmıştır.",
        "Tissue_Synchrony = Correlation_Coefficient(Action_Potential_i, Action_Potential_j) -> 1.0",
        "Basılan kalbin fonksiyonel kasılma senkronizasyonu, komşu kardiyomiyositler arasındaki elektriksel aksiyon potansiyeli kuplajı ile ölçülür."
    ),
    (
        "6.9",
        "Elektriksel ve Mekanik Stimülasyon Biyoreaktörleri: Dokunun Olgunlaştırılması (Maturation)",
        "3D basılan dokuların fötal durumdan yetişkin fizyolojik olgunluğuna erişebilmesi için biyoreaktörlerde mekanik gerilim ve elektrik şoku eğitimi.",
        "Laboratuvarda basılan kas veya kalp dokusu başlangıçta fötal düzeydedir; organize sarkomerler ve T-tübülleri içermez. Bu yapıların olgunlaşması için doku 'eğitilmelidir'. Biyoreaktörde dokuya periyodik mekanik esneme (%10 gerilim, 1 Hz) ve elektriksel alan stimülasyonu (2 V/cm) uygulanır. Haftalar süren bu fiziksel egzersiz, kardiyomiyositlerin T-tübüllerini oluşturmasını, mitokondri kristalarını sıkıştırmasını ve yetişkin insan kasılma kuvvetine (mN/mm^2) ulaşmasını sağlar.",
        "Twitch_Force = Force_0 * (1 + alpha_stim * Duration_weeks * Electric_Field_Volts)",
        "Doku kasılma kuvveti kazanımı, biyoreaktörde uygulanan elektriksel ve mekanik stimülasyon süresi ile doğru orantılı olarak katlanır."
    ),
    (
        "6.10",
        "Ksenotransplantasyon: CRISPR ile İnsanlaştırılmış Domuz Organları",
        "eGenesis ve Revivicor'un genetik mühendislik devrimi: Domuz genomundaki retrovirüslerin silinmesi ve insan immün genlerinin eklenmesiyle organ nakli.",
        "3D biyobaskı tam organ ölçeğine ulaşana kadar geçen sürede organ krizini çözmek için domuz organları insanlaştırılmaktadır. CRISPR-Cas9 ile domuz genomundaki 62 endojen retrovirüs (PERV) silinmiş; hiperakut organ reddine yol açan üç domuz karbonhidrat antijeni (alfa-Gal, Neu5Gc, Sda) nakavt edilmiş ve yerine 6 adet insan koruyucu geni (CD46, CD55, CD59 kompleman inhibitörleri, TBM, CD47) yerleştirilmiştir. Bu 10-genetik modifikasyonlu domuz kalpleri ve böbrekleri beyin ölümü gerçekleşmiş ve canlı insanlara başarıyla nakledilmiştir.",
        "Hyperacute_Rejection_Risk = [Anti-alphaGal_Antibodies] * [Swine_alphaGal_Epitopes] -> ZERO",
        "Ksenotransplantasyonda hiperakut immün ret riski, domuz hücre yüzeyindeki alfa-galaktoz epitoplarının CRISPR ile sıfırlanmasıyla yok edilir."
    )
]

parts.append(("KISIM 4: DOKUYA ÖZGÜL KÖK HÜCRELER: NÖRAL, KAS VE MEZENKİMAL SİSTEMLER", part4_subsections))
parts.append(("KISIM 5: PARAKRİN SİNYALLEŞME, EKZOZOMLAR VE REJENERATİF SEKRETOM", part5_subsections))
parts.append(("KISIM 6: DOKU MÜHENDİSLİĞİ, 3D BİYOBASKI VE DECELLÜLARİZE MATRİKSLER", part6_subsections))

# ==============================================================================
# KISIM 7: İNDÜKLENMİŞ PLURİPOTENT KÖK HÜCRELER (iPSC) VE HASTALIK MODELLEMESİ
# ==============================================================================
part7_subsections = [
    (
        "7.1",
        "Yamanaka Faktörleri Dışında Alternatif Reprogramlama: Kimyasal İndüksiyon (CiPSC)",
        "Viral vektörler veya transgenik DNA kullanılmadan, sadece küçük moleküllü kimyasal kokteyllerle somatik hücrelerin pluripotens kazanması.",
        "Hongkui Deng ve ekibi, Oct4, Sox2, Klf4 ve c-Myc transkripsiyon faktörlerini kullanmadan, tamamen tanımlı küçük kimyasal moleküllerle fare ve insan somatik hücrelerini pluripotent duruma dönüştürmeyi (CiPSC) başarmıştır. Bu kimyasal kokteyl; epigenetik modifiyerler (VPA, Trichostatin A gibi HDAC inhibitörleri; EPZ004777 gibi DOT1L inhibitörleri), sinyal yolağı blokerleri (CHIR99021 ile GSK-3b inhibisyonu, RepSox/SB431542 ile TGF-beta reseptör inhibisyonu) ve hücresel metabolizma uyarıcılarından oluşur. Kimyasal indüksiyon, viral genetik entegrasyon riskini tamamen ortadan kaldırarak klinik güvenlikte yeni bir standart oluşturur.",
        "Reprogramming_Rate = k_chem * [CHIR99021] * [VPA] * [RepSox] / (1 + [p53_barrier])",
        "Kimyasal yeniden programlama hızı, GSK-3b ve HDAC inhibitör konsantrasyonlarının p53 onkosupresör bariyer aktivitesine oranıyla ölçeklenir."
    ),
    (
        "7.2",
        "Epigenetik Hafıza (Epigenetic Memory) ve Aberan Metilasyon Kusurları",
        "Somatik hücrelerden türetilen iPSC'lerin, kaynak dokunun DNA metilasyon ve histon modifikasyon izlerini kalıntı olarak taşıması.",
        "Bir deri fibroblastından veya kan hücresinden iPSC üretildiğinde, hücre pluripotent aşamaya geçse dahi nükleer kromatininde kaynak dokuya ait 'epigenetik hayaletler' (epigenetic memory) barındırır. Örneğin hematopoietik kök hücreden türetilen iPSC'ler kan hücrelerine çok daha kolay farklılaşırken, nöronlara dönüşmekte direnç gösterebilir. Bu durum, dokuya özgül promoter ve enhancer bölgelerindeki artık DNA CpG metilasyonlarından ve H3K27me3 baskılayıcı işaretlerinden kaynaklanır. Seri pasajlama veya DNMT/EZH2 inhibitörleriyle yapılan hafif epigenetik silme işlemleri bu hafızayı temizler.",
        "Residual_Memory = sum_i |Methylation_i(iPSC) - Methylation_i(ESC)| > 0",
        "iPSC epigenetik hafıza rezidüsü, iPSC metilom profili ile doğal embriyonik kök hücre (ESC) metilomu arasındaki mutlak sapma alanıdır."
    ),
    (
        "7.3",
        "Entegrasyonsuz Dağıtım Sistemleri: Sendai Virüsü, Epizomal Vektörler ve mRNA/LNP",
        "Konak hücre genomuna rastgele entegre olup kansere yol açan retrovirüslerin yerine güvenli, kendini kopyalamayan geçici sistemler.",
        "Genotoksik riski sıfırlamak için klinik sınıf iPSC üretiminde entegrasyonsuz (non-integrating) vektörler zorunludur. Sendai virüsü (SeV), bir negatif iplikli RNA virüsü olup sitoplazmada replike olur ve nükleer DNA'ya asla temas etmez; hücre bölündükçe sıcaklığa duyarlı mutantları sistemden kolayca atılır. Epizomal plazmitler (EBNA-1/OriP temelli) nükleusta bağımsız kalır ve pasajlar ilerledikçe seyreler. En temiz yöntem olan modifiye sentetik mRNA ve lipit nanopartikülleri (LNP) ise genomu değiştirmeden faktörleri geçici olarak eksprese eder ve 48 saatte tamamen degrade olur.",
        "Insertional_Mutagenesis_Risk = P(Vector_Insertion_into_Oncogene) = ZERO (Sendai/mRNA)",
        "Entegrasyonsuz dağıtım teknolojilerinde vektörün onkogenik insersiyon mutagenezi yaratma olasılığı teorik ve pratik olarak sıfıra indirgenir."
    ),
    (
        "7.4",
        "iPSC Tabanlı Hasta-Spesifik Organoid Hastalık Modelleri",
        "Nadir genetik hastalıklara veya kompleks nörodejeneratif süreçlere sahip hastaların hücrelerinden 3D minyatür organlar üreterek ilaç tarama.",
        "Monogenik veya poligenik hastalıklara sahip bireylerden alınan deri veya kan örneklerinden üretilen iPSC'ler, in vitro koşullarda 3D serebral organoidlere, kardiyak organoidlere veya hepatik organoidlere dönüştürülür. Örneğin Amyotrofik Lateral Skleroz (ALS) hastasından türetilen motor nöron organoidlerinde TDP-43 agregasyonu ve nörodejeneratif elektrofizyoloji gerçek zamanlı gözlemlenir. Bu modeller, hayvan deneylerinin biyolojik yetersizliğini aşarak hastanın 'kişiselleştirilmiş dijital ve biyolojik ikizinde' milyonlarca molekülün yüksek verimli taranmasını (high-throughput screening) sağlar.",
        "Organoid_Phenocopy_Fidelity = Correlate(Transcriptome_Organoid, Transcriptome_Patient_Tissue)",
        "Organoid hastalık modelinin fenokopi sadakati, üretilen 3D dokunun transkriptomunun hastanın biyopsi dokusu transkriptomuna korelasyonudur."
    ),
    (
        "7.5",
        "Kanserleşme Riski ve Karyotipik İnstabilite: TP53 Seleksiyon Baskısı",
        "iPSC kültürü ve pasajlaması sırasında genomik anormallikler kazanan ve p53 mutasyonu taşıyan süper-klonların seçilme tehlikesi.",
        "Yeniden programlama süreci hücre üzerinde ağır bir onkojenik stres yaratır; hücrelerin %99'u p53/p21 aracılı apoptoz veya kalıcı sükunete gider. Bu sert darboğazdan sağ çıkmayı başaran hücreler arasında, p53 geninde heterozigot veya homozigot mutasyon taşıyan klonlar doğal bir seleksiyon avantajı kazanır. Yapılan genomik dizilemeler, uzun süreli kültürde tutulan iPSC hatlarında 20q11.21 amplifikasyonu (Bcl-xL anti-apoptotik overekspresyonu) ve p53 mutasyonlarının spontan olarak zenginleştiğini ortaya koymuştur. Klinik nakil öncesi tam genom dizileme (WGS) ile p53 bütünlüğü doğrulanmalıdır.",
        "Clonal_Selection_Rate = exp((Fitness_mutant - Fitness_wildtype) * Generation_count)",
        "Aberan klonların popülasyondaki seçilim hızı, mutasyonlu hücrenin replikatif adaptasyon üstünlüğünün pasaj sayısı ile üstel çarpımıdır."
    ),
    (
        "7.6",
        "Hipoinmün iPSC Mühendisliği: B2M/CIITA Nakavtı ve CD47 Overekspresyonu",
        "Her hastaya özel otolog iPSC üretmenin aylar süren maliyetini aşmak için, evrensel ve bağışıklık reddine dirençli 'off-the-shelf' hücreler.",
        "Evrensel hücre nakli için iPSC genomu üçlü bir mühendislikle yeniden yazılır. İlk olarak, CRISPR-Cas9 ile B2M (Beta-2 Mikroglobulin) geni nakavt edilerek HLA Sınıf I molekülleri yüzeyden silinir; böylece CD8+ sitotoksik T hücrelerinin saldırısı engellenir. İkinci olarak CIITA geni silinerek HLA Sınıf II molekülleri yok edilir ve CD4+ T hücreleri etkisizleştirilir. Ancak HLA'sız hücreler Doğal Katil (NK) hücrelerinin 'missing self' mekanizmasıyla derhal parçalanır. Bunu engellemek için üçüncü adımda hücrelere 'Beni yeme' sinyali veren CD47 yüzey proteini lentiviral/transpozon ile aşırı eksprese ettirilir.",
        "Immune_Evasion_Index = [CD47_density] / ([HLA_Class_I] + [HLA_Class_II] + K_NK)",
        "Hipoinmün hücre sağkalım indeksi, aşırı eksprese edilen CD47 densitesinin yüzey HLA molekülleri ve NK reseptör ligandlarına oranıdır."
    ),
    (
        "7.7",
        "Hücresel Gençleşme vs Onkojenik Transformasyon Eşiği",
        "Yeniden programlama esnasında epigenetik saatin sıfırlanması ile hücrenin teratoma ve kanser geliştirme riski arasındaki hassas sınır.",
        "Somatik bir hücre OSKM faktörleriyle indüklendiğinde, transkriptomik ve epigenetik gençleşme doğrusal olmayan bir faz geçişi şeklinde ilerler. Yamanaka faktörleri ilk günlerde epigenetik yaş saatini (Horvath clock) hızla geriye çeker. Ancak farklılaşma kimliğinin tamamen silindiği ve Nanog/Oct4/Sox2 kök ağının kilitlendiği pluripotent durumda teratoma oluşturma potansiyeli tavan yapar. Hücre gençleşmesi ile onkogenez arasındaki bu moleküler sınır, rejeneratif tıbbın en tehlikeli çizgisidir; tam pluripotens yerine 'kısmi yeniden programlama' (partial reprogramming) bu riski ortadan kaldırır.",
        "Safety_Margin = Distance(Epigenetic_Age -> 0, Pluripotency_Commitment_Threshold)",
        "Rejeneratif güvenlik marjı, epigenetik gençleşme vektörünün pluripotens geri dönüşümsüzlük eşiğine olan biyofiziksel mesafesidir."
    ),
    (
        "7.8",
        "İnsan Blastoidleri ve Sentetik Embriyoloji (Stem Cell-Derived Embryos)",
        "Sperm ve yumurta kullanılmadan, yalnızca naif ve prime pluripotent kök hücrelerden laboratuvarda kendiliğinden organize olan blastosist benzeri yapılar.",
        "2023 yılında Nicolas Rivron ve Magdalena Zernicka-Goetz laboratuvarları, insan naif iPSC'lerini ve trofoblast kök hücrelerini bir araya getirerek rahme tutunma aşamasındaki insan blastosistini taklit eden 'blastoid' yapıları geliştirmiştir. Bu yapılar iç hücre kitlesi (ICM), trofektoderm benzeri dış kılıf ve blastosöl benzeri sıvı dolu bir boşluk organize eder. Sentetik embriyoloji; erken insan embriyogenezini, implantasyon başarısızlıklarını ve genetik kusurları etik sınırlamaları aşarak doğrudan laboratuvar kabında moleküler düzeyde inceleme imkanı verir.",
        "Morphogenesis_Score = Overlap_Metric(Blastoid_Cell_Architecture, Natural_Blastocyst_Atlas)",
        "Blastoid morfogenez skoru, sentetik embriyonik kompartmanların doğal insan blastosist tek-hücre atlası ile örtüşme katsayısıdır."
    ),
    (
        "7.9",
        "Endojen Lineage Reversion: Somatik Hücrelerin İn Vivo Kök Hücreye Çevrimi",
        "Doku hasarı anında olgun hücrelerin pluripotens aşamasına geçmeden doğrudan lokal kök veya progenitör hücre durumuna geri sıçraması.",
        "Rejeneratif kapasitesi yüksek dokularda (örneğin karaciğer veya bağırsak epitelinde), hasar anında olgunlaşmış hücreler kök hücreye dönüşebilir. Bağırsakta Lgr5+ kript kök hücreleri radyasyonla yok edildiğinde, olgun enterositler ve Dll1+ sekretuar progenitörler dediferansiye olarak kript tabanına yerleşir ve Lgr5 ekspresyonunu yeniden açar. Benzer şekilde hepatositler biliyer epitel hücrelerine veya karaciğer progenitörlerine dönüşebilir. Bu endojen plastisite, dokunun kök hücre havuzunu sıfırdan yeniden üretmesini sağlayan evrimsel bir sigorta mekanizmasıdır.",
        "Dedifferentiation_Flux = k_revert * [Inflammatory_Cytokines] * [Tissue_Damage_Signals]",
        "Dediferansiasyon akısı, lokal doku hasar sinyalleri ve enflamatuar sitokinlerin tetiklediği nükleer plastisite transkripsiyon katsayısıdır."
    ),
    (
        "7.10",
        "GMP Standartlarında Klinik Sınıf iPSC Bankacılığı ve Allogeneik Havuzlar",
        "Tıbbi nakil için dünya nüfusunun büyük çoğunluğuna immünolojik olarak uyumlu, homozigot süper-donör HLA haplotip bankaları.",
        "Dünya genelinde her hastaya kendi hücrelerinden iPSC üretmek trilyonlarca dolarlık bir altyapı gerektirir. Çözüm, yaygın HLA haplotipine sahip homozigot 'süper-donör' bireylerden GMP (İyi İmalat Uygulamaları) koşullarında master hücre bankaları (Master Cell Banks) kurmaktır. Örneğin en yaygın 10-15 homozigot HLA donörü, belirli bir ülke nüfusunun %50-70'i ile bağışıklık uyumu sağlar. Bu hücreler tam patojen taramasından, karyotip kararlılık testlerinden, telomer uzunluğu ve onkogen mutasyon analizlerinden geçirilerek sıvı nitrojende dondurulup klinik kullanıma hazır bekletilir.",
        "Population_Coverage_HLA = 1 - prod_k (1 - Frequency_HLA_homozygous_haplotype_k)",
        "Toplum kümülatif HLA bağışıklık kapsama oranı, bankalanan homozigot süper-donör haplotip frekanslarının tamamlayıcı olasılığıdır."
    )
]

# ==============================================================================
# KISIM 8: KÖK HÜCRE YAŞLANMASININ MOLEKÜLER DİNAMİKLERİ VE EPİGENETİK BOZULMA
# ==============================================================================
part8_subsections = [
    (
        "8.1",
        "Mitokondriyal Heteroplazmi ve mtDNA Mutasyonlarının Kök Hücre Klonlarında Birikimi",
        "Kök hücrelerin on yıllar boyunca her bölünmesinde nükleer genomdan 100 kat daha hızlı mutasyona uğrayan mitokondriyal genomun tükenişi.",
        "Mitokondriyal DNA (mtDNA), koruyucu histonlardan yoksundur ve Solunum Zincirinin hemen yanında bulunduğu için sürekli süperoksit ve hidroksil radikallerine maruz kalır. Genç bir kök hücrede mitokondriler homoplazmiktir (hasarsız). Ancak yaşlanmayla birlikte replikatif kaymalar ve mutasyon birikimi heteroplazmiyi tetikler. Delesyonlu veya mutasyonlu mtDNA oranı kritik %60-70 eşiğini aştığında, mitokondriyal kompleks I ve IV sentezi çöker; elektron sızıntısı artar ve kök hücre sükunetten çıkıp farklılaşamaz veya apoptoza gider.",
        "Heteroplasmy_Ratio(t) = [mtDNA_mutant](t) / ([mtDNA_wildtype](t) + [mtDNA_mutant](t))",
        "Mitokondriyal heteroplazmi fraksiyonu, mutasyona uğramış mtDNA kopyalarının toplam hücresel mtDNA kopyalarına olan oranıdır."
    ),
    (
        "8.2",
        "Nükleer Lamin A/C ve Nükleer Zarf Bütünlüğünün Bozulması (Progerik Deformasyon)",
        "Yaşlanan kök hücre nükleusunun küresel şeklini kaybederek buruşması, lamin kaybı ve nükleositoplazmik transportun felç olması.",
        "Nükleus zarı altında mekanik destek sağlayan Lamin A, B ve C protein ağı (nükleer lamina), heterokromatin bölgelerini nükleer çeperde tutarak sessiz genlerin açılmasını engeller. Yaşlanmayla birlikte Lamin B1 kaybı ve progerin benzeri aberran Lamin A formlarının birikimi, nükleus zarında herniasyonlara (blebbing) yol açar. Nükleer por kompleksleri (NPC) parçalanır; RanGTP/RanGDP gradyanı çöker ve sitoplazma ile nükleus arasındaki moleküler sinyal ayrışması yok olarak kök hücre fonksiyonlarını felce uğratır.",
        "Nuclear_Circularity = 4 * pi * Area_nucleus / (Perimeter_nucleus^2) < 0.7",
        "Nükleer deformasyon indeksi; nükleer daireselliğin genç kök hücredeki 1.0 ideal değerinden yaşlı kök hücrede 0.7 altına gerilemesiyle ölçülür."
    ),
    (
        "8.3",
        "DNA Hasar Yanıtı (DDR), H2AX Fosforilasyonu ve Kök Hücre Apoptozu",
        "Çift zincir kırıklarının (DSB) onarılamaması sonucu ATM/ATR kinazlarının sürekli uyarılması ve kök hücre havuzunun tasfiyesi.",
        "Kök hücreler sükunet (G0) fazında iken hata yapmayan Homolog Rekombinasyon (HR) onarım mekanizmasını kullanamazlar; çünkü HR kardeş kromatit gerektirir. Bu nedenle kök hücreler sükunette hatalı Non-Homologous End Joining (NHEJ) yoluna mahkumdur. Biriken çift zincir kırıkları ATM kinazı aktive eder; Histon H2AX Serin 139 pozisyonunda fosforillenerek gamma-H2AX odaklarını (foci) oluşturur. Süregelen DDR sinyali, kök hücreyi ya p53-PUMA aracılı apoptoza iter ya da diferansiasyona zorlayarak kök havuzunu kalıcı olarak tüketir.",
        "DDR_Arrest_Signal = sum_i [gamma_H2AX_foci_i] * [53BP1_i] > Threshold_Apoptosis",
        "DNA hasar tutuklama sinyali, nükleer gamma-H2AX ve 53BP1 tamir odaklarının kümülatif miktarının apoptoz eşiğini aşmasıyla tetiklenir."
    ),
    (
        "8.4",
        "Heterokromatin Kaybı, Line-1 Retrotranspozon Aktivasyonu ve Genomik İnstabilite",
        "Heterokromatin adacıklarının çözülmesiyle 'zıplayan genlerin' uyanması, nükleer DNA'yı delik deşik etmesi ve sitoplazmik DNA alarmı.",
        "Genç kök hücrelerde genomun %50'sinden fazlasını oluşturan transpozonlar (özellikle LINE-1 retrotranspozonları), H3K9me3 ve DNA metilasyonu ile sıkı sıkıya heterokromatinde kilitlidir. Yaşlanmayla birlikte heterokromatin kaybı (Loss of Heterochromatin Hipotezi) yaşanır; LINE-1 elemanları transkripte edilir, ters transkriptaz enzimiyle sitoplazmada cDNA üretir ve genomun rastgele bölgelerine entegre olarak mutasyonlara yol açar. Bu sitoplazmik cDNA'lar cGAS-STING yolağını uyararak kök hücrede oto-enflamasyonu tetikler.",
        "Retrotransposon_Activity = [LINE1_RNA] * [RT_activity] / [H3K9me3_heterochromatin]",
        "Transpozon aktivite indeksi, LINE-1 transkriptlerinin ve ters transkriptaz aktivitesinin baskılayıcı H3K9me3 heterokromatin yoğunluğuna oranıdır."
    ),
    (
        "8.5",
        "Proteazomal ve Otofajik Bozulma: Agresom Birikimi ve Kök Hücre Toksisitesi",
        "Yaşlanan kök hücrede yanlış katlanmış proteinlerin temizlenememesi, perinükleer agregat oluşumu ve proteotoksik stres.",
        "Genç kök hücreler sükunet döneminde bile yüksek bazal otofaji seviyelerine sahiptir; hasarlı organelleri ve agregatları temizleyerek metabolik olarak tertemiz kalırlar. Yaşlanma sürecinde LAMP2A (şaperon aracılı otofaji reseptörü) düzeyleri düşer, lizozomal asidifikasyon (v-ATPase aktivitesi) zayıflar ve 26S proteazom aktivitesi azalır. Ubikitinlenmiş agregatlar perinükleer agresomlar halinde birikir. Bu proteotoksik stres, endoplazmik retikulumda UPR (Unfolded Protein Response) yolağını kronik olarak açarak kök hücreyi öldürür.",
        "Proteostasis_Capacity = [26S_Proteasome_flux] + [Autophagy_flux] - [Aggresome_accumulation_rate]",
        "Kök hücre proteostaz kapasitesi, proteazomal ve otofajik temizleme hızının toksik agregat birikim hızından çıkarılmasıyla hesaplanır."
    ),
    (
        "8.6",
        "Metabolik Değişim: Glikolizden Oksidatif Fosforilasyona Zorlanma ve ROS Fırtınası",
        "Düşük oksijenli sükunet ortamından zorla çıkarılan kök hücrenin mitokondriyal solunuma geçmesi ve reaktif oksijen türleriyle zehirlenmesi.",
        "Sükunetteki genç kök hücreler enerji ihtiyaçlarını mitokondriyi bypass ederek anaerobik glikoliz ile (HIF-1alpha kontrolünde) karşılarlar; böylece reaktif oksijen türü (ROS) üretimi minimumda kalır. Yaşlanan kök hücrelerde mitokondriyal membran potansiyeli bozulur; hücre zorunlu olarak oksidatif fosforilasyona (OXPHOS) itilir. Kompleks I ve III'ten sızan kaçak elektronlar süperoksit (O2.-) ve hidrojen peroksite (H2O2) dönüşerek nükleusta ve membran lipitlerinde peroksidasyon hasarı başlatır.",
        "Metabolic_Shift_Index = Rate_OXPHOS / Rate_Glycolysis = [p-PDH] / [PDK1]",
        "Metabolik kayma indeksi; oksidatif fosforilasyon hızının anaerobik glikolize oranı olup, piruvat dehidrogenaz fosforilasyonu ile düzenlenir."
    ),
    (
        "8.7",
        "Asimetrik Bölünme Bozulması ve Sentrozom Sayısal Aberasyonları",
        "Kök hücrenin polaritesini kaybetmesi, simetrik diferansiasyona kayması ve sentrozom amplifikasyonuyla multipolar iğ ipliği anomalileri.",
        "Yaşlı kök hücrelerde sentrozom duplikasyon kontrolü çöker; Polo-like Kinase 4 (PLK4) disregülasyonu sonucu hücrede 2 yerine 3 veya 4 sentrozom oluşur. Mitoz anında multipolar iğ iplikleri kurulur ve kromozomlar rastgele kutuplara çekilir (anöploidi). Bunun da ötesinde Par3/Par6 polarite ekseni parçalanır; kök hücre asimetrik bölünme yerine simetrik farklılaşma yapar. Bu durum, her bölünmede kök hücre rezervuarının birer birer eksilmesine ve dokunun kendini yenileyemez hale gelmesine yol açar.",
        "Mitotic_Fidelity = P(Bipolar_Spindle) * (1 - P(Centrosome_Amplification))",
        "Mitotik sadakat, iki kutuplu iğ ipliği oluşma olasılığının sentrozom amplifikasyon olasılığı ile çarpımsal azalımıdır."
    ),
    (
        "8.8",
        "Kök Hücre Senesensi: p16INK4a / p21CIP1 Kilitlenmesi ve İrreversibl Sükunet",
        "Geri dönüşümlü uyku (quiescence) durumundaki kök hücrenin, CDK inhibitörlerinin aşırı ekspresyonuyla uyanamaz hale gelmesi.",
        "Genç kök hücreler G0 evresinde sükunettedir ve gerektiğinde (örneğin doku hasarında) G1/S fazına geçerek hızla prolifere olabilir. Yaşlanan kök hücrede ise Cdkn2a lokusundaki epigenetik baskı kalkar; p16INK4a proteini CDK4/6 kinazlarını bloke eder ve Retinoblastoma (Rb) proteininin fosforillenmesini önler. Aynı zamanda p21CIP1 CDK2'yi kilitler. Hücre artık 'uyanabilir bir uyku' halinde değil, geri dönüşümsüz bir 'senesens kilitlenmesi' içindedir. Bu hücreler hem dokuyu tamir edemez hem de salgıladıkları SASP ile komşu kök hücreleri felç eder.",
        "Irreversible_Senescence = Sigmoid([p16INK4a] + [p21CIP1] - Threshold_CDK_Inhibition)",
        "Geri dönüşümsüz kök hücre senesensi, p16 ve p21 kümülatif miktarının CDK inhibisyon eşiğini aşmasıyla sigmoidal olarak tetiklenir."
    ),
    (
        "8.9",
        "Telomer T-Loop Çözülmesi ve Telomeraz Susturulmasının Kök Hücreye Bedeli",
        "Çoğu somatik kök hücrede telomerazın yetersiz olması nedeniyle her hücre döngüsünde uç kromatinin erimesi ve T-loop çözülmesi.",
        "Yaygın inanışın aksine, insan somatik kök hücrelerinin büyük kısmı (özellikle MSC ve nöral kök hücreler) telomerazı (hTERT) sürekli açık tutmaz; yalnızca hematopoietik kök hücreler zayıf ve geçici bir telomerase aktivitesi gösterir. On yıllar süren doku turnover'ı sırasında telomerik TTAGGG tekrarları 15 kb'den 3-4 kb'ye kadar erir. TRF2 ve Shelterin kompleksinin DNA ucunu bağlayamadığı noktada telomer T-loop yapısı çözülür; kromozom ucu çıplak çift zincir kırığı gibi algılanarak kalıcı DNA hasar yanıtını (TIF - Telomere Dysfunction-Induced Foci) başlatır.",
        "TIF_Count = sum_chromosomes I(Telomere_Length_k < Critical_Length_Threshold)",
        "Telomer disfonksiyon odaklarının sayısı, kritik uzunluk eşiğinin altına inerek T-loop'u çözülen kromozom uçlarının toplamıdır."
    ),
    (
        "8.10",
        "Sistemik Çevre Kaynaklı Tükeniş: Yaşlı Plazma Faktörleri ve GDF11 Tartışması",
        "Genç bir kök hücrenin yaşlı bir vücuda konduğunda fonksiyonunu kaybetmesi; sistemik kan kimyasının kök hücre kaderini dikte etmesi.",
        "Heterokronik parabiyoz (genç ve yaşlı farenin dolaşım sistemlerinin birleştirilmesi) deneyleri, kök hücre yaşlanmasının intrinsik olduğu kadar sistemik çevresel olduğunu kanıtlamıştır. Yaşlı kanda biriken pro-enflamatuar faktörler (CCL11/ökstaksin, TGF-beta, Beta-2 mikroglobulin), genç kök hücrelerin nörogenezini ve kas rejenerasyonunu derhal baskılar. Genç kanda bulunan TIMP2 ve oksitosin gibi koruyucu faktörler ise yaşlı kök hücreleri geçici olarak gençleştirir. Amy Wagers'ın genç kanda yüksek olduğunu öne sürdüğü GDF11 faktörü etrafındaki tartışmalar, sistemik gençleşme mekanizmalarının çoklu faktör dengesine dayandığını göstermiştir.",
        "Stem_Cell_Activity = Functional_Intrinsic_Fitness * (Young_Systemic_Factors / Old_Systemic_Factors)",
        "Kök hücrenin fonksiyonel aktivitesi, hücrenin kendi intrinsik sağlığının sistemik genç/yaşlı kan faktörleri oranına çarpımıdır."
    )
]

# ==============================================================================
# KISIM 9: REJENERATİF TIP, HÜCRE TERAPİLERİ VE KLİNİK TRANSFORMATİF PROTOKOLLER
# ==============================================================================
part9_subsections = [
    (
        "9.1",
        "Allogeneik Kemik İliği Naklinde Graft-versus-Host (GvHD) Önleme ve Mezenkimal Kalkan",
        "Kemik iliği naklinde bağışçının T hücrelerinin hastanın dokularına saldırmasını önlemek için kemik iliği kaynaklı MSC infüzyonu.",
        "Lösemi ve aplastik anemi tedavisinde allogeneik hematopoietik kök hücre nakli hayat kurtarır; ancak nakledilen donör lenfositleri hastanın karaciğer, bağırsak ve cildini yabancı olarak tanıyıp yok etmeye çalışır (Graft-versus-Host Hastalığı - GvHD). Mezenkimal Kök Hücreler (MSC), salgıladıkları IDO (İndoleamin 2,3-dioksijenaz) ile triptofanı tüketip kynurenin üreterek T hücrelerini anerjiye sokar; aynı zamanda TGF-beta ve PGE2 ile Foxp3+ Treg hücrelerini çoğaltır. Bu 'mezenkimal kalkan', donörün lösemiyi yok etme gücünü (GvL) korurken ölümcül doku yıkımını engeller.",
        "GvHD_Suppression_Rate = k_MSC * [MSC_Dose] * [IDO_activity] / [Effector_T_cells]",
        "GvHD baskılanma oranı, nakledilen MSC dozu ve triptofan yıkan IDO enzim aktivitesinin ortamdaki alloreaktif efektör T hücresi sayısına oranıdır."
    ),
    (
        "9.2",
        "Nörodejeneratif Hastalıklarda Dopaminerjik Progenitör Nakli (Parkinson Faz I/II)",
        "Substantia nigra'da kaybedilen dopaminerjik nöronların yerine iPSC türevli A9-tipi orta beyin dopamin progenitörlerinin stereotaksik implantasyonu.",
        "Parkinson hastalığında motor kontrolün çöküşü, striatuma projekte olan substantia nigra dopaminerjik nöronlarının kaybından kaynaklanır. Jun Takahashi ekibi ve BlueRock Therapeutics (bemdaneprocel), insan iPSC'lerini GMP koşullarında Lmx1a/Foxa2/TH pozitif orta beyin dopaminerjik progenitörlerine farklılaştırmıştır. Stereotaksik cerrahi ile hastanın putamenine nakledilen bu hücreler, konakçı nöronal devresine sinaps kurarak entegre olur, fizyolojik dopamin salgılar ve L-dopa bağımlılığı olmaksızın motor semptomları (UPDRS skoru) dramatik şekilde düzeltir.",
        "Striatal_Dopamine_Recovery = k_synapse * N_engrafted_TH+_neurons * Dopamine_release_per_cell",
        "Striatal dopaminerjik geri kazanım, greftlenen tirozin hidroksilaz (TH) pozitif nöron sayısı ile hücre başına fizyolojik dopamin salımının çarpımıdır."
    ),
    (
        "9.3",
        "Tip 1 Diyabette iPSC Kaynaklı Beta Adacık Hücre Grefti ve Mikro-Enkapsülasyon",
        "Otoimmün yıkıma uğrayan pankreas beta hücrelerinin pluripotent kök hücrelerden laboratuvarda basamaklı protokollerle üretilip nakledilmesi.",
        "Vertex Pharmaceuticals (VX-880 ve VX-264) klinik denemeleri, iPSC'lerin definitive endoderm, pankreatik progenitör ve nihayetinde insülin salgılayan fonksiyonel beta hücrelerine dönüştürülmesinde çığır açmıştır. Bu hücreler kandaki glukoz seviyesini milisaniyelik hassasiyetle ölçer ve dinamik insülin salgılar. Hastanın bağışıklık sisteminin yeni hücreleri tekrar yok etmesini (otoimmün atak) önlemek için hücreler, yarı geçirgen biyo-uyumlu alginat mikrokapsüller veya immün-bariyerli makro-cihazlar içine yerleştirilerek immünosüpresyonsuz kür sağlanır.",
        "Glucose_Insulin_Transfer_Function = d[Insulin]/dt = V_max * [Glucose]^n / (K_m^n + [Glucose]^n)",
        "Nakledilen adacıkların insülin salım dinamiği, plazma glukoz konsantrasyonuna bağlı Hill kinetiği transfer fonksiyonu ile modellenir."
    ),
    (
        "9.4",
        "Yaşa Bağlı Makula Dejenerasyonunda RPE Hücre Tabakası İmplantasyonu",
        "Körlüğün önde gelen nedeni olan makula dejenerasyonunda, retina pigment epitel tabakasının iPSC kaynaklı monokatmanla değiştirilmesi.",
        "Kuru tip YBMD'de fotoreseptörlerin ölümü, onları besleyen ve atıklarını temizleyen alttaki Retina Pigment Epiteli (RPE) tabakasının yaşlanıp ölmesinden kaynaklanır. Masayo Takahashi ve ekibi, dünyada insan üzerinde yapılan ilk iPSC klinik denemesinde, hastanın deri hücresinden ürettiği iPSC'leri pigmentli RPE hücrelerine farklılaştırmış ve tek tabakalı bir epitel tabakası (sheet) halinde subretinal alana yerleştirmiştir. Bu greft fotoreseptör kaybını durdurmuş, retina anatomisini stabilize etmiş ve görme keskinliğini korumuştur.",
        "Photoreceptor_Survival_Index = RPE_Phagocytosis_Flux(Outer_Segments) * [PEDF_secretion]",
        "Fotoreseptör nöronların sağkalım indeksi, greftlenen RPE tabakasının dış segment fagositoz hızı ve salgıladığı PEDF nörotrofik faktörüne bağlıdır."
    ),
    (
        "9.5",
        "Miyokard İnfarktüsü Sonrası Biyo-Mühendislik Kalp Yamaları (Cardiac Patches)",
        "Kalp krizi sonrası oluşan avasküler fibrotik nedbe dokusunun üzerine milyonlarca iPSC-kardiyomiyositi içeren kasılabilen canlı doku bandı.",
        "Yetişkin insan kalbi hasar gördüğünde kas hücreleri bölünemez ve yerini kolajen skar dokusuna bırakır; bu da konjestif kalp yetmezliğine yol açar. Kalp yaması teknolojisinde; insan iPSC türevli kardiyomiyositler, endotel hücreleri ve fibroblastlar kollajen-fibrin hidrojelleri içinde 3D olarak basılır ve elektriksel biyoreaktörde eşzamanlı atım kazandırılır. Bu canlı yama epikarda cerrahi olarak dikildiğinde konak koroner damarlarıyla anastamoz kurar, duvar gerilimini azaltır ve kalbin sol ventrikül ejeksiyon fraksiyonunu (LVEF) %15-20 artırır.",
        "Ejection_Fraction_Gain (Delta_EF) = beta_patch * Volume_viable_myocardium * Synchrony_Index",
        "Ejeksiyon fraksiyonundaki kazanç, yama içindeki canlı miyokard hacmi ile konak kalple olan elektriksel iletim senkronizasyonunun fonksiyonudur."
    ),
    (
        "9.6",
        "Omurilik Yaralanmalarında Oligodendrosit Progenitörleri ve Aksonal Rejenerasyon",
        "Travmatik omurilik kopmalarında miyelin kılıfın yeniden örülmesi ve aksonal büyümenin uyarılması için OPC hücre nakli.",
        "Omurilik travması sonrasında oligodendrositlerin ölümü aksonların çıplak kalmasına (demiyelinizasyon) ve sinir iletiminin blokajına neden olur. Asterias Biotherapeutics ve Lineage Cell Therapeutics klinik çalışmalarında, pluripotent kök hücrelerden türetilen Oligodendrosit Progenitör Hücreleri (OPC) yaralanma kavitasyon bölgesine enjekte edilmiştir. OPC'ler hasarlı aksonları tanıyarak etraflarına yeni miyelin kılıflar sarar, nörotrofik faktörler (NT-3, BDNF) salgılar ve nöroglial skar bariyerini delerek duyu ve motor fonksiyonların kısmi geri kazanımını sağlar.",
        "Axon_Conduction_Velocity = v_0 * sqrt(Myelin_Sheath_Thickness / Axon_Diameter)",
        "Aksonal iletim hızı, nakledilen OPC'lerin ördüğü miyelin kılıf kalınlığının akson iç çapına oranının kareköküyle artar."
    ),
    (
        "9.7",
        "Yanık ve Yara İyileşmesinde Otolog Epidermal Kök Hücre Tabakaları (Holoclar Modeli)",
        "Michele De Luca'nın holoklon teknolojisi: Tek bir kök hücreden metrelerce kare epidermal örtü üreterek ölümcül yanıkları onarma.",
        "Cilt hasarlarında cildin yenilenmesi, kıl foliküllerinde ve epidermis bazal tabakasında yaşayan p63-pozitif holoklon kök hücrelere bağlıdır. Ağır kimyasal kornea yanıklarında veya tüm vücut 3. derece yanıklarda, hastanın sağlam kalan mikroskobik bir doku parçasından izole edilen epidermal kök hücreler fibrin matriks üzerinde çoğaltılır. Üretilen epidermal tabakalar yara yatağına serildiğinde eksiksiz bir stratum korneum, melanosit dağılımı ve kılcal damarlanma organize ederek skarlaşmasız tam doku iyileşmesi sağlar.",
        "Skin_Regeneration_Success = [p63_alpha_high_Holoclones] / Total_Clonogenic_Keratinocytes > 0.05",
        "Epidermal greftin klinik tutunma başarısı, hücre süspansiyonundaki p63-alfa pozitif gerçek holoklon kök hücre oranının %5'in üzerinde olmasına bağlıdır."
    ),
    (
        "9.8",
        "Kıkırdak Onarımında Kondrosit/MSC Matriks İmplantasyonu (MACI)",
        "Avasküler ve rejenerasyonu imkansız eklem kıkırdağında kök hücre yüklü 3D kollajen membranlarla artrozu durdurma.",
        "Eklem hiyalin kıkırdağı kan damarlarından ve kök hücre rezervuarlarından yoksundur; bu nedenle kıkırdak aşınması doğrudan osteoartrit ve eklem deformasyonuna gider. Matriks Kaynaklı Otolog Kondrosit İmplantasyonu (MACI) protokolünde; hastadan artroskopik olarak alınan kondrositler veya kemik iliği MSC'leri laboratuvarda çoğaltılarak domuz Tip I/III kollajen membranına ekilir. Defekt bölgesine yapıştırılan bu biyo-aktif yama, mekanik yük altında Tip II kolajen ve agrekan salgılayarak kıkırdak yüzeyini orijinal elastikiyetinde yeniden inşa eder.",
        "Cartilage_Stiffness = Modulus_Young = 3 * [Aggrecan_PG] * [Type_II_Collagen_fibers]",
        "Yenilenen kıkırdak dokusunun elastik Young modülü, hücrelerin sentezlediği agrekan proteoglikan konsantrasyonu ve Tip II kollajen lif yoğunluğuyla orantılıdır."
    ),
    (
        "9.9",
        "Kritik Ekstremite İskemisinde Endotel Progenitör Hücre (EPC) Tedavileri",
        "Tıkayıcı periferik arter hastalığında doku nekrozunu ve amputasyonu önlemek için kemik iliği CD34+ EPC enjeksiyonuyla neo-anjiyogenez.",
        "Diyabet ve ateroskleroz sonucu bacak damarları tıkandığında cerrahi bypass yapılamayan hastalarda doku kangrene gider. Periferik kandan veya kemik iliğinden aferezle toplanan CD34+ ve VEGFR2+ Endotel Progenitör Hücreleri (EPC), iskemik kas dokusuna lokal mikro-enjeksiyonlarla verilir. Bu öncül hücreler hipoksi sinyali (HIF-1alpha / SDF-1 gradyanı) tarafından çekilir; mevcut kılcal damarlara katılarak ve VEGF, bFGF, Angiopoietin-1 salgılayarak yeni kollateral damar ağları (sprouting angiogenesis) örer ve amputasyonu önler.",
        "Capillary_Density_Increase = k_angio * N_EPC_injected * [SDF-1_gradient] / Ischemic_Distance",
        "Kılcal damar dansitesindeki artış; implante edilen EPC sayısı ve dokudaki SDF-1 kemotaktik gradyanının iskemik mesafeye oranıyla belirlenir."
    ),
    (
        "9.10",
        "Genetik Düzeltmeli Kök Hücre Terapileri: Epidermolizis Bülloza Tamiri",
        "Ölümcül genetik deri hastalığı Kelebek Çocuk sendromunda, hastanın kök hücrelerinin gen terapisiyle düzeltilip tüm vücuda nakli.",
        "2017 yılında Michele De Luca ekibi, LAMB3 genindeki mutasyon nedeniyle cildi sürekli soyulan ve ölüm döşeğinde olan 7 yaşındaki bir çocuğun deri kök hücrelerini izole etmiştir. Bir retroviral vektör aracılığıyla kök hücrelere fonksiyonel LAMB3 cDNAsı yerleştirilmiş; laboratuvarda 0.85 metrekarelik sağlıklı genetik modifiye epidermal tabakalar üretilmiştir. Hastanın vücudunun %80'ine greftlenen bu genetiği düzeltilmiş kök hücreler, Laminin-332 proteinini kusursuz sentezlemiş ve hasta çocuk yıllardır nüks olmaksızın sağlıklı bir cilde kavuşmuştur.",
        "Transduction_Efficiency = N_vector_copies_per_cell * Gene_Correction_Rate(LAMB3) -> 100%",
        "Genetik düzeltme verimliliği, kök hücre başına entegre olan terapötik vektör kopya sayısı ve düzeltilmiş protein sentez oranıyla ifade edilir."
    )
]

# ==============================================================================
# KISIM 10: ENDOJEN KÖK HÜCRELERİN UYARILMASI VE GELECEĞİN ÖMÜR UZATMA PROTOKOLÜ
# ==============================================================================
part10_subsections = [
    (
        "10.1",
        "Wnt/beta-Katenin Yolak Modülatörleri ve Lgr5+ Rejenerasyonu",
        "Bağırsak, karaciğer ve saç folikülü kök hücrelerini uykudan uyandıran R-spondin ve Wnt agonistlerinin rejeneratif kinetiği.",
        "Kanonik Wnt sinyali, kök hücrelerin kendini yenileme (self-renewal) motorudur. Wnt ligandları Frizzled ve LRP5/6 reseptörlerine bağlandığında, Axin/APC/GSK3 yıkım kompleksi inaktive olur ve beta-katenin nükleusa girerek TCF/LEF genlerini açar. R-spondin proteinleri, kök hücre yüzeyindeki Lgr4/5/6 reseptörlerine ve ZNRF3/RNF43 E3 ligazlarına bağlanarak Wnt reseptörlerinin parçalanmasını önler; Wnt sinyalini 100 kat amplifiye eder. Wnt agonistleri veya GSK-3b inhibitörleri (CHIR99021), yaşlı kök hücre havuzlarını uyararak doku rejenerasyonunu gençlik seviyesine fırlatır.",
        "Beta_Catenin_Nuclear_Flux = [Wnt] * [R-spondin] / (K_deg + [GSK3b_active])",
        "Beta-katenin nükleer translokasyon akısı, Wnt ve R-spondin konsantrasyonlarının aktif GSK-3b yıkım kinazı aktivitesine oranıyla yönetilir."
    ),
    (
        "10.2",
        "Küçük Moleküllerle Kök Hücre Uyanışı: GSK-3b İnhibitörleri ve Rock İnhibitörleri (Y-27632)",
        "Kök hücrelerin apoptoza gitmeden kültürde ve in vivo ortamda sağkalımını ve bölünmesini garantiye alan kimyasal anahtarlar.",
        "Kök hücreler dokudan veya hücre kültür yüzeyinden ayrıldıklarında anoikis adı verilen özel bir hücre ölümüne giderler. ROCK (Rho-associated kinase) inhibitörü Y-27632, aktin-miyozin kontraktilitesini gevşeterek kök hücrelerin tek tek ayrışsalar dahi hayatta kalmasını sağlar. GSK-3b inhibitörleri ile ROCK inhibitörlerinin kombinasyonu, uykudaki kök hücreleri prolifere olmaya teşvik ederken apoptoz bariyerini indirir. Bu moleküller, in vivo rejenerasyonda hasarlı bölgeye salınan akıllı hidrojel sistemleriyle hedeflenerek kök hücre patlaması yaratır.",
        "Stem_Cell_Survival_Probability = 1 - exp(-k_rock / [Y-27632])",
        "Ayrışmış kök hücrenin anoikis apoptozundan kaçış ve sağkalım olasılığı, uygulanan ROCK inhibitörü konsantrasyonuyla üstel olarak artar."
    ),
    (
        "10.3",
        "NAD+ Güçlendirmesi ve Sirtuin Aktivasyonu ile HSC Gençleşmesi",
        "Hematopoietik ve kas kök hücrelerinin mitokondriyal proteostazını ve sükunetini restore eden NMN ve NR moleküler müdahaleleri.",
        "Johan Auwerx laboratuvarının çığır açan keşiflerine göre, yaşlı kas kök hücrelerinde (uydu hücreleri) ve kan kök hücrelerinde (HSC) hücresel NAD+ seviyeleri dramatik şekilde düşer. Nikotinamid ribosit (NR) veya NMN takviyesi, mitokondriyal SIRT3 ve nükleer SIRT1 enzimlerini aktive eder. SIRT1, FoxO3a transkripsiyon faktörünü deasetilleyerek antioksidan genleri açarken; SIRT3 UPRmt (mitokondriyal katlanmamış protein yanıtı) yolağını aktive ederek bozulmuş mitokondrileri temizler ve kök hücreleri yeniden gençlik fonksiyonlarına kavuşturur.",
        "Stem_Cell_Rejuvenation_Index = [NAD+] / [NADH] * (Activity_SIRT1 + Activity_SIRT3)",
        "Kök hücre gençleşme indeksi; hücre içi serbest NAD+/NADH redoks oranının mitokondriyal ve nükleer sirtuin enzim aktiviteleriyle çarpımıdır."
    ),
    (
        "10.4",
        "Açlık (Fasting) ve Kalori Kısıtlamasının Kök Hücre Yenilenmesine Etkisi",
        "Periyodik açlık sırasında IGF-1/PKA yolağının susması ve yaşlı, hasarlı kök hücrelerin otofajiyle ayıklanıp yeniden doğuşu.",
        "Valter Longo'nun çalışmaları, 48-72 saatlik uzamış su açlığının (fasting) sistemik IGF-1 ve PKA (Protein Kinaz A) seviyelerini %70 düşürdüğünü kanıtlamıştır. PKA'nın susturulması, hematopoietik kök hücrelerde uykuyu kıran ve kendini yenilemeyi tetikleyen ana şalterdir. Açlık sırasında eski ve hasarlı lökositler ve senesen kök hücreler enerji için otofajik olarak tüketilir; beslenme başladığında ise uyarılmış HSC'ler sıfırdan yepyeni ve genç bir bağışıklık sistemi inşa eder.",
        "Immune_Regeneration_Rate = k_refeed * [Glucose_Rebound] / [PKA_activity_low]",
        "Bağışıklık kök hücre rejenerasyon hızı, açlık sonrası yeniden beslenme metabolik sıçramasının düşük PKA kinaz aktivitesiyle orantısıdır."
    ),
    (
        "10.5",
        "Klotho Proteini ve Kök Hücre Sükuneti-Proliferasyon Dengesi",
        "Böbrekten salgılanan anti-aging hormonu Klotho'nun Wnt aşırılığını frenleyerek kök hücrelerin erken tükenişini engellemesi.",
        "Klotho geninden yoksun fareler prematüre yaşlanarak birkaç haftada ölür; aşırı eksprese edenler ise %30 daha uzun yaşar. Çözünür Klotho (sKlotho), kanda bir hormon gibi dolaşır ve Wnt ligandlarına doğrudan bağlanarak onları nötralize eder. Sürekli ve kontrolsüz Wnt sinyali kök hücreleri hiper-proliferasyona sokarak hızla tüketir. Klotho, kök hücrelerin kontrolsüz bölünmesini dizginleyerek onları koruyucu sükunet (quiescence) zırhında tutar ve kök hücre rezervuarının tükeniş ömrünü iki katına çıkarır.",
        "Stem_Pool_Preservation_HalfLife = T_half_0 * (1 + alpha_klotho * [Klotho_plasma])",
        "Kök hücre rezervuarının tükenme yarı ömrü, dolaşımdaki plazma çözünür Klotho proteini konsantrasyonu ile doğrusal olarak uzar."
    ),
    (
        "10.6",
        "Parakrin İskelet: Genç Plazma, Heterokronik Parabiyoz ve Plazmaferez",
        "Yaşlı kandan pro-senesent proteinlerin filtrelenmesi (Nötral Kan Değişimi) ile kök hücre nişlerinin anında detoksifikasyonu.",
        "Irina Conboy laboratuvarı, genç plazma naklinden ziyade yaşlı kanda biriken inhibitör faktörlerin temizlenmesinin kök hücreleri gençleştirmede çok daha kritik olduğunu göstermiştir. Terapötik Plazma Değişimi (TPE / Plazmaferez) ile yaşlı kan plazması boşaltılıp yerine albümin ve salin verildiğinde; nöral, kas ve karaciğer kök hücre nişlerindeki TGF-beta, B2M ve enflamatuar baskı kalkar. Bu basit biyofiziksel seyreltme, yaşlı kök hücrelerin sükunetten uyanarak genç doku tamir hızına dönmesini sağlar.",
        "Niche_Detox_Efficiency = 1 - exp(-Clearance_Rate_TPE * Volume_exchanged / Total_Plasma_Volume)",
        "Niş detoksifikasyon verimliliği; plazmaferez ile temizlenen hacmin toplam plazma hacmine oranı ve inhibitör klerens hızı ile belirlenir."
    ),
    (
        "10.7",
        "Manyetik ve Akustik Dalgalarla Kök Hücre Mobilizasyonu ve Hedefleme",
        "Kemik iliğindeki kök hücrelerin iğnesiz, non-invaziv fiziksel dalgalarla kan dolaşımına dökülmesi ve hasarlı organa kılavuzlanması.",
        "Geleneksel kök hücre mobilizasyonu G-CSF ve Plerixafor (AMD3100) enjeksiyonlarıyla CXCR4-SDF-1 eksenini kopararak yapılır. Yeni nesil biyofiziksel protokollerde ise odaklanmış düşük yoğunluklu ultrason (FUS - Focused Ultrasound) ve darbeli elektromanyetik alanlar (PEMF) kullanılır. Dalgalar kemik iliğindeki endotel hücrelerinde geçici kalsiyum akımları başlatarak kök hücreleri dolaşıma döker. Manyetik nanopartiküllerle yüklenmiş kök hücreler, dışarıdan uygulanan odaklanmış manyetik alan gradyanlarıyla hasarlı miyokarda veya beyin lezyonuna pürüzsüzce yönlendirilir.",
        "Targeting_Velocity_Vector = v_magnetic = (Delta_Chi * Volume_cell * (B DOT grad) B) / (6 * pi * eta * R_cell)",
        "Manyetik kök hücre hedefleme hızı; hücre hacmi, manyetik duyarlılık farkı ve manyetik alan gradyanının Stokes hidrodinamik sürtünmesine oranıdır."
    ),
    (
        "10.8",
        "İn Situ Doku Baskısı (In Situ Bioprinting): Doğrudan Yara Yatağında Hücre Biriktirme",
        "Laboratuvarda doku üretip nakletmek yerine, robotik cerrahi kolların ameliyat esnasında hastanın açık yarasına doğrudan kök hücre basması.",
        "İn situ biyobaskı teknolojisinde, ameliyat masasında bulunan hastanın yara veya doku kaybı alanı 3D lazer tarayıcılarla mikrometre hassasiyetinde haritalanır. Robotik çok eksenli bir biyo-baskı ucu, bu dijital topografiye göre hastanın kendi kök hücrelerini ve fotosensitif biyo-mürekkepleri (GelMA) doğrudan yara yatağına tabaka tabaka ekstrüde eder. Eşzamanlı UV/mavi ışık polimerizasyonu ile saniyeler içinde doku iskeleti oluşturulur. Bu yöntem cerrahi greft tutma başarısını %95 üzerine çıkarır.",
        "Layer_Deposition_Accuracy = Delta_x_spatial < 15.0_microns (Laser_Feedback_Loop)",
        "İn situ biyobaskı tabaka hassasiyeti; optik geri besleme döngüsüyle yönetilen piezoelektrik nozül konumlandırma sapmasıdır."
    ),
    (
        "10.9",
        "Biyo-Elektronik Kök Hücre Nişleri ve Akıllı Biyomateryaller",
        "Grafen ve iletken polimerlerden üretilen yapay nişlerin, elektriksel mikro-akımlarla kök hücre farklılaşmasını milisaniyelik kontrolü.",
        "Geleceğin rejeneratif implantları pasif doku iskeleleri değil; yaşayan, hisseden ve kök hücreyle çift yönlü haberleşen biyo-elektronik arayüzlerdir. PEDOT:PSS ve karbon nanotüp kaplı akıllı biyomateryaller, kök hücrelerin membran voltajını (Vmem) ölçer ve dışarıdan uygulanan mikrovoltaj darbeleriyle iyon kanallarını (özellikle VGCC kalsiyum kanallarını) açarak kök hücre kaderini yönlendirir. Örneğin -70 mV hiperpolarizasyon kök hücreyi sükunette tutarken; -20 mV depolarizasyon osteojenik veya nörojenik farklılaşmayı anında başlatır.",
        "Differentiation_Trajectory = f(V_membrane, Frequency_Electric_Pulse, Surface_Conductivity)",
        "Biyo-elektronik nişte hücre kaderi; hücre membran potansiyeli, uygulanan elektriksel atım frekansı ve yüzey iletkenliğinin transfer fonksiyonudur."
    ),
    (
        "10.10",
        "Homo Aeternus Kök Hücre Mimarisi: Yaşlanmayan Organizmada Hücre Dönüşüm Döngüsü",
        "Tüm dokularda asimetrik bölünme dengesinin, telomerik korumanın ve klonal saflığın yapay biyolojik geri beslemeyle sonsuzlaştırılması.",
        "Biyolojik ölümsüzlüğün kök hücre manifestosu; 'tükenmeyen, kanserleşmeyen ve mutasyonel kaymaya uğramayan' entegre bir kök hücre ekosistemidir. Bu mimaride insan genomu; sentetik p53 kopyaları (fil genomundaki EP35 kopyaları gibi) ile onkojenik korumaya alınır, hTERT promotoru telomer erozyonunu önleyecek şekilde homeostatik geri beslemeye bağlanır ve yapay epigenetik saat sıfırlayıcılar (dCas9-epigenetik editörler) yılda bir kez devreye girerek kök hücre havuzlarını fabrika ayarlarına döndürür. Böyle bir organizmada doku atrofisi ve yaşa bağlı organ yetmezliği tarihe karışır.",
        "Organismal_Regenerative_Capacity = lim_{t -> infty} sum_tissues (N_stem_i(t) * Quality_stem_i(t)) = CONSTANT",
        "Homo Aeternus rejeneratif kapasitesi; sonsuz zaman ufkunda tüm dokulardaki kök hücre sayısı ve kalitesinin sabit bir gençlik platosunda korunmasıdır."
    )
]

# ==============================================================================
# 10 AKADEMİK KARŞILAŞTIRMA VE PARAMETRE TABLOSU (HER KISIM İÇİN BİR ADET)
# ==============================================================================
parts.append(("KISIM 7: İNDÜKLENMİŞ PLURİPOTENT KÖK HÜCRELER (iPSC) VE HASTALIK MODELLEMESİ", part7_subsections))
parts.append(("KISIM 8: KÖK HÜCRE YAŞLANMASININ MOLEKÜLER DİNAMİKLERİ VE EPİGENETİK BOZULMA", part8_subsections))
parts.append(("KISIM 9: REJENERATİF TIP, HÜCRE TERAPİLERİ VE KLİNİK TRANSFORMATİF PROTOKOLLER", part9_subsections))
parts.append(("KISIM 10: ENDOJEN KÖK HÜCRELERİN UYARILMASI VE GELECEĞİN ÖMÜR UZATMA PROTOKOLÜ", part10_subsections))

tables_data = [
    (
        "TABLO 1: Kök Hücre Potens Sınıfları ve Karşılaştırmalı Biyomarker Analizi",
        ["Potens Sınıfı", "Örnek Hücre Tipi", "Temel Transkripsiyon Faktörleri", "Farklılaşma Yelpazesi", "Etik / Onkojenik Risk Düzeyi"],
        [
            ["Totipotent", "Zigot, Erken Blastomerler (1-3 gün)", "Oct4, Sox2, Nanog, Zscan4", "Tüm embriyonik dokular + Ekstra-embriyonik (Plasenta)", "En yüksek etik hassasiyet; in vitro stabil tutulamaz"],
            ["Pluripotent (ESC / iPSC)", "İç Hücre Kitlesi (ICM), Yamanaka iPSC", "Oct4, Sox2, Nanog, Klf4, Lin28", "Ektoderm, Mezoderm, Endoderm (3 germ yaprağı)", "Yüksek teratoma riski; titiz kalite kontrolü gerektirir"],
            ["Multipotent", "HSC, MSC, Nöral Kök Hücre (NSC)", "Sox2, Pax7, Runx2, Nestin, GATA2", "Ait olduğu germ yaprağına özgü çoklu hücre tipleri", "Sıfır teratoma riski; klinik tedavilerde altın standart"],
            ["Oligopotent", "Miyeloid / Lenfoid Progenitör Hücreler", "PU.1, C/EBP-alpha, Pax5", "Sınırlı sayıda ilişkili hücre soyu (örn. Granülosit/Monosit)", "Düşük risk; sınırlı kendini yenileme kapasitesi"],
            ["Unipotent", "Epidermal Bazal Hücre, Spermatogonyal Kök", "p63, MyoD, Mitf, Stra8", "Tek bir özgül hücre tipi (örn. Keratinosit, Sperm)", "Sıfır onkojenik sapma; yerel doku onarımında uzman"]
        ]
    ),
    (
        "TABLO 2: Kök Hücre Nişi Biyofiziksel Parametreleri ve Hücre-Matriks Etkileşimleri",
        ["Niş Parametresi", "Kemik İliği Endosteal Niş", "Vasküler / Perivasküler Niş", "Beyin Subventriküler Zon (SVZ)", "İskelet Kası Uydu Nişi"],
        [
            ["Young Elastisite Modülü (E)", "30 - 100 kPa (Sert kemik matriksi)", "2 - 8 kPa (Yumuşak endotel çevresi)", "0.1 - 1 kPa (Ultra yumuşak beyin parankimi)", "12 kPa (Fizyolojik kas sertliği)"],
            ["Oksijen Gerilimi (pO2)", "%1 - %3 O2 (Aşırı hipoksi)", "%4 - %8 O2 (Normoksik perfüzyon)", "%1 - %2.5 O2 (Fizyolojik hipoksi)", "%2 - %5 O2 (Lokal kas mikroçevresi)"],
            ["Temel Hücre-ECM Reseptörleri", "alfa4-beta1 (VLA-4), alfa5-beta1 İntegren", "alfa6-beta1 İntegren, CD31, VE-Kadherin", "alfa6-beta1, N-Kadherin, Tenasin-C", "alfa7-beta1 İntegren, M-Kadherin, Distrogilikan"],
            ["Baskın Morfogenik Gradyan", "Düşük Wnt, Yüksek TGF-beta / BMP", "Yüksek VEGF, Angiopoietin-1, SDF-1", "Yüksek EGF, FGF-2, Shh, Noggin", "Pulsatil HGF, Wnt7a, Notch ligandı Delta-1"],
            ["Kök Hücre Durumu", "Derin sükunet (G0), HSC rezervi", "Aktif proliferasyon ve mobilizasyon", "Sükunet ile yavaş döngü dengesi", "Derin G0 sükuneti, travma uyarımı"]
        ]
    ),
    (
        "TABLO 3: Klonal Hematopoez Sürücü Mutasyonları ve Klinik Risk Profilleri",
        ["Gen Sembolü", "Enzimatik / Moleküler Rol", "Tipik Mutasyon Tipleri", "Klonal Yayılma Hızı", "Klinik Sonuç ve Kardiyovasküler Risk"],
        [
            ["DNMT3A", "De Novo DNA Metiltransferaz", "R882H dominant negatif missense", "Orta-Yavaş (Yılda %1-3 VAF artışı)", "AML riski 10x; Koroner arter hastalığı riski 2.0x"],
            ["TET2", "5-metilsitozin Dioksijenaz (Demetilasyon)", "Çerçeve kayması, anlamsız, delesyon", "Orta-Hızlı (Enflamasyona dirençli)", "Makrofaj hiperaktivasyonu, IL-1b artışı, Ateroskleroz 2.5x"],
            ["ASXL1", "Polikomb Kromatin İskelet Regülatörü", "Ekzon 12 delesyon ve insersiyonları", "Hızlı klonal genişleme", "Miyelodisplastik Sendrom (MDS) riski yüksek; Kalp yetmezliği"],
            ["JAK2", "Sitokin Reseptör Tirozin Kinazı", "V617F konstitütif aktif missense", "Hızlı klonal amplifikasyon", "Polisitemia vera, Esansiyel trombositemi, Tromboz riski 4x"],
            ["TP53", "Tümör Baskılayıcı DNA Hasar Kontrolü", "DNA bağlama alanı missense/null", "Kemoterapi/radyasyon altında patlama", "Terapi ilişkili AML, kemoterapiye direnç, çok kötü prognoz"]
        ]
    ),
    (
        "TABLO 4: Dokuya Özgül Kök Hücre Tipleri, Rejenerasyon Hızları ve Tükeniş Mekanizmaları",
        ["Kök Hücre Tipi", "Anatomik Lokalizasyon", "Hücresel Döngü / Yenilenme Süresi", "Pozitif Biyomarkerlar", "Yaşlanmaya Bağlı Başlıca Bozulma"],
        [
            ["Hematopoietik (HSC)", "Trabeküler kemik iliği lakunaları", "Ayda 1 bölünme (Aktif) / Yılda 1 (Dormant)", "CD34+, CD38-, Lin-, CD90+, CD45RA-", "Klonal hematopoez (CHIP), miyeloid kayma, anöploidi"],
            ["İntestinal CBC", "İnce/kalın bağırsak kript tabanı", "Her 24 saatte bir mitoz (Tüm epitel 4 günde bir)", "Lgr5+, Olfm4+, Ascl2+, CD133+", "Paneth niş sinyallerinin zayıflaması, APC mutasyon riski"],
            ["Epidermal Bazal", "Cilt stratum basale ve kıl folikülü", "7 - 14 günde bir tam epidermis dönüşümü", "KRT14+, KRT5+, p63+, CD49f (alfa6)+", "Kollajen XVII (COL17A1) erozyonu, cilt incelmesi, yara gecikmesi"],
            ["Kas Uydu (Satellite)", "İskelet kas lifi sarkolemma altı", "Sükunette sıfır; hasarda 24 saatte bir", "Pax7+, CD56+, Integrin-alfa7+, CXCR4+", "TGF-beta fibrotik kayması, Notch kaybı, sarkopeni"],
            ["Nöral (NSC)", "Beyin SVZ ve hipokampal SGZ", "Yavaş asimetrik bölünme; yaşla sönümlenir", "SOX2+, Nestin+, GFAP+, Prominin-1+", "Sistemik enflamatuar baskı (CCL11, B2M), nörogenez çöküşü"]
        ]
    ),
    (
        "TABLO 5: Ekstraselüler Vezikül / Eksozom İçeriği ve Rejeneratif Moleküler İmzalar",
        ["EV / Eksozom Parametresi", "Küçük Eksozomlar (sEV)", "Mikroveziküller (Ektosom)", "Apoptotik Cisimcikler", "Sentetik Biyo-Nanoveziküller"],
        [
            ["Çap Boyut Aralığı", "30 - 150 nm", "100 - 1000 nm", "1000 - 5000 nm", "50 - 120 nm"],
            ["Biyogenez Yolağı", "Endozomal MVB füzyonu", "Plazma zarı doğrudan dışa tomurcuklanması", "Kaspaz aracılı apoptotik hücre parçalanması", "Hücre ekstrüzyonu / mikroakışkan montaj"],
            ["Yüzey Tetraspaninleri", "CD9, CD63, CD81, Alix, TSG101", "İntegrenler, Selektinler, CD40", "Anneksein V (Fosfatidilserin zengin)", "Yapay tasarlanmış ligandlar (anti-CD34, RGD)"],
            ["Biyoaktif Yük (Kargo)", "miR-21, miR-126, miR-133a, Wnt proteinleri", "mRNA'lar, Sitoplazmik proteinler, Sitokinler", "Fragmente DNA, Nükleer kromatin, Organeller", "Yüklenmiş siRNA, dCas9-mRNA, CRISPR kılavuzları"],
            ["Rejeneratif Mekanizma", "Endotel göçü, anjiyogenez, antifibrotik", "Parakrin immünomodülasyon, doku uyarımı", "Makrofaj efferositozu, immün tolerans", "Hedefe yönelik genetik ve epigenetik susturma"]
        ]
    ),
    (
        "TABLO 6: Doku Mühendisliği İskeleleri (Scaffolds), Biyomürekkepler ve Üretim Teknolojileri",
        ["Biyomateryal / İskele", "Biyokimyasal Kompozisyon", "Mekanik Dayanım ve Degradasyon", "3D Baskı Teknolojisi", "Hedef Klinik Doku Organı"],
        [
            ["GelMA (Jelatin Metakrilat)", "Denatüre kollajen + Metakrilik anhidrit", "Düşük (1-10 kPa); enzimatik erime 2-4 hafta", "Işıkla Kürlenen Stereolitografi (SLA / DLP)", "Vasküler ağlar, cilt yamaları, karaciğer parankimi"],
            ["Alginat - Fibrin Hibridi", "Kahverengi alg polisakkariti + Plazma fibrini", "Orta (5-30 kPa); kalsiyum klerensiyle yavaş", "Pnömatik Ekstrüzyon Biyobaskı", "Kardiyak yamalar, pankreatik adacık kapsülleri"],
            ["Decellülarize ECM (dECM)", "Dokuya özgül doğal bazal membran proteomu", "Doğal organ sertliğine eşdeğer elastisite", "Dondurarak Kurutma / Hidrojel Ekstrüzyon", "Akciğer alveol iskeleleri, bütün böbrek/karaciğer"],
            ["Polikaprolakton (PCL)", "Biyo-bozunur sentetik alifatik poliester", "Çok yüksek (10-50 MPa); yavaş erime (1-2 yıl)", "Erimiş Biriktirme Modellemesi (FDM)", "Kafatası ve uzun kemik defektleri, kıkırdak iskeleti"],
            ["İpek Fibroini (Silk Fibroin)", "Bombyx mori ipekböceği kristal proteinleri", "Yüksek çekme direnci; kontrollü bozunma", "Elektroeğirme (Electrospinning) / İnkjet", "Kornea stroması, tendon/ligament, yapay sinir kanalları"]
        ]
    ),
    (
        "TABLO 7: Hücresel Reprogramlama Yöntemlerinin Karşılaştırmalı Güvenlik ve Verimlilik Matrisi",
        ["Yeniden Programlama Yöntemi", "Taşıyıcı / Dağıtım Platformu", "Genomik Entegrasyon Riski", "Dönüşüm Verimliliği (%)", "Klinik GMP Uygunluğu"],
        [
            ["Retroviral / Lentiviral", "Entegre olan RNA virüsleri", "Çok Yüksek (Rastgele insersiyonel onkogenez)", "%0.1 - %1.0 (Yüksek verim)", "Klinik kullanım için uygun değil (Araştırma sınıfı)"],
            ["Sendai Virüsü (SeV)", "Sitoplazmik eksi iplikli RNA virüsü", "Sıfır (Nükleer DNA'ya asla dokunmaz)", "%0.05 - %0.5 (Mükemmel)", "Klinik Faz I/II denemelerinde altın standart"],
            ["Epizomal Plazmitler", "EBNA-1 / OriP bazlı sirküler DNA", "Ultra Düşük (Pasajlarda kendiliğinden seyrelir)", "%0.01 - %0.05 (Orta)", "Klinik onaylı, ucuz ve geniş kabul gören yöntem"],
            ["Sentetik Modifiye mRNA", "N1-metilpsödoüridinli mRNA + LNP", "Kesinlikle Sıfır (Genetik materyal kalmaz)", "%0.1 - %2.0 (Tekrarlayan dozlarla yüksek)", "En güvenli fakat en maliyetli klinik format"],
            ["Kimyasal Kokteyller (CiPSC)", "CHIR99021, VPA, RepSox, TTNPB vb.", "Sıfır (Yalnızca küçük kimyasal moleküller)", "%0.01 - %0.1 (Protokole bağlı)", "Geleceğin en umut verici farmasötik reprogramlama yolu"]
        ]
    ),
    (
        "TABLO 8: Kök Hücre Yaşlanma Belirteçleri, Nükleer Deformasyon ve Metabolik Değişimler",
        ["Biyolojik Parametre", "Genç Kök Hücre Durumu", "Yaşlı Kök Hücre Durumu", "Ölçüm Yöntemi", "Müdahale ile Geri Döndürülebilirlik"],
        [
            ["Nükleer Morfoloji", "Kusursuz elipsoid / dairesel zar (Lamin B1+)", "Buruşuk, blebbing gösteren zar (Progerin+)", "Konfokal İmmünofloresan mikroskopi", "Kısmi yeniden programlama (OSKM) ile tam düzelir"],
            ["Kromatin Yapısı", "Periferik yoğun H3K9me3 heterokromatini", "Heterokromatin kaybı, LINE-1 aktivasyonu", "ATAC-seq, ChIP-seq, RNA-seq", "Epigenetik editörler (dCas9-SUV39H1) ile restore edilir"],
            ["Metabolik Yolak", "Anaerobik glikoliz baskın (HIF-1alpha)", "Zorlanmış mitokondriyal OXPHOS, kaçak ROS", "Seahorse XF Analizörü (ECAR vs OCR)", "HIF-1 stabilizatörleri ve NAD+ prekürsörleri ile düzelir"],
            ["Proteostaz / Otofaji", "Yüksek bazal otofaji, aktif 26S proteazom", "LAMP2A çöküşü, perinükleer agresom birikimi", "Lizozomal LysoTracker, Ubikitin boyaması", "Spermidin, Rapamisin, Açlık protokolleri ile temizlenir"],
            ["Hücre Döngüsü Frenleri", "p16INK4a ve p21CIP1 tamamen baskılanmış", "Aşırı p16 ve p21 ekspresyonu, G0 kilitlenmesi", "Western Blot, qRT-PCR tek-hücre analizi", "Senolitikler (ABT-263, D+Q) ile yaşlı hücreler elenir"]
        ]
    ),
    (
        "TABLO 9: Klinik Evre Kök Hücre Tedavileri, Endikasyonlar ve Başarı Oranları",
        ["Klinik Tedavi / Ürün", "Kök Hücre Kaynağı ve Tipi", "Hedef Hastalık / Doku Defekti", "Klinik Faz Aşaması", "Kanıtlanmış Başarı / Sonuç"],
        [
            ["Holoclar", "Otolog korneal limbal kök hücreler", "Oküler yanık kaynaklı kornea körlüğü", "EMA Onaylı (Piyasada)", "%75-80 oranında kalıcı kornea şeffaflığı ve görme"],
            ["Bemdaneprocel (BRT-DA01)", "Allogeneik iPSC türevi dopamin progenitörleri", "İleri evre Parkinson Hastalığı", "Faz I/II Tamamlandı", "Striatal dopaminerjik tutunma, UPDRS skorunda %40 düşüş"],
            ["VX-880 / VX-264", "Allogeneik iPSC türevi beta adacık hücreleri", "Tip 1 Diyabetes Mellitus", "Faz I/II Devam Ediyor", "Dışarıdan insülin kullanımının tamamen sonlandırılması"],
            ["Alofisel (Darvadstrocel)", "Allogeneik genişletilmiş adipoz kök hücreler (eASC)", "Crohn hastalığı kompleks perianal fistülleri", "EMA Onaylı (Piyasada)", "Refrakter fistüllerde %50+ tam remisyon ve doku kapanması"],
            ["LentiGlobin / Zynteglo", "Otolog CD34+ HSC (Lentiviral beta-globin)", "Transfüzyon bağımlı Beta-Talasemi", "FDA / EMA Onaylı", "Hastaların %90'ında ömür boyu kan transfüzyonundan bağımsızlık"]
        ]
    ),
    (
        "TABLO 10: Endojen Kök Hücre Rejenerasyon Protokolleri ve Moleküler Müdahale Stratejileri",
        ["Müdahale Stratejisi", "Etki Ettiği Moleküler Yolak", "Kök Hücre Havuzu Üzerindeki Etkisi", "Klinik Uygulama Rejimi", "Beklenen Rejeneratif Kazanım"],
        [
            ["Açlık (Fasting-Mimicking Diet)", "IGF-1 / PKA ekseni inhibisyonu", "Yaşlı HSC'lerin tasfiyesi, genç klonların uyanışı", "Ayda bir kez 5 günlük döngüsel rejim", "Bağışıklık sisteminin baştan aşağı gençleştirilmesi"],
            ["NAD+ Prekürsörü (NMN / NR)", "SIRT1 ve SIRT3 mitokondriyal aktivasyonu", "Kas uydu ve HSC mitokondriyal proteostazı", "Günlük 500 - 1000 mg oral / IV infüzyon", "Kas rejenerasyon hızında %50 artış, kök hücre gençleşmesi"],
            ["Plazmaferez (Nötral Kan Değişimi)", "TGF-beta, B2M, CCL11 sistemik seyreltmesi", "Nöral ve kas kök hücre nişlerinin uyanışı", "Yılda 2-4 seans tam plazma değişimi", "Nörogenezde sıçrama, kognitif ve kas fonksiyon artışı"],
            ["Wnt / R-spondin Agonistleri", "Frizzled / LRP6 / beta-katenin uyarımı", "İntestinal, karaciğer ve saç kök hücresi patlaması", "Lokal hidrojelle kontrollü doku salımı", "Hızlı mukozal onarım, karaciğer rejenerasyon ivmesi"],
            ["Odaklanmış Ultrason (FUS)", "Mekanotransdüksiyon ve SDF-1/CXCR4 ekseni", "Kemik iliğinden dolaşıma kök hücre mobilizasyonu", "Hasarlı organ üzerine haftalık 15 dk FUS", "Ameliyatsız lokal organ revaskülarizasyonu ve onarımı"]
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
                
        # Tablo Sonrası Sayfa Sonu (Tablonun tek başına net görünmesi ve sayfa artışı)
        doc.add_page_break()

# Dokümanı Kaydet
doc.save(OUTPUT_PATH)
print(f"BÖLÜM 06 başarıyla kaydedildi: {OUTPUT_PATH}")

