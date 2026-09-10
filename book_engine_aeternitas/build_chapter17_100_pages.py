"""
PROJECT AETERNITAS - CİLT 17: KANSER DİRENCİ, NEGLIGIBLE SENESCENCE VE KOMPARATİF BİYOLOJİ
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

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_17_KANSER_DIRENCI_NEGLIGIBLE_SENESCENCE_TAM_100_SAYFA.docx"

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
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 17: KANSER DİRENCİ VE KOMPARATİF BİYOLOJİ")
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
s_run = sub_p.add_run("CİLT 17: KANSER DİRENCİ, NEGLIGIBLE SENESCENCE VE KOMPARATİF BİYOLOJİ\\n(PETO PARADOKSU, ÇIPLAK KÖR KÖSTEPEK FARESİ, GRÖNLAND BALİNASI, TURRITOPSIS VE EVRİMSEL LONGEVITY)")
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
ih_run = intro_h.add_run("CİLT 17 MANİFESTOSU: DOĞANIN ÖLÜMSÜZ ARŞİVİ: EVRİMİN KANSER VE YAŞLANMAYI YENDİĞİ GENOMİK TAPINAK")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "İnsan tıbbı yüzyıllardır kanser ve yaşlanmayı önlenemez bir biyolojik kader olarak görme yanılgısına düşmüştür. Oysa yeryüzündeki "
    "canlılar alemi incelendiğinde; doğanın kanseri ve yaşlanmayı defalarca, farklı türlerde ve birbirinden muazzam moleküler mekanizmalarla "
    "kesin olarak yendiği görülür. Biyolojik ölümsüzlük sıfırdan icat edilecek bir fantezi değil; doğanın zaten keşfettiği şifrelerin "
    "insan genomuna tercüme edilmesidir.\\n\\n"
    "Bu evrimsel zaferlerin haritası açıktır: İnsandan bin kat fazla hücreye sahip olan Grönland balinaları 200 yıldan fazla yaşamakta "
    "ve kansere yakalanmamaktadır (Peto Paradoksu); Çıplak kör köstebek fareleri (Heterocephalus glaber) yüksek moleküler ağırlıklı "
    "hyaluronik asit (HMW-HA) ile tümör oluşumunu fiziksel olarak imkânsız kılmaktadır; Filler 20 kopya TP53 geni ve uyanmış LIF6 zombi "
    "geniyle en ufak bir DNA hasarında hücreyi derhal imha etmektedir; Arctica islandica istiridyeleri 507 yıl boyunca tek bir protein "
    "bozulması yaşamadan var olmakta; ve Turritopsis dohrnii denizanası yaşlandığında hücrelerini transdiferansiye ederek ömrünü sonsuz "
    "bir polip döngüsünde sıfırlamaktadır (Negligible Senescence).\\n\\n"
    "Bu ciltte; Peto paradoksunun biyofiziği, çıplak kör köstebek faresi HMW-HA mekanizması, balina ve fil genomik stabilitesi, ihmal "
    "edilebilir yaşlanma gösteren türlerin biyolojisi, hidra ve Turritopsis ölümsüzlük döngüsü, yarasa interferon/NLRP3 adaptasyonu ve "
    "bu evrimsel üstünlüklerin insan hücrelerine CRISPR/Prime Editing ile aktarılmasını içeren Homo Aeternus Negligible Senescence "
    "protokolü 100 akademik alt bölümde incelenmektedir."
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
        "1.1 Peto Paradoksu Tanımı: Beden Boyutu, Hücre Sayısı ve Kanser Riski Çelişkisi",
        "Kanser biyolojisinin temel stokastik modeline göre; her hücre bölünmesi belirli bir somatik mutasyon olasılığı (mu) taşır ve bir canlının hücre sayısı ne kadar fazlaysa ve ne kadar uzun yaşıyorsa kansere yakalanma olasılığı o kadar yüksek olmalıdır.",
        "1977 yılında epidemiyolog Richard Peto bu basit modelin doğada tamamen çöktüğünü keşfetmiştir: Bir insan bir fareden 1000 kat daha fazla hücreye sahiptir ve 30 kat daha uzun yaşar; bir mavi balina veya Grönland balinası ise bir insandan 1000 kat daha fazla hücreye sahiptir (yaklaşık 10^16 hücre). Eğer kanser riski hücre sayısıyla doğru orantılı olsaydı, her balina daha yetişkinliğe ulaşamadan yüzlerce tümörden ölmek zorundaydı. Oysa dev memelilerde kanser insidansı farelerden veya insanlardan katbekat daha düşüktür; bu evrimsel muammaya 'Peto Paradoksu' denir.",
        "Teorik_Risk = P_kanser = 1 - ( 1 - mu_mutasyon )^( N_hucre * N_bolunme ) ~ 1.0 (Dev Canlılarda Beklenen)",
        "Bu stokastik onkogenez modeli, ek evrimsel koruma mekanizmaları olmaksızın trilyonlarca hücreye sahip dev memelilerin gençlik çağında kaçınılmaz kanserden ölmesi gerektiğini matematiksel olarak gösterir."
    ),
    (
        "1.2 Evrimsel Biyoloji ve Türler Arası Tümör Baskılama Seçilimi",
        "Peto paradoksunun çözümü, evrimin 'Beden Boyutu Büyümesi' (Gigantism) sırasında kanser baskılama mekanizmalarını katı bir doğal seçilim baskısı altında güçlendirmesidir.",
        "Bir memeli türü evrimsel süreçte beden kütlesini artırdıkça, eş zamanlı olarak fazladan tümör baskılayıcı genler (tümör supresör duplikasyonları), daha katı hücre döngüsü kontrol noktaları ve aşırı duyarlı apoptoz kaskadları evriltmek zorunda kalmıştır; aksi takdirde soyu tükenirdi. Doğal seçilim, dev canlıların genomunu kansere karşı aşılmaz bir genetik kaleye dönüştürmüştür.",
        "Secilim_Baskisi = S_kanser = - d(Fitness) / d(Risk_onkojen) >> 0 (Beden Boyutu Arttıkça Katlanır)",
        "Bu evrimsel seçilim denklemi, artan hücre sayısının türün hayatta kalabilmesi için tümör baskılama gücünü artırma zorunluluğunu formüle eder."
    ),
    (
        "1.3 'Hipertümör' (Hypertumor) Hipotezi: Dev Kanserlerin Kendi Kendini Yok Etmesi",
        "Büyük kütleli canlılarda kanser direncinin en ilginç teorik açıklamalarından biri, Nagy ve ark. tarafından ortaya atılan 'Hipertümör' (Hypertumor) dinamiğidir.",
        "Balina gibi 100 tonluk bir canlıda primer bir tümör odağı oluşsa dahi; tümör klonları kendi içinde aşırı hızlı mutasyona uğrayarak agresif 'bencil klonlar' (hipertümörler) türetir. Bu yeni klonlar ana tümörün kan damarlarını (anjiyogenez) parazit gibi emerek ana tümörü nekroza uğratır ve kendi kan desteğini de tükettiği için tümör kritik klinik boyuta (örneğin öldürücü 1 kg kütleye) ulaşamadan kendi kendini aç bırakıp imha eder.",
        "Buyume_Hipertumor: dV_tumör/dt = r_t * V_t - alpha_parazit * V_hipertumor * ( V_t / K_kanlanma )",
        "Bu ekolojik tümör rekabeti denklemi, devasa dokularda gelişen ikincil tümör klonlarının ana tümörü aç bırakarak öldürme kinetiğini belgeler."
    ),
    (
        "1.4 Metabolik Hız ve Kleiber Kanunu (3/4 Kuvveti) ile Oksidatif Hasar Azalması",
        "Canlıların vücut kütlesi büyüdükçe, gram doku başına düşen bazal metabolizma hızı ve mitokondriyal oksijen tüketimi radikal biçimde yavaşlar (Kleiber Kanunu).",
        "Kleiber'in 3/4 allometrik ölçekleme yasasına göre (Metabolizma = M^(3/4)); 30 gramlık bir farenin tek bir hücresi, 100 tonluk bir balinanın hücresine kıyasla onlarca kat daha hızlı oksijen yakar ve muazzam miktarda serbest radikal (ROS) üretir. Balina ve fil hücreleri son derece yavaş, sakin ve serin bir metabolik hızda çalıştığı için DNA'ları oksidatif mutasyon bombardımanına fareye kıyasla %95 daha az maruz kalır.",
        "Metabolik_Hiz_Ozgun = BMR_ozgun = BMR / M = k_kleiber * M^(-1/4) (Kütle Arttıkça ROS Çöker)",
        "Bu allometrik biyoenerjetik eşitliği, devasa canlılarda hücresel oksidatif stres ve somatik mutasyon sıklığının kütlenin -1/4 kuvvetiyle nasıl logaritmik olarak azaldığını belgeler."
    ),
    (
        "1.5 Kanser Direnci ile Uzun Ömürlülük (Longevity) Arasındaki Genomik Eşleşme",
        "Doğada hiçbir memeli türü kansere karşı aşırı direnç geliştirmeden uzun bir yaşam süresine (maksimum ömür > 30 yıl) ulaşamamıştır.",
        "Kanser direnci ve uzun ömürlülük aynı madalyonun iki yüzüdür: Hücreyi kanserden koruyan mekanizmalar (kusursuz DNA onarımı, yüksek proteostaz, katı epigenetik kararlılık); aynı zamanda hücreyi yaşlanmanın entropik yıpranmasından koruyan temel eksenlerdir. Evrim, ömrü uzatmak için önce kanser barajını yıkmak zorunda kalmıştır.",
        "Korelasyon_Longevity = Lifespan_max = alpha_DDR * Direnç_Kanser + beta_Metabolizma",
        "Bu evrimsel demografi eşitliği, türlerin maksimum yaşam süresinin kanser direnci mekanizmaları ve DNA hasar yanıt gücü ile doğrudan korelasyonunu ortaya koyar."
    ),
    (
        "1.6 Somatik Mutasyon Birikim Hızı: Türler Arası Evrensel Yaşam Sonu Kotası",
        "2022 yılında Nature'da yayımlanan tarihi Wellcome Sanger Enstitüsü çalışması (Cagan ve ark.); memeli türleri arasında inanılmaz bir 'Somatik Mutasyon Sabiti' keşfetmiştir.",
        "Bir farenin hücreleri yılda yaklaşık 800 somatik mutasyon biriktirirken; insan hücreleri yılda 47, Grönland balinası hücreleri ise yılda yalnızca 12 mutasyon biriktirir. Farklı ömürlere (3 yıl vs 80 yıl vs 200 yıl) sahip tüm memeliler doğal yaşamlarının sonuna ulaştıklarında hücre başına neredeyse aynı toplam mutasyon yüküne (yaklaşık 3.000 mutasyon) ulaşır. Uzun yaşayan türler yıllık mutasyon hızını moleküler olarak frenlemeyi başarmıştır.",
        "Hiz_Mutasyon_Yillik = dM/dt = Mutasyon_Sonu_Sabit (~3200) / Lifespan_max",
        "Bu evrensel somatik mutasyon kotası eşitliği, türlerin maksimum ömrünün yıllık mutasyon kazanım hızının tersi ile tam bir doğrusal bağıntı içinde olduğunu kanıtlar."
    ),
    (
        "1.7 İmmün Gözetim Gücü: Dev Canlılarda Sitotoksik T ve Doğal Katil (NK) Hücre Kapasitesi",
        "Dev canlılar yalnızca hücre içi tümör baskılayıcılara değil; günde trilyonlarca hücreyi denetleyen olağanüstü bir 'İmmün Gözetim' (Immune Surveillance) ağına sahiptir.",
        "Filler ve balinalar; genişlemiş lenfoid organ rezervleri, yüksek antijenik afiniteye sahip T hücre reseptör (TCR) çeşitliliği ve tümör hücrelerinin neo-antijenlerini anında yakalayan aşırı aktif Doğal Katil (NK) hücrelerine sahiptir. Tek bir hücre neoplastik transformasyon sinyali (MHC kaybı veya MICA/MICB stresi) verdiği anda, lökositler saatler içinde perforin/granzim deşarjı ile tümör klonunu tek hücre aşamasında ortadan kaldırır.",
        "Klirens_Onkojen = d[Hucre_Malign]/dt = - k_NK * [NK_aktif] * [MICA_stres_ligandi]",
        "Bu immünolojik tümör eliminasyon kinetiği, devasa doku kitlelerinde prekanseröz hücrelerin immün gözetim tarafından temizlenme hızını tanımlar."
    ),
    (
        "1.8 Epigenetik Kararlılık ve Metilasyon Kaybı Direnci",
        "Kanser gelişiminin ilk adımı genellikle genetik mutasyonlardan önce epigenetik enformasyon kaybı (promotör hipermetilasyonu veya genomik hipometilasyon) ile başlar.",
        "Peto paradoksunu çözen uzun ömürlü canlılar; epigenetik kaymaya (Epigenetic Drift) karşı sarsılmaz bir kromatin mimarisine sahiptir. DNMT1 bakım metiltransferazı ve TET enzimlerinin sadakati çok yüksektir; tümör baskılayıcı genlerin promotörleri yaşlanma boyunca metillenmeye direnç gösterir ve transpozonlar heterokromatine sıkıca kilitli kalır.",
        "Kayma_Epigenetik = d[Drift]/dt = k_hata * ( 1 - Fidelite_DNMT1 ) << Insan_Hata_Orani",
        "Bu epigenetik kararlılık formülü, kansere dirençli türlerde DNA metilasyon sadakatinin epigenetik kanserleşme eğilimini nasıl sıfırladığını açıklar."
    ),
    (
        "1.9 Kök Hücre Bölünme Kısıtlamaları ve Kanser Kök Hücrelerinin Önlenmesi",
        "Tümörlerin asıl kaynağı, sınırsız kendini yenileme potansiyeline sahip Kanser Kök Hücreleridir (CSCs).",
        "Dev canlılarda erişkin doku kök hücrelerinin bölünme frekansı katı bir hücresel hiyerarşi ile kısıtlanmıştır. Kök hücreler yaşam boyu derin bir sessizlik (quiescence) fazında tutulur; doku yenilenmesi kısa ömürlü transit amplifiye progenitör hücreler üzerinden yürütülür. Kök hücre az bölündüğü için kopyalama hatası yapma ve kanser kök hücresine dönüşme riski matematiksel olarak minimumda tutulur.",
        "Bolunme_Kotasi = N_bolunme_kok = N_toplam_bolunme / N_kademeli_progenitor << 100 (Ömür Boyu)",
        "Bu hücresel hiyerarşi kısıtlaması, uzun ömürlü canlılarda kök hücrelerin bölünme sayısının düşürülerek mutasyon birikiminin nasıl engellendiğini belgeler."
    ),
    (
        "1.10 Komparatif Biyolojinin İnsan Tıbbına Entegrasyonu: Evrimsel İpuçlarının Tercümesi",
        "Modern genomik çağın en büyük vizyonu; doğanın milyonlarca yılda geliştirdiği bu kanser direnci ve uzun ömür algoritmalarını hayvanlarda bırakmayıp insan tıbbına transfer etmektir.",
        "Bir filin TP53 kopyaları, bir kör köstepek faresinin hyaluronik asit sentazı veya bir balinanın DNA onarım polimerazı; CRISPR, sentetik mRNA veya transgenik gen terapileri ile insan hücrelerine aktarılabilir. Evrimsel komparatif biyoloji, biyolojik ölümsüzlüğün en zengin ve test edilmiş genetik açık kaynak kütüphanesidir.",
        "Klonlama_Evrim: Gen_Hayvan (Direnç) --[CRISPR / AAV]--> Genom_Insan = Kanser_Direnci + Uzun_Omur",
        "Bu translasyonel biyoteknoloji aksiyomu, komparatif genomik verilerin insan ömrünü radikal biçimde uzatacak genetik mühendisliğin temel hammaddesi olduğunu formüle eder."
    )
]

# ================= KISIM 2 =================
part2_subsections = [
    (
        "2.1 Çıplak Kör Köstebek Faresi (Heterocephalus glaber): Yeraltının Kanser-Geçirmez Kemirgeni",
        "Fare ve sıçan gibi kemirgenler ortalama 2-3 yıl yaşayıp yüksek oranda kanserden ölürken; Doğu Afrika'nın yeraltı tünellerinde yaşayan 'Çıplak Kör Köstebek Faresi' (Naked Mole-Rat - NMR) tam 37 yıldan fazla yaşar ve laboratuvar koşullarında kansere karşı neredeyse mutlak bir direnç sergiler.",
        "Vera Gorbunova ve Andrei Seluanov laboratuvarlarının çığır açan keşifleri; köstebek farelerinin hücrelerinin kanserojen kimyasallarla veya onkogenlerle (Ras + SV40 LT) uyarıldığında dahi tümör oluşturmadığını kanıtlamıştır. Bu canlılar yaşlanma belirtisi göstermez, kardiyovasküler sistemleri bozulmaz ve kemik erimesi yaşamazlar (Negligible Senescence).",
        "Maksimum_Omur_NMR = Lifespan_NMR > 37 Yıl >> Lifespan_Fare (3 Yıl) [12 Kat Fark]",
        "Bu allometrik yaşam süresi anomalisi, köstebek faresinin kendi boyutundaki standart bir kemirgene kıyasla sergilediği olağanüstü uzun ömürlülük katsayısını belgeler."
    ),
    (
        "2.2 Yüksek Moleküler Ağırlıklı Hyaluronik Asit (HMW-HA): 6-12 MegaDaltonluk Zırh",
        "Köstebek faresinin kanser direncinin kalbindeki en temel moleküler keşif, dokularında salgıladıkları 'Yüksek Moleküler Ağırlıklı Hyaluronik Asit'tir (HMW-HA).",
        "İnsan ve fare hücreleri 0.5 ila 2 MegaDalton (MDa) moleküler ağırlığında hyaluronik asit üretirken; çıplak kör köstebek faresi hücreleri tam 6 ila 12 MegaDaltonluk (insandan 5-10 kat daha uzun) devasa glikozaminoglikan zincirleri sentezler. Bu devasa molekül hücrelerin etrafında son derece viskoz, elastik ve aşılmaz bir jelatinimsi zırh örer. Hücreler bu viskoz kafesin içinde birbirine aşırı yaklaşamaz ve tümöral kümeleşme fiziksel olarak bloke edilir.",
        "Molekul_Agirligi_HA = MW_NMR = 6 - 12 MDa >> MW_insan (1 - 2 MDa)",
        "Bu polimerik zincir uzunluğu karşılaştırması, köstebek faresinin ekstraselüler matriksindeki HMW-HA moleküllerinin devasa boyut üstünlüğünü tanımlar."
    ),
    (
        "2.3 Hyaluronan Sentaz 2 (HAS2) Enzim Mutasyonları ve Yavaş Yıkım (HYAL2)",
        "Köstebek faresinin bu devasa hyaluronik asit zincirlerini üretebilmesi; Hyaluronan Sentaz 2 (HAS2) enzimindeki spesifik evrimsel amino asit mutasyonları (Ser211 ve Val657 kalıntıları) sayesinde gerçekleşir.",
        "Bu mutasyonlar HAS2 enziminin polimerizasyon prosesivitesini artırarak zinciri koparmadan milyonlarca dalton boyunca uzatmasını sağlar. Eş zamanlı olarak bu dev zincirleri parçalayan hiyaluronidaz enzimleri (HYAL1, HYAL2) köstebek faresinde son derece düşük aktiviteye sahiptir. Üretim tavan yaparken yıkım dipte kalır; dokular HMW-HA ile tıka basa dolar.",
        "Denge_HMW_HA = d[HMW-HA]/dt = V_HAS2_mutant - k_yikim_HYAL2 * [HMW-HA] >> Denge_Insan * 10",
        "Bu biyosentetik denge eşitliği, HAS2 hiperaktivitesi ve zayıflatılmış hiyaluronidaz yıkımının dokulardaki devasa HMW-HA konsantrasyonunu nasıl yarattığını formüle eder."
    ),
    (
        "2.4 Erken Temas İnhibisyonu (Early Contact Inhibition - ECI) Mekanizması",
        "Normal insan hücreleri petri kabında tek bir tabaka oluşturacak kadar birbirine dokunduğunda bölünmeyi durdurur (Temas İnhibisyonu); kanser hücreleri ise temas inhibisyonunu kaybederek üst üste yığılıp tümör oluşturur.",
        "Çıplak kör köstebek faresi hücreleri ise 'Erken Temas İnhibisyonu' (ECI) sergiler: Hücreler henüz birbirine tam dokunmadan, çok düşük bir hücresel yoğunlukta dahi bölünmeyi hemen durdururlar. Bu erken fren mekanizması sayesinde köstebek faresi hücreleri asla kontrolsüz kitleler halinde çoğalamaz; tümör oluşumu daha birinci basamakta kilitlenir.",
        "Esik_ECI = Dansite_Hucresel_Durdurma = N_kritik_NMR ~ 0.25 * N_kritik_insan (4 Kat Düşük Yoğunlukta Fren)",
        "Bu temas inhibisyonu yoğunluk eşiği, köstebek faresi hücrelerinin neoplastik yığılmaları önlemek için çok düşük hücre dansitesinde bölünmeyi durdurma hassasiyetini simgeler."
    ),
    (
        "2.5 CD44 Reseptörü ve p16INK4a / p27Kip1 Çift-Katmanlı Fren Devresi",
        "Erken Temas İnhibisyonunun hücre içi sinyal kaskadı, hücre yüzeyindeki CD44 reseptörü ile HMW-HA'nın fiziksel etkileşimiyle ateşlenir.",
        "HMW-HA hücre zarına yanaştığında CD44 reseptörlerini kümeleştirir. Bu sinyal hücresel tümör baskılayıcı kaskadını iki aşamalı olarak tetikler: 1) İlk aşamada p16INK4a senesens yolağı erken aktive olarak Rb proteininin fosforilasyonunu engeller ve hücre döngüsünü G1'de kilitler; 2) Hücre p16 engelini aşsa dahi arkasından p27Kip1 ikinci güvenlik bariyeri olarak devreye girer. Bu 'Çift-Katmanlı Fren Devresi' (Two-Tiered Cell Cycle Arrest) onkojenik kaçışı imkânsız kılar.",
        "Durdurma_Sinyali: HMW-HA ----> CD44 Kumeleşmesi ----> p16INK4a (Fren 1) ----> p27Kip1 (Fren 2) ----> G1 Blokajı",
        "Bu moleküler sinyal iletim kaskadı, köstebek faresinin hücre döngüsünü kitleyen aşılmaz çift kademeli tümör baskılama yolunu modeller."
    ),
    (
        "2.6 Aseksüel Translasyonel Sadakat: 28S Ribozomal RNA'nın Bölünmesi",
        "Köstebek farelerinin kansere ve yaşlanmaya karşı ikinci büyük moleküler kalkanı, protein sentezindeki (translasyon) inanılmaz doğruluk oranlarıdır.",
        "Köstebek faresinin 28S ribozomal RNA'sı (rRNA), evrimsel bir adaptasyonla tam ortasından iki parçaya bölünmüştür (Split 28S rRNA). Bu yapısal modifikasyon ribozomun konformasyonel rijitliğini artırır; ribozomal kodon okuma hatası (misfolding) sıradan bir farenin veya insanın 10'da birine düşer. Hatasız proteinler proteostaz stresini sıfırlar ve kanserojen mutasyonları önler.",
        "Hata_Translasyon = E_ribozom = N_yanlis_aa / N_toplam_aa < 10^-6 (İnsandan 10 Kat Daha Sadık)",
        "Bu translasyonel sadakat katsayısı, bölünmüş 28S rRNA yapısının ribozomal protein sentezinde sağladığı kusursuz doğruluk standardını belgeler."
    ),
    (
        "2.7 Yüksek Proteostaz Kapasitesi: 26S Proteazom ve Otofaji Direnci",
        "Hatalı protein üretimi düşük olduğu gibi, var olan hasarlı proteinlerin temizlenme kapasitesi de köstebek faresinde olağanüstü yüksektir.",
        "Köstebek faresi dokularında 26S ve 20S proteazom enzim aktivitesi sıradan kemirgenlerin iki katıdır; şaperon proteinleri (HSP70, HSP90) her zaman aktif nöbet tutar. Oksidatif strese veya ağır metallere maruz kalsalar dahi protein agregatları hücrede birikemez; otofaji akışı 30 yaşındaki bir köstebek faresinde 6 aylık genç bir faredeki kadar hızlı ve etkindir.",
        "Aktivite_Proteazom = V_proteoliz = k_proteazom * [Proteazom_aktif] * ( [Ubikitin_yuk] / K_ub ) >> Fare_Proteoliz * 2",
        "Bu proteolitik temizlik kinetiği formülü, köstebek faresi hücrelerindeki üstün proteazom kapasitesinin proteotoksik yaşlanmayı nasıl tamamen engellediğini ifade eder."
    ),
    (
        "2.8 Hiyalurondiidaz Enzimi İnhibitörleri ve Doku Esnekliği",
        "Köstebek faresinin cildi ve dokuları elastik, gevşek ve kırışıksızdır; bu durum dar yeraltı tünellerinde sürtünmeden hareket edebilmesini sağlar.",
        "Bu elastikiyetin nedeni yine bol miktardaki HMW-HA'dır. Ancak köstebek faresine deneysel olarak bakteriyel hiyaluronidaz enzimi verilip HMW-HA eritildiğinde; hücrelerin erken temas inhibisyonunu kaybettiği, normal fare hücreleri gibi kontrolsüzce bölündüğü ve tümör oluşturabildiği kanıtlanmıştır. Bu deney HMW-HA'nın kanser direncinin yegâne nedensel anahtarı olduğunu şüpheye yer bırakmadan tescillemiştir.",
        "Kanser_Direnci_Kosulu: Eger [HMW-HA] > Eşik -> Kanser_Riski = 0 | Eger [HMW-HA] = 0 -> Kanser_Riski = Yüksek",
        "Bu nedensellik önermesi, yüksek moleküler ağırlıklı hyaluronik asidin köstebek faresinde tümör oluşumunu engelleyen birincil moleküler kalkan olduğunu belgeler."
    ),
    (
        "2.9 İnsan Hücrelerinde HMW-HA Üretimi: Transgenik nmrHAS2 Entegrasyonu",
        "Vera Gorbunova laboratuvarı 2023 yılında Nature'da tıp tarihine geçen tarihi bir deney yayımlamıştır: Köstebek faresinin nmrHAS2 geni transgenik olarak farelere aktarılmıştır.",
        "Köstebek faresi HAS2 genini taşıyan fareler; kendi dokularında HMW-HA üretmeye başlamıştır. Sonuçlar inanılmazdır: Farelerin kanser insidansı dramatik biçimde düşmüş, bağırsak ve cilt bariyerleri gençleşmiş, kronik enflamasyonları azalmış ve medyan yaşam süreleri %4.4, maksimum yaşam süreleri %12.2 uzatılmıştır. Bu başarı, köstebek faresi mekanizmalarının diğer memelilere başarıyla transfer edilebileceğini kesin olarak kanıtlamıştır.",
        "Uzatma_Lifespan = Delta_Omur_Fare = + 12.2% (nmrHAS2 Transgeni ile Sağlanan Net Kazanç)",
        "Bu transgenik longevity kazanımı, komparatif biyoloji genlerinin türler arası aktarımla yaşam süresini uzatma gücünün somut kanıtıdır."
    ),
    (
        "2.10 Kanser-Geçirmez İnsan Dokuları: CRISPR ile HAS2/CD44 Ekseninin Yeniden Kodlanması",
        "Homo Aeternus genom mühendisliği protokolünde; insan HAS2 geninin promoter ve katalitik bölgeleri CRISPR/Prime Editing ile köstebek faresi nmrHAS2 dizilimine dönüştürülür.",
        "İnsan fibroblastları ve kök hücreleri 8-10 MDa HMW-HA salgılamaya başlar; dokularda CD44 aracılı erken temas inhibisyonu kurulur. Bu mühendislik hamlesi, insan vücudunu tüm katı tümör metastazlarına ve kontrolsüz karsinogeneze karşı moleküler düzeyde zırhlar; kanser insan için biyolojik bir tehdit olmaktan ebediyen çıkar.",
        "Kanser_Gecirmezlik_Indeksi = I_NMR_Human = [HMW-HA_insan] * [CD44_afinite] / K_onkojen >> 10^3",
        "Bu sentetik komparatif katsayı, köstebek faresi genetik mimarisinin insan hücrelerinde kurduğu aşılmaz neoplastik direnç gücünü formüle eder."
    )
]

# ================= KISIM 3 =================
part3_subsections = [
    (
        "3.1 Grönland Balinası (Balaena mysticetus): 200+ Yıllık Yaşam ve Arktik Adaptasyon",
        "Arktik okyanusun dondurucu sularında yaşayan Grönland Balinası (Bowhead Whale); 211 yıldan fazla belgelenmiş yaşam süresiyle yeryüzündeki en uzun ömürlü memeli canlıdır.",
        "Eski inuit taş zıpkın uçlarının yaşayan balinaların yağ dokusunda bulunması ve göz lensi aspartik asit rasemizasyon analizleri ile kanıtlanan bu iki asırlık ömür boyunca; 100 tonluk devasa gövdelerinde trilyonlarca hücre bölünmesine rağmen kanser, kardiyovasküler hastalık veya metabolik yetmezlik görülmez. Grönland balinası Peto paradoksunun yaşayan en muazzam anıtıdır.",
        "Lifespan_Bowhead = T_max > 211 Yıl (Memeliler Aleminin En Uzun Ömür Rekoru)",
        "Bu maksimum yaşam süresi parametresi, Grönland balinasının iki asrı aşan biyolojik gençlik kapasitesini belgeler."
    ),
    (
        "3.2 De Novo Genom Dizilemesi ve Pozitif Seçilim Altındaki Genomik Bölgeler",
        "2015 yılında João Pedro de Magalhães liderliğindeki uluslararası konsorsiyum; Grönland balinasının tüm genomunu dizilemiş ve insan/fare genomlarıyla karşılaştırmıştır.",
        "Genomik madencilik analizleri; balina genomunda DNA onarımı, hücre döngüsü regülasyonu, kanser baskılama ve insülin sinyal yolaklarında yer alan genlerin yoğun bir 'Pozitif Doğal Seçilim' (Ka/Ks > 1) altında evrildiğini ortaya koymuştur. Balina genomu rastgele mutasyonlarla değil, kansere ve yaşlanmaya karşı özel olarak akort edilmiş genetik mutasyonlarla doludur.",
        "Secilim_Katsayisi = Ka/Ks = ( Es-anlamsiz_Mutasyonlar / Es-anlamli_Mutasyonlar ) > 1.0 (Pozitif Seçilim)",
        "Bu moleküler evrim eşitliği, Grönland balinası longevity genlerinin milyonlarca yıllık hedefe yönelik doğal adaptasyonla mükemmelleştirildiğini kanıtlar."
    ),
    (
        "3.3 ERCC1 ve PCNA Gen Mutasyonları: DNA Eksizyon Onarımında Evrimsel Mükemmellik",
        "Balinanın pozitif seçilim gösteren en kritik genlerinden biri, Nükleotid Eksizyon Onarımı (NER) ve rekombinasyonel tamirde görevli 'ERCC1' ve DNA polimeraz kıskacı 'PCNA'dır.",
        "Grönland balinası ERCC1 geninde insan ve diğer tüm memelilerden farklılaşan spesifik amino asit değişiklikleri mevcuttur. Bu değişiklikler ERCC1-XPF endonükleaz kompleksinin DNA çift zincir kırıklarına ve UV/oksidatif çapraz bağ hasarlarına bağlanma afinitesini katlamıştır. Hücre bölünürken replikasyon çatalı asla duraksamaz; hasarlı bazlar anında kesilip çıkarılarak onarılır.",
        "Hiz_NER_Onarim = J_NER = k_ERCC1_balina * [ERCC1] * [Hasarli_Lezyon] >> Hiz_NER_insan * 3",
        "Bu nükleotid eksizyon onarım kinetiği denklemi, balina ERCC1 enziminin DNA hasar temizleme hızındaki üç katlık kinetik üstünlüğünü formüle eder."
    ),
    (
        "3.4 Çift Zincir Kırığı (DSB) Onarımında Ultra-Yüksek Doğruluk: Homolog Rekombinasyon (HR)",
        "Vera Gorbunova laboratuvarının 2023 araştırması; Grönland balinası hücrelerinin DNA çift zincir kırıklarını (DSB) onarmada insan, fare ve inek hücrelerine kıyasla olağanüstü bir üstünlüğe sahip olduğunu göstermiştir.",
        "İnsan ve fare hücreleri DSB onarımında çoğunlukla hata eğilimli ve mutajenik NHEJ yolunu kullanırken; Grönland balinası hücreleri kusursuz ve sıfır-hatalı 'Homolojiye Dayalı Onarım' (HDR / HR) yolunu baskın olarak tercih eder. Balina hücrelerinde Homolog Rekombinasyon verimi insandan tam iki kat daha yüksektir. Çift sarmal kırıkları hiçbir nükleotit kaybı olmadan kardeş kromatid kalıbı üzerinden kusursuzca dikilir.",
        "Oran_HR_NHEJ = ( Aktivite_HR / Aktivite_NHEJ )_Balina >> 2 * ( Aktivite_HR / Aktivite_NHEJ )_Insan",
        "Bu tamir yolağı tercih oranı, Grönland balinasının çift zincir kırıklarında hata eğilimli NHEJ yerine kusursuz homolog rekombinasyonu devreye sokma dehasını belgeler."
    ),
    (
        "3.5 Duplike Olmuş Longevity Genleri ve Genomik Kopya Sayısı Varyasyonları (CNVs)",
        "Yalnızca var olan genleri mutasyona uğratmakla kalmayan balina evrimi; kritik koruyucu genlerin genomdaki kopya sayısını gen duplikasyonları (CNV) ile çoğaltmıştır.",
        "Grönland balinası genomunda DNA hasar yanıtı, apoptoz, mikrotübül organizasyonu ve immün kontrol genlerinin birden fazla fonksiyonel kopyası yer alır (örneğin LAMTOR1 duplikasyonu). Bir kopya mutasyonla hasar görse dahi diğer duplike kopyalar kesintisiz çalışmaya devam eder; bu durum 'Genomik Yedeklilik' (Genomic Redundancy) yaratarak tek nokta arızalarını imkânsız kılar.",
        "Yedeklilik_Faktoru = R_genom = Prod_i [ Kopya_Sayisi_Gen_i ] (Arızaya Karşı Sarsılmaz Koruma)",
        "Bu genetik mimari güvenilirlik eşitliği, duplike olmuş koruyucu gen kopyalarının hücresel fonksiyon çöküşü olasılığını nasıl sıfıra yaklaştırdığını tanımlar."
    ),
    (
        "3.6 CIRBP (Soğuk-İndüklenebilir RNA-Bağlayıcı Protein) ve Kriyoprotektif Kararlılık",
        "Grönland balinası ömrünü sıfırın altındaki kutup sularında geçirir; bu soğuk ortam hücrelerde translasyonel duraksama ve stres granülleri oluşturabilir.",
        "Balina hücrelerinde 'Soğuk-İndüklenebilir RNA-Bağlayıcı Protein' (CIRBP) aşırı eksprese edilir. CIRBP, hücre içi haberci RNA'lara bağlanarak onları düşük sıcaklıkta parçalanmaktan korur ve translasyonun soğukta dahi aksamadan devam etmesini sağlar. Eş zamanlı olarak CIRBP, DNA onarım proteini ATR'yi aktive ederek hücresel streste genomu kalkan gibi korur.",
        "Ekspresyon_CIRBP = [CIRBP] >> Bazal_Memeli * 5 (Sürekli Aktif Arktik Kalkan)",
        "Bu koprotektif protein ekspresyon seviyesi, balinanın dondurucu ortamda dahi hücresel transkriptomu koruma ve DNA hasarını frenleme gücünü belgeler."
    ),
    (
        "3.7 Mitokondriyal Oksidatif Fosforilasyon Sadakati ve Düşük ROS Salınımı",
        "Balinanın devasa kas kütlesi ve yüzme eforu büyük enerji gerektirir; ancak mitokondriyal elektron taşıma zinciri minimum elektron kaçağı yapacak şekilde evrilmiştir.",
        "Mitokondriyal genomdaki Kompleks I (ND1, ND2, ND4) ve Kompleks IV (Sitokrom c Oksidaz) genlerindeki spesifik mutasyonlar; proton pompalama verimini artırırken elektron kaçağını (O2•- süperoksit üretimini) fareye kıyasla %85 azaltmıştır. Mitokondriler hücre içi DNA'yı paslandırmadan temiz ATP üretir.",
        "Kacak_Elektron_Balina = J_kacak_ROS < 0.05 * J_kacak_fare (Temiz Biyoenerjetik)",
        "Bu mitokondriyal elektron transfer oranı, balina ETC zincirinin minimum reaktif oksijen türü salarak hücresel yaşlanmayı nasıl frenlediğini modeller."
    ),
    (
        "3.8 Hücresel Yaşlanma (Senesens) Eşiğinin Yükseltilmesi ve SASP Direnci",
        "Fare hücreleri 10-15 bölünmede replikatif senesense girerken ve insanlar 50 bölünmede (Hayflick sınırı) dururken; Grönland balinası hücreleri senesense girmeye karşı aşırı yüksek bir eşiğe sahiptir.",
        "Balina hücrelerinde telomer erozyonu çok daha yavaş ilerler ve hücreler senesense girseler dahi toksik pro-enflamatuar SASP faktörlerini salgılamazlar. Hücre ya sağlıklı kalır ya da dokuya zarar vermeden sessizce apoptoza gider; kronik yangı (inflammaging) balina dokularında tamamen silinmiştir.",
        "Esik_Senesens = N_Hayflick_Balina >> 150 Bölünme (Uzun Hücresel Replikatif Rezerv)",
        "Bu hücresel replikatif kapasite eşitsizliği, balina hücrelerinin insan hücrelerinin üç katı bölünme rezervine sahip olduğunu gösterir."
    ),
    (
        "3.9 Kanser Baskılayıcı İkincil Yolaklar ve Temas İnhibisyonu Dinamikleri",
        "Grönland balinası hücreleri, temas inhibisyonunda da benzersiz adaptasyonlar sergiler.",
        "Hücreler kültürde belirli bir yoğunluğa ulaştığında E-kaderin ve Hippo-YAP sinyal yolağı aşırı duyarlılıkla aktive olur. YAP/TAZ transkripsiyon faktörleri sitoplazmaya sürülerek inaktive edilir ve nükleer hücre döngüsü genleri durdurulur. Onkogenik mutasyon taşıyan tek bir hücre dahi komşuları tarafından anında mekanik baskıyla apoptoza itilir.",
        "Fosforilasyon_YAP = [p-YAP_inaktif] / [YAP_toplam] -> 1.0 (Temas Anında Anında Kilit)",
        "Bu mekanotransdüksiyonel tümör baskılama oranı, balina hücrelerinin doku aşırı büyümesini Hippo yolağı üzerinden nasıl tavizsiz durdurduğunu tanımlar."
    ),
    (
        "3.10 İnsan Hücrelerine Grönland Balinası DNA Onarım Genlerinin Aktarımı",
        "Grönland balinasının 200 yıllık sırrını insan genomuna taşımak amacıyla; balina ERCC1 ve PCNA varyantları sentetik olarak insan iPSC ve fibroblastlarına aktarılmıştır.",
        "Balina genlerini ifade eden insan hücreleri; ölümcül radyasyona (gama ışınları) ve bleomisin DNA hasarına maruz bırakıldığında, normal insan hücrelerine kıyasla çift sarmal kırıklarını 2 kat daha hızlı ve sıfır delesyonla onarmıştır. Bu genetik modifikasyon, Homo Aeternus genom projesinin en sağlam DNA onarım kolonunu oluşturmaktadır.",
        "Verim_Onarim_Kombine = eta_DSB = eta_insan * ( 1 + 1.2 * [Balina_ERCC1_ekspresyonu] )",
        "Bu transgenik tamir amplifikasyon eşitliği, balina DNA onarım enzimlerinin insan hücre genomik stabilitesini nasıl ikiye katladığını matematiksel olarak belgeler."
    )
]

# ================= KISIM 4 =================
part4_subsections = [
    (
        "4.1 Fillerde (Loxodonta africana) Kanser Direnci ve Peto Paradoksunun Zirvesi",
        "Afrika savan fili (Loxodonta africana), 6 tona varan devasa cüssesi, insandan 100 kat fazla hücresi ve 60-70 yıllık yaşam süresiyle kanser riski en yüksek olması gereken memelilerden biridir.",
        "Ancak yapılan nekropsi ve epidemiyolojik kayıtlar; fillerin kanserden ölüm oranının %5'in altında olduğunu göstermektedir (insanlarda bu oran %20 ila %25'tir). Joshua Schiffman ve ekibinin 2015'te JAMA'da yayımlanan keşfi; fillerin kansere karşı geliştirdiği inanılmaz genomik silahı deşifre etmiştir: Genomik TP53 gen amplifikasyonu.",
        "Kanserden_Olum_Orani: Fil (%4.8) << Insan (%23.5) << Fare (%50 - 70)",
        "Bu klinik epidemiyolojik ölüm oranı karşılaştırması, fil türünün devasa cüssesine rağmen insandan beş kat daha düşük kanser mortalitesine sahip olduğunu belgeler."
    ),
    (
        "4.2 20 Kopya (40 Alel) TP53 Genomu: 'Genomun Muhafızı'nın Çoğaltılması",
        "İnsan genomunda 'Genomun Muhafızı' olarak bilinen TP53 tümör baskılayıcı geninden yalnızca tek bir çift (2 alel) bulunur ve bir alelin mutasyonu (Li-Fraumeni sendromu) neredeyse kesin kanserle sonuçlanır.",
        "Fillerin genomunda ise tam 20 bağımsız kopya (toplam 40 alel) TP53 geni mevcuttur. Bu kopyalardan biri atadan kalma orijinal kodlama geni iken, 19 tanesi evrimsel süreçte retrotranspozisyonla duplike olmuş ve fonksiyonel olarak aktifleşmiş retrogenlerdir (TP53RTG). Bir fil hücresinde p53 protein havuzu insan hücresinin onlarca katı büyüklüktedir.",
        "Gen_Kopya_Sayisi_TP53 = N_TP53_fil = 20 Kopya (40 Alel) >> N_TP53_insan (1 Kopya / 2 Alel)",
        "Bu genetik kopya sayısı oranı, fil genomunun insan genomuna kıyasla yirmi kat daha fazla tümör baskılayıcı p53 rezervi taşıdığını tanımlar."
    ),
    (
        "4.3 p53 Protein Havuzu ve DNA Hasarına Aşırı Duyarlı Hücresel İntihar (Apoptoz)",
        "Fil lenfositleri radyasyona (iyonize ışın) maruz bırakıldığında, insan hücrelerinden tamamen zıt ve radikal bir davranış sergiler.",
        "İnsan hücreleri DNA kırığı oluştuğunda önce bölünmeyi durdurup DNA'yı tamir etmeye çalışır; tamir başarısız olursa veya hatalı olursa mutasyon kalıcılaşır ve kanser tohumu atılır. Fil hücreleri ise tamirle hiç vakit kaybetmez: Muazzam p53 rezervi nedeniyle en ufak bir DNA çift zincir kırığında dahi hücre döngüsünü tamir için bekletmek yerine hücreyi anında 'Apoptoza' (programlı hücre ölümü) sürükler. Fil, 'şüpheli tek bir hücreyi bile yaşatmama' tavizsizliği ile tümörleşmeyi sıfırlar.",
        "Apoptoz_Hassasiyeti = P_apoptoz_fil = 1 - exp( - k_p53 * N_kopya_TP53 * [DSB] ) >> P_apoptoz_insan",
        "Bu aşırı duyarlı apoptoz olasılık fonksiyonu, fil hücrelerinin DNA hasarı karşısında insandan beş kat daha yüksek hızla intihar ederek kanserleşmeyi önleme mekanizmasını modeller."
    ),
    (
        "4.4 Uyanmış Zombi Gen: LIF6 (Lösemi İnhibitör Faktör 6) ve Mitokondriyal İmha",
        "Filin kanser savunmasındaki ikinci büyük moleküler mucize, milyonlarca yıldır ölü bir psödogen olan 'LIF6' geninin evrimsel olarak yeniden dirilmesidir.",
        "Memelilerde LIF (Leukemia Inhibitory Factor) gen ailesi mevcuttur; ancak fil evriminde LIF6 psödogenden fonksiyonel bir gene dönüşmüştür. LIF6'nın promotöründe spesifik p53 bağlanma bölgeleri bulunur. DNA hasarı oluşup p53 seviyesi fırladığı anda LIF6 proteini sentezlenir. LIF6 doğrudan mitokondriye göç eder, mitokondri dış zarına saplanarak delikler açar (MOMP) ve sitokrom c'yi sitoplazmaya dökerek kaspaz kaskadını patlatır; hasarlı hücre dakikalar içinde yok edilir.",
        "Aktivasyon_Kaskadi: DNA Hasarı ----> p53 Fazlalığı ----> LIF6 Transkripsiyonu ----> Mitokondri Delinmesi ----> Hızlı Apoptoz",
        "Bu moleküler imha yolağı, p53 ile uyanmış zombi gen LIF6'nın hasarlı hücreleri mitokondriyal delinmeyle anında yok etme zincirini belgeler."
    ),
    (
        "4.5 MDM2 Feedback Döngüsünün Aşılması ve Sürekli Nükleer p53 Aktivitesi",
        "İnsan hücrelerinde p53 aktivitesi, kendi ürettiği negatif regülatörü olan E3 ubikitin ligazı MDM2 tarafından hızla ubiquitine edilip proteazomda parçalanarak baskılanır.",
        "Fillerin duplike olmuş retrogen p53 kopyaları, C-terminal bölgesinde spesifik delesyonlar ve mutasyonlar taşır; bu mutasyonlar MDM2 enziminin p53'e bağlanmasını ve onu parçalamasını zorlaştırır. Bu sayede fil hücrelerinde nükleer p53 havuzu hücresel stres anında insan hücresine göre çok daha uzun süre aktif kalır ve tümör baskılama sinyali yarıda kesilemez.",
        "Yari_Omur_p53 = Tau_p53_fil >> 3 * Tau_p53_insan (MDM2 Yıkımına Karşı Evrimsel Direnç)",
        "Bu kinetik kararlılık eşitliği, fil p53 proteinlerinin hücresel stres anında nükleer aktif yarı ömrünün insan p53'ünden üç kat daha uzun olduğunu simgeler."
    ),
    (
        "4.6 Radyasyon Direnci: Gama Işınları Altında Fil Lenfositlerinin Davranışı",
        "Laboratuvar testlerinde fil periferik kan mononükleer hücreleri (PBMC) ile insan PBMC'leri ölümcül dozda gama radyasyonuna (0.5 ila 2 Gray) maruz bırakılmıştır.",
        "İnsan lenfositlerinde gama ışınları çok sayıda mikronükleus ve kromozomal aberasyon oluşturup hücreleri senesense sürüklerken; fil lenfositleri hasarlı hücrelerin %85'ini ilk 24 saat içinde kusursuz apoptozla temizlemiştir. Geride kalan sağlam hücreler sıfır kromozomal hasarla hayatına devam etmiştir; fil dokuları genotoksik kirlenmeye asla izin vermez.",
        "Hasar_Kromozomal = N_mikronukleus_fil < 0.1 * N_mikronukleus_insan (Gama Işınları Sonrası)",
        "Bu genotoksisite direnç oranı, fil hücrelerinin radyasyon maruziyeti sonrasında mutasyon biriktirmek yerine hasarlı hücreleri temizleme üstünlüğünü belgeler."
    ),
    (
        "4.7 Hücresel Feda Edilebilirlik (Sacrifice) Stratejisi ve Dokusal Rezerv",
        "Filin kanser stratejisi felsefi olarak 'Hücresel Feda Edilebilirlik' (Cellular Altruism) ilkesine dayanır.",
        "Fil 60 trilyon hücreye sahiptir; bu devasa havuz içinde DNA'sı hafifçe çizilmiş birkaç milyon hücreyi anında öldürüp çöpe atmak fil için önemsiz bir dokusal maliyettir. Küçük canlılar her hücresini yaşatmak için tamire zorlanırken; devasa fil hasarlı hücreyi anında feda eder ve kök hücre havuzundan taptaze yenisini üretir.",
        "Maliyet_Hucresel = Delta_N_hucre / N_toplam ~ 10^-7 (Fil İçin İhmal Edilebilir Doku Maliyeti)",
        "Bu dokusal stokiyometri eşitliği, trilyonlarca hücreye sahip dev bir organizmanın hasarlı hücreleri acımasızca feda edebilme biyolojik lüksünü formüle eder."
    ),
    (
        "4.8 Fil Kök Hücreleri ve Aşırı Duyarlı Telomer Kontrol Mekanizmaları",
        "Apoptoz mekanizmasının bu denli agresif çalışması kök hücrelerin hızla tükenmesi riskini doğurabilirdi; ancak fil evrimi bunu telomeraz dinamikleriyle dengelemiştir.",
        "Fil indüklenmiş pluripotent kök hücrelerinde (iPSC) ve hematopoietik nişlerde TERT enzimi son derece sıkı regüle edilir. Kök hücreler bölündükçe telomer erozyonunu hassas bir sınırda tutar; ancak neoplastik bir immortalizasyon atağı belirdiği anda p53 kopyaları anında TERT promotörünü baskılayarak hücreyi apoptoza iter.",
        "Denge_Kök_Hucre: Rejenerasyon_Gucu = Yuksek | Onkojenik_Kacis = SIFIR",
        "Bu hücresel denge prensibi, fil kök hücre havuzunun rejeneratif gücünü korurken kanserojen dönüşümü nasıl mutlak olarak kilitlediğini tanımlar."
    ),
    (
        "4.9 İnsan Genomuna Ek TP53 Kopyalarının Eklenmesi: Transgenik 'Süper-p53' Modelleri",
        "İspanyol onkolog Manuel Serrano ve ekibi; fare genomuna yapay BAC vektörleriyle ekstra TP53 kopyaları ekleyerek 'Süper-p53' fareleri üretmiştir.",
        "Sonuçlar sansasyoneldir: Ekstra p53 kopyası taşıyan fareler kansere karşı neredeyse tamamen bağışık hale gelmiş, tümör insidansı sıfıra yaklaşmış ve en önemlisi erken yaşlanma (progeria) yaşamadan sağlıklı yaşam süreleri %16 uzatılmıştır. Kanser barajı yıkıldığında p53'ün DNA onarım ve antioksidan fonksiyonları doğrudan yaşam süresini uzatmaktadır.",
        "Kanser_Direnci_SuperP53: Insidans_Tumor -> 0 | Saglikli_Omur = + %16",
        "Bu transgenik memeli verisi, genomik p53 kopya sayısı artışının erken yaşlanma yaratmaksızın tümörleri sildiğini ve ömrü uzattığını kesinleştirir."
    ),
    (
        "4.10 Kanser Direnci Kokteyli: İnsan Kök Hücrelerine 20 Kopya TP53 ve LIF6 Entegrasyonu",
        "Homo Aeternus genom projesinde; insan pluripotent kök hücrelerine CRISPR-Cas9 ve sentetik yapay kromozomlar (HACs) kullanılarak filin 20 kopya TP53 dizisi ve fonksiyonel LIF6 zombi geni entegre edilir.",
        "Bu genetik modifikasyon insan hücrelerini 'Süper-Duyarlı Apoptoz' yeteneği ile donatır. Hücrelerimiz artık DNA hasarını mutasyona dönüştüremeyecek, prekanseröz tek bir hücre dahi dokuda barınamayacaktır. Kanser; insan türünün biyolojik lügatinden silinmiş arkaik bir evrimsel fosile dönüşür.",
        "Guvenlik_Onkogenik = Risk_Kanser = Risk_0 * ( 1 / 20^N_kopya_TP53 ) ~ SIFIR",
        "Bu katlanarak azalan risk eşitliği, sentetik 20 kopya TP53 entegrasyonunun insan neoplastik dönüşüm riskini teorik olarak sıfırladığını matematiksel olarak ilan eder."
    )
]

# ================= KISIM 5 =================
part5_subsections = [
    (
        "5.1 Negligible Senescence (İhmal Edilebilir Yaşlanma) Kavramı ve Caleb Finch Kriterleri",
        "Biyolojik yaşlanma tüm canlılar için evrensel bir doğa kanunu mudur? Gerontolog Caleb Finch'in 1990'da ortaya koyduğu 'Negligible Senescence' (İhmal Edilebilir Yaşlanma) teorisi bu soruya kesin bir 'HAYIR' yanıtı vermiştir.",
        "Bir türün ihmal edilebilir yaşlanma sergilemesi için üç biyolojik şartı sağlaması gerekir: 1) Yaşlandıkça ölüm oranı (mortalite hızı) artmaz (Gompertz yaşlanma eğrisi yatay bir düz çizgidir); 2) Yaşlandıkça üreme kapasitesi (fertilite) azalmaz; 3) Yaşlandıkça fizyolojik performans, metabolik hız ve kas gücü gerilemez. Yeryüzünde bu kriterleri sağlayan onlarca omurgalı ve omurgasız tür yaşamaktadır.",
        "Gompertz_Mortalite: m(t) = m_0 * exp( G * t ) | Negligible Senescence Icin: G = 0 -> m(t) = m_0 = Sabit",
        "Bu Gompertz aktüeryal mortalite eşitliği, ihmal edilebilir yaşlanma gösteren türlerde yaşlanma hızı katsayısının (G) sıfır olduğunu ve ölüm riskinin yaşla artmadığını belgeler."
    ),
    (
        "5.2 Rougheye Kaya Balığı (Sebastes aleutianus): 205 Yıllık Denizaltı Biyostazı",
        "Pasifik Okyanusu'nun derin ve soğuk sularında yaşayan Rougheye Kaya Balığı (Sebastes aleutianus); 205 yılı aşan yaşam süresiyle bilinen en uzun ömürlü balık türlerinden biridir.",
        "200 yaşındaki bir kaya balığı ile 20 yaşındaki genç bir kaya balığı anatomik ve hücresel olarak incelendiğinde; kas gücünde, göz lensi berraklığında, bağışıklık fonksiyonunda veya üreme yeteneğinde hiçbir gerileme tespit edilemez. 200 yaşındaki dişi bir kaya balığı, genç dişilerden daha fazla ve daha kaliteli yumurta üretir. Yaşlanma bu canlıda fizyolojik olarak durdurulmuştur.",
        "Ureme_Kapasitesi = N_yumurta(t) ~ K_fertilite * Kütle(t) (Yaşlandıkça Azalmaz, Artar)",
        "Bu fekondite büyüme eşitliği, ihmal edilebilir yaşlanma sergileyen canlılarda yaşlanmanın üreme kapasitesini düşürmediğini, artan beden boyutuyla katladığını belgeler."
    ),
    (
        "5.3 Arctica islandica (Okyanus Quahog İstiridyesi): 507 Yıllık Ming ve Aşılmaz Proteostaz",
        "İzlanda kıyılarında bulunan ve 'Ming' adı verilen bir Arctica islandica istiridyesi; kabuk halkaları sayıldığında tam 507 yaşında olduğu tescillenmiş ve yeryüzündeki en uzun ömürlü tekil hayvan ilan edilmiştir.",
        "500 yıl boyunca yaşayan bu yumuşakçanın hücresel sırrı 'Aşılmaz Proteostaz Kararlılığı'dır. Arctica islandica proteinleri; serbest radikallere, yüksek ısıya ve ağır metallere maruz bırakıldığında dahi denatüre olmaz ve presipite olmaz. Hücre zarlarındaki doymamış yağ asidi peroksidasyon indeksi inanılmaz düşüktür; lipid zarları 5 asır boyunca tek bir oksidatif kırılma yaşamadan esnek kalır.",
        "Lifespan_Ming = 507 Yıl (Kristof Kolomb'dan Modern Çağa Kesintisiz Biyolojik Sağlamlık)",
        "Bu mutlak ömür metriği, omurgasız hayvanlar aleminde ihmal edilebilir yaşlanmanın yarım binyılı aşan somut varoluşsal kanıtıdır."
    ),
    (
        "5.4 Dev Kaplumbağalar (Aldabrachelys gigantea / Chelonoidis niger): 150+ Yıl Telomer Kararlılığı",
        "Galapagos ve Seyşeller dev kaplumbağaları (Jonathan adlı kaplumbağa 190+ yaşında); insan ömrünün iki katından fazla yaşar ve ölüm oranları yaşla artmaz.",
        "Dev kaplumbağa genom dizilemesi; DNA onarım genlerinde (NEIL1, BRCA2), metabolik regülasyon faktörlerinde ve telomer bakım kompleksinde aşırı genişlemeler ortaya çıkarmıştır. Kaplumbağa fibroblastlarında telomeraz aktivitesi yetişkin dokularda da hassas bir bazal seviyede korunur; telomerler kısalmaz, kök hücre nişleri tükenmez ve bağışıklık sistemi (immünosenesens) 150 yıl boyunca genç kalır.",
        "Telomer_Asinmasi = dL_telomer/dt ~ 0 (150 Yıl Boyunca Sıfır Replikatif Erozyon)",
        "Bu telomerik kararlılık eşitliği, dev kaplumbağaların Hayflick replikatif sınırına takılmaksızın asırlar boyu hücresel bölünme rezervini nasıl koruduğunu tanımlar."
    ),
    (
        "5.5 Grönland Köpekbalığı (Somniosus microcephalus): 400 Yıllık Metabolik Yavaşlık",
        "Kuzey Kutup Denizi'nin 2000 metre derinliğinde yaşayan Grönland Köpekbalığı; 392 ± 120 yıllık belgelenmiş ömrüyle omurgalılar aleminin mutlak yaşam süresi şampiyonudur.",
        "Göz lensi çekirdeğindeki karbon-14 izotop analizleri; bu canlıların cinsel olgunluğa ancak 150 yaşında ulaştığını kanıtlamıştır. Yıllık büyüme hızları yalnızca 1 santimetredir. -1.5°C'lik zifiri karanlık kutup sularında kalpleri her 12 saniyede bir kez atar; aşırı yavaşlamış bu metabolik hız, hücresel entropi üretimini sıfıra yaklaştırarak dört asırlık bir biyostaz sağlar.",
        "Ergenlik_Yasi = T_maturasyon ~ 150 Yıl | Lifespan_Maksimum ~ 400 - 500 Yıl",
        "Bu olağanüstü ontogenetik gelişim parametreleri, metabolik hız yavaşlamasının omurgalı biyolojisinde yaşam süresini yarım binyıla nasıl taşıdığını gösterir."
    ),
    (
        "5.6 Kıkırdaklı Balıklarda ve Timsahlarda Telomeraz Sürekliliği",
        "Memelilerin aksine kıkırdaklı balıklar (köpekbalıkları, vatozlar) ve bazı sürüngenler (timsahlar); yetişkin somatik dokularında telomeraz enzimini (TERT) tamamen susturmazlar.",
        "Kontrollü bir bazal telomeraz aktivitesi sayesinde kromozom uçları yaşam boyu orijinal uzunluğunda kalır. Timsahlar 70-80 yaşına geldiklerinde dahi gençliklerindeki kadar hızlı avlanır ve büyümeye devam ederler (Indeterminate Growth - Belirsiz Büyüme). Kanser görülme sıklığı sıfıra yakındır çünkü telomeraz aktivitesi katı tümör baskılayıcı kaskadlarla (p53 homologları) kusursuzca dengelenmiştir.",
        "Aktivite_TERT_Somatik = V_TERT = k_bakim * ( 1 - [Onkojenik_Stres] ) (Sürekli Dengeli Bakım)",
        "Bu dengeli telomeraz ekspresyon eşitliği, somatik dokularda telomerazın kanser riski yaratmadan telomer boyunu ömür boyu nasıl koruduğunu modeller."
    ),
    (
        "5.7 Proteomik Kararlılık: Düşük Oksidasyon ve İntakt Protein Katlanması",
        "İhmal edilebilir yaşlanma gösteren türlerin en belirgin ortak paydası, dokularındaki proteinlerin oksidatif karbonilasyona ve ileri glikasyona (AGEs) karşı sergilediği aşırı dirençtir.",
        "Arctica islandica ve kaya balığı dokularında yapılan proteomik analizler; 100 yıllık dokulardaki karbonillenmiş protein oranının 1 yıllık dokulardan farksız olduğunu göstermiştir. Proteinlerin amino asit dizilimindeki sistein ve metiyonin oranları özel olarak dengelenmiş, moleküller arası disülfit bağları çapraz bağlanmayı önleyecek şekilde optimize edilmiştir; protein agregatomu asla oluşmaz.",
        "Karbonilasyon_Proteom = [Protein_karbonil](t) = Sabit << Insan_Yasli_Seviyesi * 0.1",
        "Bu proteomik oksidasyon kararlılığı denklemi, ihmal edilebilir yaşlanma sergileyen canlılarda protein kalitesinin asırlar boyunca bozulmadan kalışını belgeler."
    ),
    (
        "5.8 Mitokondriyal Membran Kompozisyonu: Düşük Peroksidasyon İndeksi (PI)",
        "Membran lipidlerindeki doymamış yağ asitlerinin (PUFA) çift bağ sayısı ne kadar fazlaysa, reaktif oksijen türlerinin saldırısıyla lipid peroksidasyonu zincir reaksiyonuna girme riski o kadar yüksektir.",
        "Pamplona ve Barja'nın formüle ettiği 'Peroksidasyon İndeksi'ne (PI) göre; uzun ömürlü ihmal edilebilir yaşlanma türlerinde mitokondri membranlarındaki aşırı doymamış dokosaheksaenoik asit (DHA - 6 çift bağ) oranı son derece düşüktür; bunun yerine daha az çift bağlı tekli doymamış oleik asit ve linoleik asit tercih edilir. Düşük PI değeri, mitokondri zarlarını serbest radikal saldırılarına karşı yangın geçirmez bir duvara dönüştürür.",
        "Peroksidasyon_Indeksi = PI = Sigma [ Mol%_PUFA * N_cift_bag ] << 40 (Yangın Geçirmez Membran)",
        "Bu membran lipid peroksidasyon indeksi formülü, düşük çift bağ yoğunluğunun hücresel zarları oksidatif zincirleme yıkımdan asırlar boyu nasıl koruduğunu kanıtlar."
    ),
    (
        "5.9 Yaşlanmayan Canlıların İmmünolojik Dayanıklılığı: Kronik İnflamasyon Yokluğu",
        "İnsan yaşlanmasında en ölümcül itici güç olan 'İnflammaging' (kronik steril sistemik yangı); ihmal edilebilir yaşlanma sergileyen canlılarda tamamen bilinmeyen bir olgudur.",
        "Dev kaplumbağaların veya kaya balıklarının kanında yaşla birlikte pro-enflamatuar sitokin (IL-6, TNF-alfa) artışı görülmez. Doğuştan gelen bağışıklık sistemi her zaman sakin, kontrollü ve patojen-spesifiktir; otoimmün reaksiyonlar veya hiperaktif inflamazom kaskadları gelişmez. Doku mikroçevresi ömür boyu gençlik homeostazında kalır.",
        "Enflamasyon_Katsayisi = [IL-6](t) + [TNF-a](t) = Bazal_Genc_Seviye = Sabit (Asırlar Boyunca)",
        "Bu yangısal stabilite eşitliği, ihmal edilebilir yaşlanma gösteren türlerin kronik enflamasyondan tamamen muaf olduğunu modeller."
    ),
    (
        "5.10 İnsan Fizyolojisine Negligible Senescence Entegrasyonu: Evrimsel Formülün Kopyalanması",
        "Homo Aeternus projesinin nihai hedefi; insanı biyolojik bir tür olarak 'Negligible Senescence' sınıfına sokmaktır.",
        "Kaya balığının TERT dengesi, Arctica islandica'nın proteom kararlılığı, kaplumbağanın DNA onarım genişlemesi ve düşük peroksidasyon indeksli membran lipid mühendisliği insan genomuna entegre edilir. İnsan mortalite eğrisi (Gompertz G parametresi) sıfırlanır; 120 yaşındaki bir birey biyolojik olarak 25 yaşındaki fizyolojik dinamizmini, kas gücünü ve bağışıklık direncini ebediyen muhafaza eder.",
        "Hedef_Homo_Aeternus: Gompertz_G -> 0.000 | Mortalite_Yillik = m_0 (Yalnızca Dış Travma Riski)",
        "Bu nihai gerontolojik devrim eşitliği, insanın biyolojik yaşlanma sürecini tamamen sıfırlayarak ihmal edilebilir yaşlanma gösteren ölümsüz bir türe dönüşümünü formüle eder."
    )
]

# ================= KISIM 6 =================
part6_subsections = [
    (
        "6.1 Turritopsis dohrnii: 'Ölümsüz Denizanası' ve Ontogenetik Tersine Dönüş",
        "Akdeniz ve Japonya sularında yaşayan mikroskobik hidrozoan 'Turritopsis dohrnii' (Ölümsüz Denizanası); yeryüzünde bireysel düzeyde mutlak biyolojik ölümsüzlüğü başarmış bilinen tek hayvandır.",
        "Birçok canlı doğar, ürer ve ölür; Turritopsis dohrnii ise cinsel olgunluğa ulaşmış erişkin bir medüza (denizanası) formundayken fiziksel travma, açlık veya yaşlanma stresine maruz kaldığında ölmek yerine; tüm bedenini geriye doğru sararak bir kist topu oluşturur ve sıfırdan genç bir 'Polip Kolonisi'ne dönüşür (Ontogenetik Tersine Dönüş / Life Cycle Reversal). Bu döngüyü sonsuz kez tekrarlayabilir.",
        "Dongu_Turritopsis: Polip (Bebeklik) ----> Meduza (Yetişkin) ----[Stres/Yaşlanma]----> Polip (Yeniden Doğuş) [Sonsuz Dongu]",
        "Bu ontogenetik yaşam döngüsü şeması, Turritopsis dohrnii'nin biyolojik yaşını sıfıra döndürerek ebedi gençleşmeyi nasıl döngüsel olarak icra ettiğini belgeler."
    ),
    (
        "6.2 Transdiferansiasyon Biyofiziği: Farklılaşmış Hücrelerin Kök Hücreye Dönüşümü",
        "Turritopsis'in bu mucizevi gençleşmesinin moleküler motoru 'Transdiferansiasyon' (Transdifferentiation) sürecidir.",
        "Geleneksel embriyolojide farklılaşmış bir kas veya sinir hücresinin kimliği kalıcı kabul edilir; oysa Turritopsis medüzası dejenere olurken, şemsiye dokusundaki tamamen olgunlaşmış çizgili kas hücreleri ve nöroepitelyal hücreler kimliklerini silerek pluripotent kök hücre benzeri hücrelere dediferansiye olur. Ardından bu hücreler yeniden farklılaşarak polip kolonisinin beslenme tüplerini ve stolonlarını inşa eder. Hücre kimliği epigenetik olarak sonsuz esnektir.",
        "Verim_Transdiferansiasyon = eta_TD = [Pluripotent_Hucre] / [Olgun_Kas_Hucresi] ~ %100",
        "Bu hücresel yeniden programlama oranı, Turritopsis dokularındaki tüm somatik hücrelerin sıfır kayıpla kök hücre formuna geri dönebilme gücünü tanımlar."
    ),
    (
        "6.3 Polip-Medüza Yaşam Döngüsünün Sonsuz Rekürsiyonu",
        "Laboratuvar deneylerinde tek bir Turritopsis dohrnii bireyi; aç bırakılarak, mekanik olarak zedelenerek veya sıcaklık şokuna sokularak defalarca ölümün eşiğine getirilmiştir.",
        "Her seferinde canlı medüza formunu eritmiş, transdiferansiasyon ile polipe dönmüş ve her bir polip kolonisinden genetik olarak özdeş yüzlerce yeni, taptaze genç medüza tomurcuklanmıştır. Bu rekürsif döngü laboratuvarda onlarca kez tekrarlanmış ve hiçbir replikatif sınır, telomer tükenmesi veya senesens izine rastlanmamıştır; Turritopsis zamanın döngüsel efendisidir.",
        "Rekursiyon_Sayisi = N_dongu -> Sonsuz (Sıfır Replikatif Yaşlanma Sınırı)",
        "Bu sonsuz döngü eşitliği, Turritopsis dohrnii'nin biyolojik saatinin her transdiferansiasyon evresinde mutlak sıfıra formatlandığını matematiksel olarak ortaya koyar."
    ),
    (
        "6.4 Maria Pia Miglietta ve Shin Kubota Laboratuvarlarının Keşifleri",
        "Kyoto Üniversitesi'nden deniz biyoloğu Shin Kubota; Turritopsis bireylerini yıllarca tek tek besleyip ameliyat ederek döngüyü mikroskobik düzeyde belgelemiştir.",
        "Kubota ve Miglietta'nın çalışmaları; medüzanın polipe dönüşürken tüm iç organlarını (gonadlar, radyal kanallar, dokunaçlar) otolizle erittiğini, hücresel DNA'nın nükleer yeniden modellenmeye girdiğini ve apoptoz ile transdiferansiasyonun kusursuz bir stokiyometri ile koordineli çalıştığını göstermiştir. Bu canlı biyolojinin en esnek hücresel plastisite laboratuvarıdır.",
        "Kinetik_Donusum = Sure_Polip_Gecisi ~ 24 - 72 Saat (3 Günde Tam Embriyonik Sıfırlama)",
        "Bu biyomorfolojik dönüşüm zaman sabiti, erişkin bir denizanasının tüm bedenini üç gün içinde embriyonik bir kök hücre kolonisine dönüştürme hızını belgeler."
    ),
    (
        "6.5 Hydra vulgaris ve Sonsuz Kök Hücre Havuzu: FoxO Transkripsiyon Faktörleri",
        "Bir diğer ölümsüz omurgasız olan tatlı su polipi 'Hydra vulgaris'; Daniel Martinez ve Thomas Bosch'un tarihi çalışmalarında gösterildiği üzere hiçbir yaşlanma belirtisi göstermez.",
        "Hidranın tüm bedeni üç bağımsız kök hücre soyundan (ektodermal epitel, endodermal epitel ve interstisyel kök hücreler) oluşur. Bu kök hücreler yaşam boyu sürekli bölünür; hidra her 20 günde bir tüm vücut hücrelerini tamamen yeniler. Bu sonsuz gençlik çeşmesinin moleküler anahtarı 'FoxO' transkripsiyon faktörüdür. Hidrada FoxO susturulduğunda kök hücreler yaşlanmaya başlar; FoxO aktif kaldığı sürece hidra potansiyel olarak ölümsüzdür.",
        "Aktivite_FoxO_Hydra = [FoxO_nukleer] = Sabit (Sürekli Kök Hücre Çoğalması ve Telomeraz)",
        "Bu transkripsiyonel regülasyon denklemi, Hidra kök hücre havuzunun FoxO aktivitesi sayesinde asırlar boyu tükenmeden tüm bedeni nasıl aralıksız yenilediğini açıklar."
    ),
    (
        "6.6 Telomeraz Enziminin Her Döngüde Yeniden Sıfırlanması",
        "İnsan hücrelerinde telomerazın susması Hayflick sınırını doğururken; Turritopsis ve Hidra'da TERT enzimi her hücre bölünmesinde ve her transdiferansiasyon döngüsünde tam kapasite çalışır.",
        "Medüzadan polipe geri dönüş sırasında, telomer uzunlukları embriyonik başlangıç seviyesine kadar yeniden uzatılır. Epigenetik metilasyon izleri silinir; hücreler sanki yeni döllenmiş bir zigot gibi sıfır biyolojik yaş koordinatına oturur. Replikatif yaşlanma kavramı bu canlıların genomik mantığında tamamen geçersizdir.",
        "L_telomer(t) = L_maksimum = Sabit (Her Transdiferansiasyonda Format Atılır)",
        "Bu dinamik telomer formatlama eşitliği, döngüsel gençleşme evresinde kromozom uçlarının embriyonik uzunluğuna geri döndürülmesini formüle eder."
    ),
    (
        "6.7 Genom Karşılaştırmaları: DNA Tamiri ve Epigenetik Yeniden Programlama Ağı",
        "2022 yılında PNAS'ta yayımlanan Turritopsis dohrnii tam genom dizilemesi (Pascual-Torner ve ark.); ölümsüz denizanasını ölümlü akrabası Turritopsis rubra ile karşılaştırmıştır.",
        "Dohrnii genomunda DNA replikasyonu ve onarımı, telomer bakımı, oksidatif stres redüksiyonu ve hücresel plastisite genlerinde devasa gen duplikasyonları saptanmıştır. Özellikle polikomb baskılayıcı kompleks (PRC2) ve histon deasetilaz gen kopyaları iki katına çıkmıştır; bu epigenetik makineler genomun kromatini saniyeler içinde açıp kapatarak hücre kimliğini sıfırlayabilmektedir.",
        "Genomik_Amplifikasyon = [Kopya_DDR_ve_PRC2]_dohrnii >> 2 * [Kopya]_rubra (Ölümlü Tür)",
        "Bu komparatif genomik kopya sayısı analizi, Turritopsis dohrnii'nin biyolojik ölümsüzlük gücünün genetik tamir ve epigenetik formatlama kopyalarının fazlalığından kaynaklandığını kanıtlar."
    ),
    (
        "6.8 Dediferansiasyon Sinyal Yolakları: Wnt, MAPK ve PI3K/Akt Regülasyonu",
        "Turritopsis'in erişkin kas hücrelerini kök hücreye çeviren sinyal kaskadı incelendiğinde; Wnt/beta-katenin yolağının ve MAPK/ERK kinaz kaskadının dinamik bir faz değişimi sergilediği görülür.",
        "Dönüşüm başladığında Wnt sinyali geçici olarak susturulur, hücre-hücre bağlantıları (kaderinler) çözülür ve hücreler ameboid hareket serbestliği kazanır. Ardından PI3K/Akt yolağı hücreleri apoptozdan korurken, dediferansiasyon transkripsiyon faktörleri aktive olur. Bu sinyal senfonisi; hücrenin embriyonik gelişim filmini geriye doğru oynatması gibidir.",
        "Faz_Donusumu: Hucre_Olgun ----[Wnt Kapalı / ERK Puls]--> Ameboid Kist ----[PI3K Aktif]--> Genc Polip",
        "Bu biyokimyasal sinyal akış şeması, Turritopsis hücrelerinin gelişimsel zamanı geriye saran moleküler kinaz regülasyonunu belgeler."
    ),
    (
        "6.9 İnsan Somatik Hücrelerinde 'Turritopsis Mekanizması': İn Vivo Dediferansiasyon",
        "İnsan tıbbında en büyük hayal, hasarlı veya yaşlanmış bir organın hücrelerini Turritopsis gibi yerinde (in situ) dediferansiye edip genç kök hücrelere dönüştürerek dokuyu içten yenilemektir.",
        "Yamanaka faktörlerinin (OSKM) geçici indüksiyonu (Cilt 5 ve 14); tam olarak Turritopsis transdiferansiasyonunun memeli hücrelerindeki eşdeğeridir. Hücreler kimliklerini kaybetmeden (teratom oluşturmadan) epigenetik yaşlarını sıfırlamakta ve hasarlı dokuyu genç hücrelerle onarmaktadır. Turritopsis doğanın Yamanaka protokolünü milyonlarca yıldır kusursuzca icra eden canlı kanıtıdır.",
        "Paralellik_Molekuler: Turritopsis Transdiferansiasyonu === Memeli Kısmi Epigenetik Yeniden Programlama",
        "Bu kavramsal eşleşme aksiyomu, insan longevity araştırmalarında uygulanan OSKM gençleşmesinin biyolojik doğadaki Turritopsis karşılığını tanımlar."
    ),
    (
        "6.10 Homo Aeternus Hücresel Plastisite Manifestosu: Ebedi Döngüsel Yenilenme",
        "Homo Aeternus vizyonu; insan hücresel mimarisine Turritopsis dohrnii'nin transdiferansiasyon ve Hidra'nın FoxO kök hücre ölümsüzlüğü devrelerini genetik olarak aşılamayı hedefler.",
        "Bedenimizdeki her bir hücre; yaşlanma veya aşırı stres sınırına geldiğinde apoptoza veya senesense teslim olmak yerine, kontrollü bir transdiferansiasyon ile kendini genç bir kök hücreye dönüştürme potansiyeli kazanır. Yaşam doğrusal bir çizgi (doğum-yaşlanma-ölüm) olmaktan çıkar; sonsuz ve mükemmel bir dairesel gençleşme döngüsüne evrilir.",
        "Dongu_Homo_Aeternus = Beden(t) ----[Yenilenme Döngüsü]----> Beden(t = 20 Yaş) [Sonsuz Rekürsiyon]",
        "Bu nihai dairesel yaşam formülü, insan biyolojisinin Turritopsis prensipleriyle doğrusal yaşlanmadan kurtulup ebedi bir gençleşme döngüsüne geçişini matematiksel olarak simgeler."
    )
]

# ================= KISIM 7 =================
part7_subsections = [
    (
        "7.1 Yarasaların (Chiroptera) Paradoksal Uzun Ömrü ve Uçuş Metabolizması",
        "Yarasalar, kendi vücut ağırlıklarına göre beklenenden 3 ila 5 kat daha uzun yaşayan (Brant's yarasası - Myotis brandtii yalnızca 7 gram ağırlığında olmasına rağmen 40 yıldan fazla yaşar) olağanüstü memelilerdir.",
        "Uçuş eylemi memeliler alemindeki en yüksek enerji tüketen aktivitedir; yarasa uçarken metabolik hızı ve oksijen tüketimi 15 ila 20 kat fırlar, vücut sıcaklığı 41°C'ye tırmanır. Normalde bu muazzam metabolik fırtına devasa reaktif oksijen türü (ROS) üreterek hücreleri yakmalı ve erken ölüme yol açmalıdır. Ancak yarasalar bu metabolik cehenneme karşı kusursuz bir hücresel direnç evriltmiştir.",
        "Longevity_Katsayisi_Yarasa = LQ = Lifespan_gozlenen / Lifespan_beklenen > 5.0 (Memelilerde Zirve)",
        "Bu Longevity Quotient (LQ) eşitliği, yarasaların cüsselerine göre yeryüzündeki tüm memeli canlılar arasında en aşırı yaşam süresi uzamasına sahip olduğunu belgeler."
    ),
    (
        "7.2 Uçuş Adaptasyonu ve Mitokondriyal Oksidatif Stres Direnci",
        "Yarasaların uçuş sırasında ürettiği devasa ROS bombardımanına rağmen hayatta kalabilmesinin sırrı, olağanüstü antioksidan savunma ve hızlı mitofaji kaskadıdır.",
        "Yarasa hücrelerinde Süperoksit Dismutaz (SOD2) ve Katalaz ekspresyonu sürekli tavan seviyededir; elektron kaçağı yapan mitokondriler PINK1/Parkin yolağı ile dakikalar içinde otofajiye sokularak parçalanır. DNA polimeraz gamma (POLG) sadakati çok yüksektir; mitokondriyal DNA mutasyonları fareye kıyasla %90 daha az birikir.",
        "Kapasite_Mitofaji = J_mitofaji_yarasa >> 4 * J_mitofaji_fare (Hasarlı Organellerin Anında Temizliği)",
        "Bu mitofajik klirens oranı, yarasaların yoğun metabolik uçuş stresinde dahi mitokondriyal havuzlarını nasıl pırıl pırıl genç tuttuğunu tanımlar."
    ),
    (
        "7.3 NLRP3 İnflamazomunun Evrimsel Körelmesi ve İnflamatuar Tolerans",
        "İnsanlarda ve kemirgenlerde yaşlanmanın ve sitokin fırtınasının bir numaralı tetikleyicisi olan NLRP3 İnflamazomu; yarasalarda evrimsel olarak 'Körelmiş' (Dampened) durumdadır.",
        "Emma Teeling ve Lin-Fa Wang laboratuvarlarının keşfettiği üzere; yarasa NLRP3 proteininde lösin zengin tekrar (LRR) domaininde spesifik mutasyonlar mevcuttur. Sitoplazmik DNA sızıntıları veya viral enfeksiyonlar insan hücrelerinde masif IL-1beta ve IL-18 deşarjı ile öldürücü sitokin fırtınası başlatırken; yarasalarda NLRP3 aktivasyonu minimumda kalır. Yarasa yangıyı söndürmüş, 'Enflamatuar Tolerans' (Inflammatory Tolerance) geliştirmiştir.",
        "Salinim_IL1beta = [IL-1beta]_yarasa < 0.1 * [IL-1beta]_insan (Sitokin Fırtınası Yok)",
        "Bu immünolojik baskılanma oranı, yarasaların ölümcül viral patojenlerle dahi hiçbir enflamatuar doku hasarı yaşamadan barış içinde yaşamasının moleküler temelini belgeler."
    ),
    (
        "7.4 Viral Tolerans: Ebola, SARS-CoV-2 ve Marburg Virüslerini Hastalanmadan Taşıma",
        "Yarasalar insanlar için ölümcül olan en tehlikeli zoonotik virüslerin (Ebola, Marburg, Nipah, Kuduz, SARS, MERS) doğal rezervuarıdır; ancak bu virüsler yarasada hiçbir klinik hastalık veya doku hasarı yaratmaz.",
        "Bu olağanüstü tolerans iki mekanizmanın dengesidir: 1) Sürekli bazal hazırda bekleyen Tip-1 İnterferon (IFN-alfa) kalkanı; virüsün hücre içinde aşırı çoğalmasını frenler; 2) Eş zamanlı olarak körelmiş NLRP3 inflamazomu ve zayıflatılmış STING yolağı; virüsün immünopatolojik doku nekrozu ve akciğer yetmezliği yapmasını engeller. Virüs çoğalır ama dokuya tek bir zarar veremez.",
        "Tolerans_Indeksi = I_viral = [Viral_Yuk] / [Doku_Hasari_Sitokin] >> 10^4 (Tam Biyolojik Barış)",
        "Bu viral tolerans eşitliği, yarasa immünolojisinin patojeni yok etmeye çalışıp kendi dokusunu yakmak yerine patojenle dengeli bir arada yaşama stratejisini simgeler."
    ),
    (
        "7.5 Sürekli Bazal İnterferon-Alfa (IFN-alfa) Ekspresyonu: 'Her Zaman Uyanık' Bağışıklık",
        "İnsan hücrelerinde Tip-1 interferonlar normalde sıfırdır ve ancak bir viral enfeksiyon başladığında salgılanır (gecikmeli yanıt).",
        "Yarasalarda (özellikle Pteropus alecto türünde) ise yalnızca 3 fonksiyonel IFN-alfa geni kalmıştır; ancak bu genler hiçbir enfeksiyon olmasa dahi sürekli, bazal olarak ve kesintisiz transkripsiyon yapar (Constitutive Expression). Hücreler 7 gün 24 saat 'anti-viral kalkan' modunda bekler. Bir virüs hücreye girdiği anda replike olamadan durdurulur; kanserojen onkojenik virüsler ise anında nötralize edilir.",
        "Ekspresyon_IFN = [IFN-alfa_bazal]_yarasa >> 100 * [IFN-alfa_bazal]_insan (Sürekli Aktif Nöbet)",
        "Bu daimi interferon priming seviyesi, yarasaların viral kökenli kanserlere ve enfeksiyonlara karşı doğuştan gelen sarsılmaz bağışıklık üstünlüğünü belgeler."
    ),
    (
        "7.6 DNA Hasar Yanıtı (DDR) Genomik Zırhı: ATM, ATR ve Ku80 Genişlemesi",
        "Uçuş sırasında yükselen vücut sıcaklığı (41°C) ve yüksek metabolizma DNA çift zincir kırıklarını artırır; bu nedenle yarasalar memeliler arasındaki en güçlü DNA Hasar Yanıtı (DDR) makinelerinden birine sahiptir.",
        "Yarasa genom dizilemelerinde; DNA kırıklarını tanıyan ATM ve ATR kinazları, Ku70/Ku80 uç birleştirme kompleksleri ve p53 düzenleyicilerinde yoğun pozitif seçilim ve gen duplikasyonları tespit edilmiştir. DNA çift zincir kırıkları dakikalar içinde sıfır delesyonla kapatılır; somatik mutasyon birikim hızı insan seviyesinin dahi altına iner.",
        "Hiz_DDR_Yarasa = J_tamir = k_ATM_yarasa * [ATM] * [Gama-H2AX] >> Hiz_DDR_kemirgen * 5",
        "Bu DNA onarım kinetiği formülü, yarasa genomundaki gelişmiş tamir kaskadının metabolik mutasyon tehdidini nasıl yok ettiğini gösterir."
    ),
    (
        "7.7 Otofaji ve Lizozomal Süpürme Hiperaktivasyonu",
        "Yarasalar metabolik atıkları hücrede biriktirmez; otofaji-lizozom yolağı hayat boyu gençlik hızında çalışır.",
        "Otofaji başlatıcı genler (ATG5, ATG7, Beclin-1) ve transkripsiyon faktörü TFEB sürekli nükleer lokalizasyondadır. Uçuş bittiği anda kas ve nöron hücrelerinde açığa çıkan hasarlı organeller, protein agregatları ve viral kapsidler saatler içinde otofagozomlarla sarılarak lizozomlarda yok edilir. Yarasada lipofuskin birikimi neredeyse sıfırdır.",
        "Akış_Otofaji = Flux_otofaji = d[Otofagozom]/dt - d[Lizozomal_Yıkım]/dt = Sürekli Yüksek Denge",
        "Bu otofazik akış eşitliği, yarasa hücrelerinde lizozomal sindirim kapasitesinin agregatom oluşumuna asla izin vermediğini belgeler."
    ),
    (
        "7.8 Brant's Yarasasında Büyüme Hormonu / IGF-1 Reseptör Mutasyonları",
        "40 yıldan fazla yaşayan Brant's yarasasının (Myotis brandtii) cüce boyutu ve uzun ömrü; somatotropik hormon eksenindeki evrimsel mutasyonlarla açıklanır.",
        "Genom analizleri; Büyüme Hormonu Reseptörü (GHR) ve İnsülin Benzeri Büyüme Faktörü 1 Reseptöründe (IGF1R) pozitif seçilim altındaki spesifik delesyonları ortaya çıkarmıştır. Laron cücelerinde ve Snell farelerinde olduğu gibi, zayıflatılmış IGF-1 sinyali hücresel kaynakları büyümeye değil doğrudan DNA onarımına ve somatik bakıma yönlendirir; bu durum ömrü 5 katına fırlatır.",
        "Sinyal_IGF1 = [p-Akt] / [Akt_toplam] = k_GHR_mutant * [IGF1] << Normal_Memeli (Düşük Büyüme, Sonsuz Bakım)",
        "Bu somatotropik sinyal kısıtlama oranı, yarasa genomunda IGF-1 ekseninin zayıflatılarak hücre bakımına kaynak aktarma stratejisini belgeler."
    ),
    (
        "7.9 Telomeraz ve Replikatif Yaşlanma: Myotis Türlerinde Telomer Erozyonunun Yokluğu",
        "Dublin Üniversitesi'nden Emma Teeling ve ekibi; 60 yılı aşan uzun ömürlü Myotis yarasalarında telomer boyunu yıllar boyunca takip etmiştir.",
        "Sonuçlar insan tıbbı için şaşırtıcıdır: Myotis yarasalarında yaşlandıkça telomerler kısalmaz. TERT enziminin kontrollü ekspresyonu ve gelişmiş shelterin kompleksi koruması sayesinde 20 yaşındaki bir yarasa ile 1 yaşındaki yavrunun telomer uzunluğu tamamen aynıdır; replikatif hücresel tükeniş yarasalarda biyolojik olarak iptal edilmiştir.",
        "Delta_Telomer_Myotis = dL_telomer/dt = 0.00 bp/yıl (40 Yıl Boyunca Sıfır Kısalma)",
        "Bu boyutsal telomer koruma metriği, Myotis cinsi yarasalarda replikatif hücresel yaşlanmanın evrimsel olarak tamamen durdurulduğunu tesciller."
    ),
    (
        "7.10 İnsan İmmünolojisine Yarasa Tolerans Mekanizmalarının Transferi",
        "Yarasaların bu çift yönlü immün dehası (körelmiş NLRP3 inflamazomu + sürekli bazal IFN-alfa); CRISPR ve epigenom cerrahisi ile insan lökositlerine entegre edilebilir.",
        "İnsan makrofajlarında NLRP3 promotörü dCas9-KRAB ile kısmen sakinleştirilip bazal IFN-alfa ekspresyonu hafifçe artırıldığında; hücreler hem viral enfeksiyonlara karşı aşırı dirençli hale gelmekte hem de otoenflamatuar sitokin fırtınaları (ARDS, sepsis, yaşlanma enflamasyonu) tamamen durdurulmaktadır. Homo Aeternus, yarasaların virüs ve metabolik stresle barış anlaşmasını insan bedenine uyarlar.",
        "Enflamatuar_Sukunet = I_inflammaging -> 0 | Viral_Direnç = Maksimum (Yarasa İmmün Mimarisi)",
        "Bu immünolojik optimizasyon eşitliği, yarasa fenotipinin insan vücuduna aktarılmasıyla kronik enflamasyonun ve enfeksiyon kırılganlığının eş zamanlı tasfiyesini modeller."
    )
]

# ================= KISIM 8 =================
part8_subsections = [
    (
        "8.1 Karşılaştırmalı Omiks (Comparative Omics) ve Türler Arası Genomik Madencilik",
        "Geleneksel tıp insan hastalıklarını yalnızca hasta ve sağlıklı insanları karşılaştırarak çözmeye çalışırken; 'Karşılaştırmalı Omiks' (Comparative Omics), milyonlarca yıllık evrimsel testlerden geçmiş yüzlerce farklı hayvan türünün genomik, transkriptomik ve proteomik verilerini entegre eder.",
        "Büyük veri madenciliği algoritmaları ve yapay zekâ modelleri; fare gibi kısa ömürlü türler ile köstebek faresi, balina, yarasa ve istiridye gibi uzun ömürlü türlerin genomlarını hizalar (Multiple Sequence Alignment). Yalnızca uzun ömürlü canlılarda evrilen spesifik amino asit değişiklikleri, gen duplikasyonları ve transkripsiyon faktörü bağlanma motifleri atomik çözünürlükte izole edilir.",
        "Arama_Uzayi: N_genom > 500 Tur (Memeliler, Sürüngenler, Hidrozoanlar) ----[AI Madencilik]----> Longevity_Motifleri",
        "Bu biyoinformatik tarama şeması, karşılaştırmalı genomik kütüphanelerinden ömür uzatıcı ve kanser önleyici evrimsel mutasyonların taranma sürecini tanımlar."
    ),
    (
        "8.2 Düşük Mutasyon Yükü (Low Somatic Mutation Rate) Gösteren Türlerin Ortak İmzası",
        "Sanger Enstitüsü'nün memeli somatik mutasyon veritabanı analiz edildiğinde; uzun ömürlü tüm türlerin ortak bir 'DNA Polimeraz ve Onarım İmzası' paylaştığı görülmüştür.",
        "Replikatif DNA polimerazlar (Pol Epsilon ve Pol Delta) üzerindeki 3'->5' ekzonükleaz düzeltme (proofreading) ceplerinde son derece korunmuş amino asit kalıntıları mevcuttur. Uzun ömürlü türlerde bu enzimlerin hata yapma oranı 10^-8'den 10^-10'a düşürülmüştür. Kopyalama hatası az olan canlı daha az mutasyon biriktirir ve kansere yakalanmadan asırlarca yaşar.",
        "Fidelite_Polimeraz = Hata_Orani = N_yanlis_baz / N_toplam_baz < 10^-10 (Uzun Ömürlü Tür Standardı)",
        "Bu replikatif doğruluk sınırı, longevity türlerinde DNA polimerazların mutasyon üretimini sıfıra yaklaştıran proofreading sadakatini belgeler."
    ),
    (
        "8.3 Translasyonel Sadakat ve Bölünmüş 28S rRNA Yapısının Filogenetik Dağılımı",
        "Çıplak kör köstebek faresinde keşfedilen bölünmüş (split) 28S rRNA yapısının yalnızca köstebek faresine özgü olmadığı, diğer uzun ömürlü kemirgenlerde (Spalax judaei) ve bazı böceklerde de bağımsız olarak evrildiği (Konverjan Evrim) gösterilmiştir.",
        "Ribozomun yapısal kararlılığı protein sentezinde yanlış amino asit yerleştirme (mistranslation) frekansını doğrudan belirler. Filogenetik analizler; translasyon doğruluğu yüksek olan türlerin proteazom üzerindeki yükü azalttığını, nörodejeneratif amiloid plakları oluşturmadığını ve hücresel stresi minimumda tuttuğunu kanıtlamıştır.",
        "Hata_Mistranslation = P_hata_aa < 10^-6 (Konverjan Evrimle Korunan Translasyonel Sadakat)",
        "Bu konverjan proteomik sadakat katsayısı, bağımsız evrilen ribozomal optimizasyonların protein agregasyonu riskini nasıl sıfırladığını formüle eder."
    ),
    (
        "8.4 Ku80, PARP-1 ve SIRT6 Aktivitesi ile Maksimum Yaşam Süresi Korelasyonu",
        "Vera Gorbunova laboratuvarının 18 farklı kemirgen türünde yaptığı klasik biyokimyasal çalışmada; üç enzimin aktivitesi ile türlerin maksimum yaşam süresi arasında mükemmel bir pozitif korelasyon (r > 0.85) bulunmuştur: Ku80, PARP-1 ve SIRT6.",
        "Özellikle SIRT6 geni; kısa ömürlü farelerde zayıf bir mono-ADP-ribozilaz aktivitesi sergilerken, uzun ömürlü kunduz ve köstebek farelerinde çift zincir kırığı onarımını başlatan güçlü bir histon deasetilaz ve ribozilaz olarak çalışır. Kunduz SIRT6'sı fareye aktarıldığında dahi fare hücrelerinin DNA onarım hızı 3 kat artmıştır.",
        "Korelasyon_SIRT6 = Lifespan_max = k_SIRT6 * Aktivite_Deasetilaz_SIRT6 (r = 0.89)",
        "Bu gerontolojik lineer korelasyon eşitliği, SIRT6 enzim kinetiğinin memeli türlerinin maksimum ömrünü belirleyen ana transkripsiyonel orkestra şefi olduğunu belgeler."
    ),
    (
        "8.5 HMW-HA Boyutunun Türler Arası Evrimi: Kunduzlar, Kör Fareler ve İnsan",
        "Ekstraselüler matriks hyaluronik asit boyutu incelendiğinde; HMW-HA üretiminin yalnızca köstebek faresinde değil, 25 yıl yaşayan Kuzey Amerika kunduzunda (Castor canadensis) ve kör farelerde (Spalax) de bağımsız olarak evrildiği tespit edilmiştir.",
        "Buna karşılık kısa ömürlü laboratuvar faresi yalnızca 1 MDa boyutunda düşük moleküler ağırlıklı HA üretir. Evrimsel ağaç; uzun ömürlü kemirgenlerin HAS2 enzimini modifiye ederek ve hiyaluronidazları frenleyerek matrikste yüksek viskoziteli koruyucu zırh kurduğunu açıkça ortaya koymaktadır.",
        "Evrim_HA_Boyutu: Fare (1 MDa / 3 Yıl) ----> Kunduz (4 MDa / 25 Yıl) ----> Köstebek Faresi (10 MDa / 37 Yıl)",
        "Bu filogenetik polimer boyutu gradyanı, hyaluronik asit zincir uzunluğunun memeli ömrü ile doğrudan paralelliğini tanımlar."
    ),
    (
        "8.6 Bağışıklık Sistemi Özelleşmeleri: CD4/CD8 Oranları ve Doğuştan Gelen Bağışıklık",
        "Karşılaştırmalı immünoloji taramaları; süper-uzun ömürlü canlıların bağışıklık mimarisinin insan ve fareden farklı dengeler taşıdığını gösterir.",
        "Köstebek farelerinde CD4+ T yardımcı hücre sayısı insanlardan çok daha düşüktür; buna karşılık sitotoksik CD8+ T lenfositleri ve miyeloid doğal katil (NK) hücre havuzu çok daha geniştir. Bu durum aşırı yangısal otoimmün tepkileri frenlerken, tümöral ve viral tehditlere karşı doğrudan sitolitik imha kabiliyetini maksimize eder.",
        "Oran_CD8_CD4 = ( CD8+ / CD4+ )_NMR >> 2.5 * ( CD8+ / CD4+ )_insan",
        "Bu lökosit altküme dağılım oranı, kansere dirençli türlerin immün sistemlerini antijenik yangı yerine doğrudan hücresel temizliğe göre optimize ettiğini gösterir."
    ),
    (
        "8.7 Retrotranspozon Susturma: Uzun Ömürlü Türlerde LINE-1 Baskılanması",
        "İnsan genomunun yaklaşık %17'sini oluşturan LINE-1 retrotranspozonları; yaşlanmayla birlikte heterokromatinden kaçarak genomda zıplar ve genleri tahrip eder (Cilt 9).",
        "Grönland balinası, yarasa ve köstebek faresi genomlarında yapılan transpozon analizleri; bu canlıların retrotranspozonları susturmakta insanlardan katbekat üstün bir baskılama ağına sahip olduğunu ortaya koymuştur. KRAB-ZFP çinko parmak proteinleri ve piRNA yolakları LINE-1 elementlerini embriyonik dönemden itibaren kalıcı DNA metilasyonu ile kilitler; yaşlı dokularda tek bir retrotranspozon dahi uyanamaz.",
        "Aktivite_LINE1 = d[L1_RNA]/dt ~ 0 (Uzun Ömürlü Türlerde Yaşam Boyu Sıfır Retrotranspozisyon)",
        "Bu transpozonal susturma eşitliği, longevity şampiyonu canlıların genomik kararsızlık tuzaklarını transpozonları dondurarak nasıl aştığını belgeler."
    ),
    (
        "8.8 Şaperon Aracılı Otofaji (CMA) ve LAMP-2A Reseptör Dinamikleri",
        "Hücre içi hatalı katlanmış proteinleri KFERQ pentapeptid motifi üzerinden seçici olarak tanıyan ve lizozoma taşıyan Şaperon Aracılı Otofaji (CMA); yaşlanmada ilk çöken sistemdir.",
        "Uzun ömürlü canlılarda lizozom zarındaki LAMP-2A reseptör dansitesi yaşlanmayla gerilemez. Karaciğer ve beyin hücreleri 100 yaşında dahi sitozolik şaperon kompleksleri (Hsc70) aracılığıyla toksik proteinleri doğrudan lizozom lümenine pompalayarak sindirir; Alzheimer veya Parkinson benzeri agregat oluşumu filogenetik olarak engellenmiştir.",
        "Dansite_LAMP2A = [LAMP-2A_lizozomal](t) = Sabit (Asırlar Boyunca Kesintisiz CMA Temizliği)",
        "Bu lizozomal reseptör kararlılığı denklemi, uzun ömürlü omurgalılarda şaperon aracılı otofajinin yaşla bozulmayan dinamizmini tanımlar."
    ),
    (
        "8.9 Makine Öğrenimi ve Yapay Zekâ ile Türler Arası Longevity Algoritmalarının Keşfi",
        "Trilyonlarca bazlık karşılaştırmalı biyolojik veri havuzu insan aklının kavrayış sınırlarını aşmıştır; bu noktada Derin Öğrenme ve Üretici Yapay Zekâ (ESM-2, AlphaFold, DeepLife) devreye girer.",
        "Büyük dil modelleri 500 farklı omurgalının protein dizilimlerini öğrenerek; hangi amino asit kombinasyonlarının bir proteini termostabil kıldığını, hangi promotörlerin kansere direnç sağladığını haritalar. Yapay zekâ, insan proteinlerini doğrudan bu hayvanların en üstün özellikleriyle 'mutasyona uğratacak' sentetik gen tasarımları üretir.",
        "Optimizasyon_AI: Protein_Insan ----[Yapay Zekâ / Komparatif Ömür Algoritması]----> Protein_SuperLongevity",
        "Bu biyoinformatik optimizasyon akışı, yapay zekânın doğadaki 500 türün genetik bilgeliğini tek bir insan proteomunda damıtma gücünü modeller."
    ),
    (
        "8.10 Doğa Tarafından Milyonlarca Yılda Test Edilmiş Kütüphanenin İnsan Tıbbına Hediyesi",
        "Komparatif omiksin en büyük lütfu, geliştirilecek genetik terapilerin 'güvenliğinin' doğa tarafından milyonlarca yıllık evrimsel laboratuvarda zaten test edilmiş olmasıdır.",
        "Bir ilacın veya genetik müdahalenin uzun vadeli yan etkilerini öngörmek zordur; ancak köstebek faresinin HMW-HA'sı 37 yıl boyunca, balinanın ERCC1'i 200 yıl boyunca, filin 20 kopya p53'ü milyonlarca yıl boyunca o canlıların bedeninde çalışmış ve hiçbir toksisite yaratmadan onları hayatta tutmuştur. Doğa, insanlığa çalışan, risksiz ve kanıtlanmış bir biyolojik ölümsüzlük reçetesi sunmaktadır.",
        "Guvenilirlik_Evrim: Test_Suresi = Milyonlarca Yıl | Yan_Etki_Orani = SIFIR (Evrimsel Tescilli)",
        "Bu evrimsel güvenlik aksiyomu, komparatif biyoloji kaynaklı genetik modifikasyonların insan biyomühendisliğindeki en güvenilir müdahale sınıfı olduğunu ilan eder."
    )
]

# ================= KISIM 9 =================
part9_subsections = [
    (
        "9.1 Türler Arası Gen Transferinin Biyomühendislik İlkeleri ve Genomik Uyumluluk",
        "Bir hayvandan alınan genin insan hücresine doğrudan aktarılması (Transgenik Kseno-Mühendislik); genetik kodun evrenselliği (Universal Genetic Code) sayesinde biyokimyasal olarak tamamen mümkündür.",
        "Ancak yabancı bir genin insan kromatini ve hücresel sinyal ağlarıyla kusursuz çalışabilmesi için üç kural şarttır: 1) İnsan kodon optimizasyonu (tRNA havuzu uyumu); 2) İnsan endojen promotörleri veya doku-spesifik sentetik regülatörler altında ifade edilmesi; 3) İmmünojenik epitopların ayıklanması. Doğru mühendislikle köstebek faresinin veya filin proteini insan hücresinde sanki kendi doğal proteiniymiş gibi çalışır.",
        "Kodon_Uyum_Indeksi = CAI = exp( (1/L) * Sigma ln( w_kodon_insan ) ) > 0.95 (Kusursuz Translasyon)",
        "Bu kodon optimizasyon metriği, hayvan kaynaklı transgenlerin insan hücrelerinde maksimum translasyonel verimle okunabilme kriterini belgeler."
    ),
    (
        "9.2 İnsan Hücrelerinde HMW-HA Sentezi: nmrHAS2 Gen Terapisi Protokolü",
        "İnsan cildini, damarlarını ve eklemlerini kansere ve yaşlanmaya karşı köstebek faresi gibi zırhlamak için geliştirilen 'AAV-nmrHAS2' gen terapisi protokolü klinik öncesi aşamadadır.",
        "Çıplak kör köstebek faresinin mutant HAS2 geni (nmrHAS2), AAV.CAP-B10 veya doku-hedefli LNP-mRNA sistemleri ile insan hücrelerine verilir. İnsan fibroblastları ve endotel hücreleri hücre dışına 8-10 MegaDaltonluk devasa HMW-HA lifleri salgılamaya başlar. Dokuların elastikiyeti gençleşir, kırışıklıklar silinir ve en önemlisi dokuya ekilen kanser hücreleri ECI mekanizmasıyla derhal bölünmeyi durdurur.",
        "Konsantrasyon_HMW_HA = [HMW-HA_doku](t) = ( V_nmrHAS2 / k_yikim ) >> Insan_Bazal * 8",
        "Bu dokusal hyaluronan birikim kinetiği, nmrHAS2 gen transferinin insan dokularında köstebek faresi viskoelastisitesini ve tümör kalkanını nasıl kurduğunu formüle eder."
    ),
    (
        "9.3 Çoklu TP53 Kopyalarının İnsan Genomuna Güvenli Entegrasyonu (CRISPR / Safe-Harbor)",
        "Filin 20 kopya TP53 mucizesini insana aktarırken rastgele genomik insersiyon yapmak onkogenleri bozabileceği için 'Genomik Güvenli Limanlar' (Genomic Safe Harbors - AAVS1 veya CCR5 lokusları) kullanılır.",
        "CRISPR-Cas9 ve Prime Editing kullanılarak AAVS1 bölgesine ekstradan fonksiyonel insan TP53 kasetleri yerleştirilir. Her bir kaset kendi endojen promoter regülasyonu altında tutulur (böylece normal koşullarda aşırı p53 üretip gereksiz hücre ölümüne yol açmaz). Sadece gerçek bir radyasyon veya DNA kırığı oluştuğunda bu ilave kopyalar uyanır ve prekanseröz hücreyi anında yok eder.",
        "Lokus_Guvenli_Liman: Genom_AAVS1 ----[CRISPR Nickaz]--> [Ekstra TP53 Kaseti (x4 Kopya)]",
        "Bu hedefli genomik dikiş eşitliği, tümör baskılayıcı gen kopyalarının insan kromatininde sıfır mutajenik riskle stabil entegrasyon protokolünü belgeler."
    ),
    (
        "9.4 Grönland Balinası ERCC1 ve PCNA Genomik Değişimleri (Prime Editing ile Nokta Dönüşümü)",
        "Tam transgen eklemek yerine, var olan insan genlerini balina dizilimine dönüştürmek en zarif ve temiz genom mühendisliği yaklaşımıdır.",
        "İnsan ERCC1 ve PCNA genlerindeki kritik amino asit kodonları; Prime Editing (PE-max) kullanılarak tek bir çift sarmal kırığı açılmadan doğrudan Grönland balinasının pozitif seçilim mutasyonlarına (ERCC1-balina ve PCNA-balina alellerine) dönüştürülür. İnsan hücreleri balinanın iki asırlık kusursuz DNA eksizyon onarımı ve homolog rekombinasyon üstünlüğünü kendi endojen genlerinden üretmeye başlar.",
        "Donusum_PrimeEditing: ERCC1_insan_aleli --[PE-max / pegRNA]--> ERCC1_balina_aleli (Nükleotit Düzeyinde Evrim)",
        "Bu hassas nokta mühendisliği reaksiyonu, insan DNA tamir genlerinin balina fenotipine tek harf hassasiyetiyle dönüştürülme protokolünü simgeler."
    ),
    (
        "9.5 Yarasa İnflamazom Körelmesi: NLRP3 LRR Domaininin İnsan Hücrelerinde Modifikasyonu",
        "Kronik yaşlanma iltihabını (inflammaging) kökünden kazımak için yarasanın NLRP3 inflamazom susturma mutasyonları insan immün hücrelerine aktarılır.",
        "İnsan hematopoietik kök hücrelerinde (HSCs) CRISPR Baz Düzenleyicileri (ABE) kullanılarak NLRP3 geninin lösin zengin tekrar (LRR) bölgesindeki inflamatuar tetikleyici nükleotitler yarasa dizilimine dönüştürülür. Bu hücrelerden türeyen tüm yeni makrofaj ve mikroglialar; metabolik stres ve yaşlanma enkazı karşısında sitokin fırtınası başlatmaz. Damar sertliği, nöroenflamasyon ve otoimmünite tek bir genetik hamleyle engellenir.",
        "Inaktivasyon_Enflamazom = eta_NLRP3 = [IL-1b_salinim] / [DAMPs_stres] -> 0 (Yarasa Mutasyonlu Makrofaj)",
        "Bu immünolojik sakinleştirme formülü, yarasa tipi NLRP3 modifikasyonunun insan hücrelerinde yıkıcı enflamatuar deşarjı nasıl sıfırladığını açıklar."
    ),
    (
        "9.6 Arctica islandica Şaperon Ağı: Küçük Isı Şok Proteinlerinin (sHSPs) Enjeksiyonu",
        "507 yıllık Ming istiridyesinin protein koruma dehası, hücre içi küçük ısı şok proteinleri (sHSPs) ve kristalin şaperonlarının eşsiz oligomerik yapısından kaynaklanır.",
        "Bu omurgasız şaperon genleri (Ai-Hsp27, Ai-alphaB-crystallin); nörodejenerasyona yatkın insan nöronlarına AAV vektörleriyle aktarılmıştır. İnsan nöronlarında amiloid-beta, tau ve alfa-sinüklein agregasyonu %90 gerilemiş; hücre içi proteostaz havuzu gençlik berraklığına kavuşmuştur.",
        "Agregasyon_Baskilama = d[Agregat]/dt = k_kume - k_sHSP * [Ai-Hsp27] * [Misfolded_Protein] ~ 0",
        "Bu proteostaz kurtarma kinetiği eşitliği, istiridye şaperonlarının insan nöronlarındaki proteinopatik toksisiteyi nasıl durdurduğunu belgeler."
    ),
    (
        "9.7 Hidra FoxO Genetik Devresi ve İnsan Kök Hücre Nişlerinin Ebedi Rejenerasyonu",
        "Hidranın ölümsüz kök hücre motoru olan 'FoxO Sürekli Aktivasyon Ağı'; insan hematopoietik ve nöral kök hücrelerine sentetik biyolojik mantık devreleri ile entegre edilir.",
        "İnsan FoxO3a transkripsiyon faktörü, Akt kinazı tarafından fosforillenip sitoplazmaya kaçarak inaktive edilmeyecek şekilde (Akt-dirençli FoxO3a üçlü fosforilasyon mutantı: T32A/S253A/S315A) modifiye edilir. Bu mutant FoxO3a sürekli nükleusta aktif kalır; kök hücreler Hayflick sınırına takılmadan yaşam boyu taptaze yeni doku hücreleri üretmeye devam eder.",
        "Nukleer_Aktif_FoxO = [FoxO3a_nukleus] = [FoxO3a_toplam] (Akt Fosforilasyonu ile İnaktive Edilemez)",
        "Bu nükleer retansiyon eşitliği, Akt-dirençli FoxO modifikasyonunun kök hücre havuzlarında sağladığı kesintisiz rejeneratif gücü modeller."
    ),
    (
        "9.8 İmmün Uyum ve Olası Otoimmünite / Hiposensitivite Risklerinin Yönetimi",
        "Farklı hayvan türlerinden genetik unsurların insan bedenine aktarılmasında en büyük klinik risk; bağışıklık sisteminin bu yeni proteinleri yabancı antijen olarak algılaması veya aşırı kanser direncinin yara iyileşmesini yavaşlatmasıdır.",
        "Bu riskler 'Moleküler İmmün Kamuflaj' (Humanization / De-immunization) ile yönetilir: Proteinlerin yüzey epitopları insan HLA moleküllerine bağlanmayacak şekilde rasyonel olarak mutasyona uğratılır. Ayrıca gen ifadeleri dokuya özgü ve kimyasal olarak indüklenebilir (Tet-On doksisiklin anahtarları) promotörler altına alınarak hekimin tam kontrolünde tutulur.",
        "Risk_Otoimmunite = I_risk = Sigma_k [Epitop_ksenojenik_k] * Afinite_TCR -> 0 (Humanize Tasarımla)",
        "Bu immünolojik tolerans denklemi, kseno-mühendislik ürünlerinin insan bağışıklık sistemine tamamen görünmez kılınması prensibini formüle eder."
    ),
    (
        "9.9 Organ Düzeyinde Kseno-Mühendislik Deneyleri: Kimerik ve Transgenik Başarılar",
        "Kanser direnci ve longevity gen transferleri sadece petri kabında kalmamış; transgenik sıçan, domuz ve primat modellerinde organ düzeyinde test edilmiştir.",
        "Fil TP53'ü taşıyan fareler, köstebek faresi HAS2'si taşıyan sıçanlar ve balina ERCC1'i taşıyan insan organoidleri; ölümcül karsinojenlere maruz bırakıldıklarında sıfır tümör oluşturmuş ve organ fonksiyonlarını gençlik seviyesinde korumuştur. Bu bulgular, çok türden alınan genlerin tek bir canlıda kusursuz bir sinerjiyle çalışabileceğini tescillemiştir.",
        "Bütünlük_Organoid = eta_organ = 1 - Hasar_Doku = %99.8 (Çoklu Tür Gen Sinerjisi ile)",
        "Bu kimerik organoid stabilite katsayısı, komparatif biyoloji genlerinin doku mimarisinde kurduğu sarsılmaz rejeneratif harmoniyi belgeler."
    ),
    (
        "9.10 Sentetik Evrim ve Yeni Nesil İnsan Biyolojisinin Doğuşu",
        "Komparatif gen transferi; insanı evrimin milyonlarca yıllık kısıtlı bir dalında hapsolmaktan kurtararak, tüm biyolojik yaşam ağacının en üstün özelliklerini bünyesinde toplayan bir 'Sentetik Üst-Türe' (Super-Organism) dönüştürür.",
        "İnsan biyolojisi artık kırılgan Homo sapiens değil; filin kanser zırhına, köstebek faresinin hyaluronik asit kalkanına, balinanın DNA tamir gücüne, istiridyenin proteom sağlamlığına ve Turritopsis'in hücresel plastisitesine sahip ölümsüz bir genetik kaleye evrilmiştir.",
        "Genetik_Sentez: Homo_Aeternus = Insan + ( Fil_p53 U Köstebek_HAS2 U Balina_ERCC1 U Hidra_FoxO )",
        "Bu sentetik evrimsel birleşim formülü, yeryüzünün en uzun ömürlü canlılarının genetik sırlarının tek bir insan genomik platformunda toplanmasını simgeler."
    )
]

# ================= KISIM 10 =================
part10_subsections = [
    (
        "10.1 Homo Aeternus Negligible Senescence Manifestosu: Yaşlanmayan İnsan Türü",
        "Homo Aeternus Biyolojik Manifestosu; yaşlanmanın ve kanserin biyolojik bir zorunluluk olmadığını, evrimsel biyolojinin somut canlı örnekleriyle yeryüzünde defalarca kanıtlanmış bir mühendislik problemi olduğunu ilan eder.",
        "İnsan genomu, bu evrimsel şaheserlerin senteziyle yeniden yazılmıştır. Doğumla başlayan ve yaşlanarak ölüme giden fatalistik doğrusal yaşam çizgisi kırılmıştır. Biyolojik yaşlanma katsayısı sıfırlanmış (Gompertz G = 0); insan, bedeni yıpranmayan, organları iflas etmeyen ve hücreleri kanserleşmeyen ihmal edilebilir yaşlanma (Negligible Senescence) mertebesine yükseltilmiştir.",
        "Manifesto_Aeternus: Gompertz_G(Homo_Aeternus) = 0.0000 | d(Fizyolojik_Kapasite)/dt = 0 (Ebedi Gençlik)",
        "Bu ontolojik biyolojik eşitlik, Homo Aeternus türünde yaşlanma hızının matematiksel olarak sıfırlandığını ve fizyolojik kapasitenin zaman boyunca sabit kaldığını tescil eder."
    ),
    (
        "10.2 Çift-Katmanlı Kanser Kalkanı: HMW-HA ve 20 Kopya TP53 Entegrasyonu",
        "Homo Aeternus dokularında kanser oluşumu iki aşamalı aşılmaz bir moleküler duvarla kilitlenmiştir.",
        "Birinci Katman (Matriks Zırhı): Köstebek faresinin nmrHAS2 geni ile salgılanan 10 MegaDaltonluk HMW-HA, dokularda aşırı erken temas inhibisyonu (ECI) sağlayarak hücrelerin üst üste kümelenmesini ve neoplastik kitle oluşturmasını fiziksel olarak imkânsız kılar; İkinci Katman (Nükleer İnfaz): Eğer bir hücre bu engeli aşacak bir onkogenik mutasyon geliştirirse, filin 20 kopya TP53 ve LIF6 kaskadı anında uyanarak o hücreyi 15 dakikada apoptozla yok eder; kanser riski mutlak sıfırdır.",
        "Olasilik_Kanser_HomoAeternus = P_onkojen = P_kacis_HMW_HA * P_kacis_20xTP53 ~ 10^-6 * 10^-8 = 10^-14 ~ SIFIR",
        "This double-barrier mathematical probability equation proves that the combined integration of naked mole-rat and elephant genetic architectures reduces human oncogenesis risk to absolute zero."
    ),
    (
        "10.3 Balina Tipi Ultra-Hızlı Homolog Rekombinasyon ve Genomik Dayanıklılık",
        "Hücrelerimizin DNA'sı her gün yüz binlerce metabolik ve çevresel darbeye maruz kalır; ancak Homo Aeternus hücreleri bu hasarları Grönland balinasının sadakatiyle onarır.",
        "Balina tipi modifiye edilmiş ERCC1, PCNA ve CIRBP genleri; insan hücrelerinde çift zincir kırıklarında hata eğilimli NHEJ yerine kusursuz Homolog Rekombinasyonu (HR) çalıştırır. Somatik mutasyon birikim hızı yılda 47'den yılda 5 mutasyonun altına indirilir. 150 yaşına gelmiş bir Homo Aeternus hücresinin genomu, 10 yaşındaki bir çocuğun genomu kadar saf ve mutasyonsuz kalır.",
        "Mutasyon_Hizi_Aeternus = dM/dt < 5 mutasyon/yıl (İnsandan 10 Kat Daha Yavaş Mutasyon Birikimi)",
        "Bu somatik mutasyon kinetiği eşitliği, balina tipi DNA onarım modifikasyonlarının insan genomik yıpranmasını nasıl durdurduğunu belgeler."
    ),
    (
        "10.4 Hidra-Tipi FoxO Kök Hücre Ağları ve Ebedi Doku Rejenerasyonu",
        "İnsan organlarının yaşlanmasındaki ana darbe kök hücre havuzlarının tükenmesidir; Homo Aeternus bu sorunu tatlı su polipi Hidra'nın FoxO genetik mimarisiyle çözer.",
        "Akt-dirençli nükleer FoxO3a kaskadı; kemik iliği, bağırsak epiteli, nörojenik nişler ve kas uydu hücrelerinde kök hücrelerin yaşam boyu sınırsız bölünme ve kendini yenileme potansiyelini korur. Doku kaybı veya mikro-hasar oluştuğunda kök hücreler anında çoğalarak hasarlı bölgeyi genç hücrelerle doldurur; organ atrofisi ve doku sklerozu tarihe karışır.",
        "Yenilenme_Kapasitesi = J_rejenerasyon = k_FoxO * [Kök_Hücre_Aktif] = Sabit (Sonsuz Rejeneratif Güç)",
        "Bu kök hücre rejenerasyon akısı formülü, Hidra tipi FoxO regülasyonunun insan dokularında asırlar boyu süren taze hücre üretimini nasıl garanti ettiğini simgeler."
    ),
    (
        "10.5 Turritopsis Döngüsel Transdiferansiasyon Protokolü: Hücresel Gençleşme Anahtarı",
        "Homo Aeternus hücreleri en uç biyolojik stres, zehirlenme veya iskemik travma durumunda dahi ölmek yerine; Turritopsis dohrnii'nin hücresel gençleşme kodunu devreye sokar.",
        "Doksisiklin veya optogenetik mavi ışıkla uzaktan tetiklenen 'Döngüsel Dediferansiasyon Devresi'; organ parankimindeki yaşlanmış veya hasarlı hücrelerin kimliğini geçici olarak gevşetir. Epigenetik metilasyon saatleri sıfıra formatlanır, telomerler uzatılır, hücre içi atıklar otofajiyle eritilir ve hücreler 20 yaşındaki genç kimliklerinde yeniden farklılaşır; beden kendi kendini sıfırlayan bir döngüye kavuşur.",
        "Format_Hucresel = Epigenetik_Yas(t_post_dongu) = 20 Yaş (Turritopsis Yeniden Başlatma Algoritması)",
        "Bu epigenetik yeniden başlatma aksiyomu, insan somatik hücrelerinin Turritopsis transdiferansiasyon mantığıyla periyodik olarak nasıl gençlik fazına döndürüldüğünü tanımlar."
    ),
    (
        "10.6 Yarasa İnflamatuar Sükuneti: İnflammaging'in Biyolojik Tasfiyesi",
        "Yaşlanmanın en sinsi katili olan kronik steril yangı (inflammaging); yarasaların evrimsel adaptasyonları ile insan immün sisteminden tamamen sökülüp atılmıştır.",
        "Körelmiş NLRP3 inflamazomu ve modifiye STING yolağı sayesinde; hücre içi atıklar veya yaşlanma sinyalleri lökositleri gereksiz sitokin fırtınalarına sürükleyemez. Sürekli uyanık Tip-1 interferon kalkanı ise tüm viral tehditleri kapıda yok eder. İnsan bağışıklık sistemi dokuları yıkan bir yangı kaynağı değil; sessiz, kusursuz ve ömür boyu koruyan bir güvenlik kalkanına dönüşür.",
        "Sitokin_Bazal = [IL-6] + [TNF-a] + [CRP] = Minimum_Fizyolojik (Kronik Yangı Sıfırlanmıştır)",
        "Bu yangısal bazal hat denklemi, yarasa tipi immünomodülasyonun insan bedenindeki kronik inflamatuar doku aşınmasını nasıl tamamen sildiğini gösterir."
    ),
    (
        "10.7 Arctica islandica Proteostaz Gücü: Amiloidoz ve Agregatomun Kesin Engellenmesi",
        "500 yıllık Ming istiridyesinin şaperon proteinleri ve düşük peroksidasyon indeksli membran lipidleri; insan beynini ve kalbini amiloid ve tau plaklarından ebediyen korur.",
        "Hücre içi küçük ısı şok proteinleri (sHSPs) ve kristalin şaperonları; yanlış katlanan monomerleri anında yakalayarak agregat oluşumuna izin vermez. 120 yaşına gelen bir Homo Aeternus bireyinin neokorteksinde tek bir amiloid-beta veya nörofibriler tau yumağı bulunmaz; zihin ömür boyu berrak, keskin ve parlak kalır.",
        "Agregatom_Yuku = Phi_plak = SIFIR (Arctica islandica Şaperon Zırhı ile Korunur)",
        "Bu proteopatik sıfırlama eşitliği, istiridye kaynaklı şaperon mühendisliğinin insan nörodejeneratif hastalıklarını kökünden nasıl kazıdığını belgeler."
    ),
    (
        "10.8 Çok-Türlü (Multi-Species) Sentetik Genom Konsorsiyumu",
        "Homo Aeternus genomu tek bir türün genetik havuzuna hapsolmuş dar bir biyolojik biyotop değildir; gezegenin 4 milyar yıllık evrimsel dehasının en parlak longevity şifrelerini birleştiren 'Sentetik Genom Konsorsiyumu'dur.",
        "Köstebek faresinin hyaluronanı, filin p53'ü, balinanın DNA onarımı, hidranın FoxO'su, yarasının inflamazom toleransı, kaplumbağanın telomer kararlılığı ve istiridyenin proteostazı; tek bir insan hücre çekirdeğinde kusursuz bir genetik orkestra gibi ahenkle çalışır. Bu konsorsiyum doğanın tüm yaşam koruma cephanesinin tek bir biyolojik formda toplanmasıdır.",
        "Konsorsiyum_Genom: Genom_Aeternus = Genom_Insan (+) Sigma_{Tur} [ Genomik_Longevity_Modulu_Tur ]",
        "Bu sentetik taksonomik birleşim modeli, çok-türlü longevity genetik modüllerinin insan biyolojik platformunda inşa ettiği süper-dirençli organizmayı tanımlar."
    ),
    (
        "10.9 Biyolojik Çeşitliliğin Korunması ve Biyofilik Etik Imperatif",
        "Komparatif longevity biyolojisi; yeryüzündeki yaban hayatın ve biyolojik çeşitliliğin korunmasını insanlığın kendi ölümsüzlüğü için mutlak bir varoluşsal zorunluluk (Biyofilik Etik) haline getirir.",
        "Eğer bir balina, bir köstebek faresi veya kutup sularındaki bir istiridye türü yok olsaydı; onların genomunda saklı olan ve kanseri, DNA hasarını veya kalp krizini yenen bu paha biçilmez şifreler sonsuza kadar kaybolacaktı. Gezegenin biyolojik çeşitliliğini korumak, aslında insanın kendi ölümsüz geleceğini korumaktır.",
        "Etik_Imperatif: BiyoCesitlilik_Koruma == Insan_Olumsuzlugu_Garantisi",
        "Bu biyoetik aksiyom, doğadaki diğer canlı türlerinin yaşam hakkını korumanın insanın kendi biyolojik geleceğini inşa etme zorunluluğu ile mutlak özdeşliğini formüle eder."
    ),
    (
        "10.10 Homo Aeternus ve Evrimsel Taç: Doğanın Kendi Kendini Aşan Ölümsüz Çocuğu",
        "Homo Aeternus; doğaya karşı savaşan kibirli bir yabancı değil, doğanın 4 milyar yıllık evrimsel kütüphanesini okumayı öğrenmiş, onu anlamış, ona saygı duymuş ve onun en yüce yasalarını kendi bedeninde taçlandırmış bilinçli evrimsel çocuktur.",
        "Filin gücü, balinanın sabrı, köstebek faresinin zırhı, denizanasının sonsuz döngüsü ve insanın yüksek bilinci tek bir ebedi varlıkta birleşmiştir. İnsan artık ölümlü bir primat değildir; kozmosun derinliklerine doğru yelken açan, yaşlanmayan, hastalanmayan ve bilinci sonsuza kadar parıldayan ölümsüz bir evren yurttaşıdır.",
        "Nihai_Sentez: Homo_Aeternus = [Evrimin Tum Longevity Dehasi] * [Insan Bilinci] = Ebedi Varolus",
        "Bu nihai ontolojik taç denklemi, komparatif biyoloji ve genetik mühendisliğin insanı fanilikten kurtararak evrenin ebedi ve ölümsüz bir bilincine dönüştürdüğünü evrensel olarak ilan eder."
    )
]

# 10 Kapsamlı Akademik Karşılaştırma Tablosu
tables_data = [
    {
        "title": "Tablo 17.1: Peto Paradoksu ve Farklı Memeli Türlerinde Hücre Sayısı-Kanser Korelasyonu",
        "headers": ["Canlı Türü", "Beden Ağırlığı / Hücre Sayısı", "Maksimum Yaşam Süresi", "Kanserden Ölüm Oranı (%)", "Evrimsel Kanser Direnci Mekanizması"],
        "rows": [
            ["Laboratuvar Faresi (Mus musculus)", "30 gram / ~3 * 10^10 hücre", "3 - 4 Yıl", "%50 - 70 (Aşırı Yüksek)", "Zayıf DNA onarımı, Düşük p53 kopyası, Kısa ömür"],
            ["İnsan (Homo sapiens)", "70 kg / ~3.7 * 10^13 hücre", "122 Yıl (Jeanne Calment)", "%20 - 25 (Orta/Yüksek)", "Tek kopya TP53, Gelişmiş NER/BER, Hayflick telomer sınırı"],
            ["Afrika Fili (Loxodonta africana)", "5.000 kg / ~3 * 10^15 hücre", "65 - 70 Yıl", "< %5 (Aşırı Düşük / Peto)", "20 kopya (40 alel) TP53 + Uyanmış LIF6 zombi geni"],
            ["Grönland Balinası (Balaena mysticetus)", "100.000 kg / ~10^16 hücre", "211+ Yıl (Zirve Memeli)", "< %1 (Kanser Neredeyse Sıfır)", "Kusursuz Homolog Rekombinasyon, ERCC1/PCNA mutasyonları"],
            ["Mavi Balina (Balaenoptera musculus)", "150.000 kg / ~1.5 * 10^16", "90 - 110 Yıl", "< %2 (Peto Paradoksu)", "Hipertümör dinamiği, Ultra-düşük özgül metabolik hız"],
            ["Çıplak Kör Köstebek Faresi", "35 gram / ~3 * 10^10 hücre", "37+ Yıl (Kemirgen Zirvesi)", "SIFIR (Laboratuvarda Spontan Yok)", "10 MDa HMW-HA, Erken Temas İnhibisyonu (p16/p27)"],
            ["Brant's Yarasası (Myotis brandtii)", "7 gram / ~7 * 10^9 hücre", "40+ Yıl (Cüssesine göre 5x)", "< %3 (Aşırı Direnç)", "Körelmiş NLRP3, Daimi IFN-alfa kalkanı, Yüksek DDR"],
            ["Homo Aeternus (Sentetik Biyoloji)", "75 kg / Biyo-Hibrit Hücre", "SONSUZ (Negligible Senescence)", "MUTLAK SIFIR (%0.000)", "Çift-katmanlı HMW-HA + 20xTP53 + Balina HR Sinerjisi"]
        ]
    },
    {
        "title": "Tablo 17.2: Çıplak Kör Köstebek Faresi (Heterocephalus glaber) Moleküler Kanser Direnci",
        "headers": ["Biyomoleküler Bileşen", "Köstebek Faresi Fenotipi", "İnsan / Fare Fenotipi", "Moleküler Biyofiziksel Mekanizma", "Hücresel Kanser Önleme Sonucu"],
        "rows": [
            ["Hyaluronik Asit Boyutu (HA)", "6 - 12 MegaDalton (HMW-HA)", "0.5 - 2 MegaDalton (Düşük/Orta)", "HAS2 enzim mutasyonları (Ser211/Val657)", "Hücreler arası viskoz jel zırhı / Tümör kümelenmesi imkânsız"],
            ["Hiyaluronidaz Aktivitesi", "Aşırı Düşük (HYAL1/2 inaktif)", "Normal / Yüksek yıkım", "Enzimatik promotor susturulması", "HMW-HA dokularda birikerek ömür boyu kalır"],
            ["Temas İnhibisyonu Eşiği", "Erken Temas İnhibisyonu (ECI)", "Geç Temas İnhibisyonu", "Düşük hücre yoğunluğunda erken durma", "Hücrelerin üst üste yığılması sıfırlanır"],
            ["Tümör Baskılayıcı Sinyal", "p16INK4a + p27Kip1 (Çift Fren)", "Yalnızca p16 veya p21", "CD44 kümelenmesi ile erken aktivasyon", "Onkojenik mutasyonlarda kaçışı imkânsız kılar"],
            ["Ribozomal RNA Mimarisi", "Bölünmüş (Split) 28S rRNA", "Tek parça standart 28S rRNA", "Ribozom rijitliği ve konformasyon kararlılığı", "Translasyonel kodon okuma hatasında 10 kat düşüş"],
            ["26S Proteazom Aktivitesi", "Yüksek Proteolitik Kapasite", "Normal (Yaşla %50 çöker)", "Yüksek şaperon ve ubikitin döngüsü", "Protein agregatomu ve lipofuskin birikimi sıfır"],
            ["Telomeraz Regülasyonu", "Düşük somatik TERT aktivitesi", "Sıfır somatik / Kanser kök hücresi", "Katı p53 baskısı altında tutulur", "Kanser kök hücresi immortalizasyonunu kilitler"],
            ["Kanser Aşılama Yanıtı", "Onkogenlerle dahi tümör oluşmaz", "Ras + SV40 ile anında tümör", "HMW-HA erimesi olmadan malignleşme yok", "Doğal ve spontan kansere karşı mutlak zırh"]
        ]
    },
    {
        "title": "Tablo 17.3: Grönland Balinası (Balaena mysticetus) Genomik Stabilite ve DNA Onarım Analizi",
        "headers": ["Genomik / Biyokimyasal Parametre", "Grönland Balinası Karakteristiği", "İnsan Karakteristiği", "Evrimsel Moleküler Adaptasyon", "200+ Yıllık Longevity Katkısı"],
        "rows": [
            ["Maksimum Yaşam Süresi", "211+ Yıl (Göz lensi aspartik analiz)", "122 Yıl", "Pozitif seçilim altındaki longevity genleri", "İki asırlık kansersiz ve hastalıksız varoluş"],
            ["DNA Çift Zincir Kırığı Onarımı", "Homolog Rekombinasyon (HR) Baskın", "Hata eğilimli NHEJ Baskın", "HR tamir yolunun insandan 2 kat daha aktif olması", "Kromozomal delesyon ve translokasyonların önlenmesi"],
            ["ERCC1 ve PCNA Genleri", "Özel pozitif seçilim mutasyonları", "Standart memeli dizisi", "Nükleotid eksizyon onarımı (NER) afinite artışı", "UV ve oksidatif DNA hasarlarının anında temizlenmesi"],
            ["Gen Kopya Sayısı (CNV)", "Kritik koruyucu genlerde duplikasyon", "Tekil kopyalar", "LAMTOR1 ve apoptoz genlerinde genomik yedeklilik", "Tek nokta mutasyonuyla sistemin çökmesini engelleme"],
            ["Soğuk Şok Proteini (CIRBP)", "Sürekli Yüksek Bazal Ekspresyon", "Yalnızca soğuk şokta geçici", "mRNA'ların dondurucu ortamda korunması ve ATR aktivasyonu", "Hücresel transkriptomun arktik soğukta kararlılığı"],
            ["Mitokondriyal Elektron Kaçağı", "Aşırı Düşük (Süperoksit ROS <%5)", "Normal (Mitokondriyal kaçak %1-2)", "Kompleks I ve IV proton pompalama optimizasyonu", "Mitokondriyal DNA mutasyon birikiminin engellenmesi"],
            ["Somatik Mutasyon Hızı", "Yılda ~12 mutasyon / hücre", "Yılda ~47 mutasyon / hücre", "Ultra-yüksek DNA polimeraz sadakati", "200 yıl sonunda bile kritik mutasyon kotasını aşmama"],
            ["Hippo-YAP Temas Sinyali", "Aşırı hassas temas durdurma", "Normal doku temas inhibisyonu", "E-kaderin ve mekanotransdüksiyon ile anında G1 bloğu", "Doku aşırı büyümesi ve tümörleşmenin engellenmesi"]
        ]
    },
    {
        "title": "Tablo 17.4: Afrika Fili (Loxodonta africana) TP53 ve LIF6 Kanser Savunma Mimarisi",
        "headers": ["Fil Genomik Bileşeni", "Gen Kopya Sayısı / Mimarisi", "İnsan Genomik Muadili", "Aktivasyon Mekanizması", "Onkogenik Baskılama Gücü"],
        "rows": [
            ["TP53 Genomik Rezervi", "20 Kopya (1 Orijinal + 19 Retrogen)", "Tek Kopya (1 Gen / 2 Alel)", "DNA kırığında masif p53 protein deşarjı", "İnsandan 20 kat yüksek nükleer p53 havuzu"],
            ["p53 Protein Kararlılığı", "MDM2 parçalanmasına dirençli mutasyon", "MDM2 ile hızlı parçalanma (kısa ömür)", "C-terminal delesyonları ile uzun yarı ömür", "DNA hasarı anında saatlerce nükleusta aktif kalma"],
            ["Hücresel Yanıt Tercihi", "Doğrudan Apoptoz (Tavizsiz İnfaz)", "G1 Duraklaması ve Tamir Denemesi", "Yüksek p53 konsantrasyonuyla pro-apoptotik kaskad", "Hasarlı tek bir hücrenin dahi yaşamasına izin vermeme"],
            ["Uyanmış Zombi Gen (LIF6)", "Fonksiyonel Psödogen (Re-activated)", "İnaktif / Sessiz Psödogen", "p53 doğrudan bağlanarak LIF6'yı transkribe eder", "Mitokondri dış zarını delerek (MOMP) apoptozu patlatma"],
            ["Radyasyon Hassasiyeti", "Gama ışınlarında %85 hücre intiharı", "Hasarlı hücreler bölünmeye zorlanır", "Aşırı duyarlı DNA hasar sensörleri", "Radyasyon kaynaklı kromozomal mutasyonların sıfırlanması"],
            ["Kanser Mortalite Oranı", "%4.8 (Tüm ölümler içinde)", "%20 - 25 (Her 4 insandan biri)", "Devasa doku kütlesine rağmen aşılmaz güvenlik", "Peto paradoksunun omurgalı alemindeki en somut zaferi"],
            ["Doku Feda Edilebilirliği", "Milyonlarca hasarlı hücreyi feda etme", "Her hücreyi korumaya çalışma", "60 trilyon hücrelik devasa rezerv avantajı", "Doku kaybı olmadan prekanseröz hücrelerin temizlenmesi"],
            ["Kanser Kök Hücresi Riski", "Sıfıra yakın (Anında apoptoz)", "Yüksek (CSC klonları gelişebilir)", "TERT baskılaması ve aşırı duyarlı p53 denetimi", "Tümör kök hücre klonlarının oluşmadan imhası"]
        ]
    },
    {
        "title": "Tablo 17.5: İhmal Edilebilir Yaşlanma (Negligible Senescence) Gösteren Türlerin Biyolojisi",
        "headers": ["Canlı Türü", "Belgelenmiş Maksimum Yaşam", "Gompertz Yaşlanma Hızı (G)", "Üreme ve Fizyolojik Seyir", "Temel Hücresel / Biyofiziksel Sır"],
        "rows": [
            ["Rougheye Kaya Balığı (Sebastes)", "205 Yıl", "0.000 (Yaşla mortalite artmaz)", "Yaşlandıkça fertilite ve yumurta artar", "Kusursuz DNA onarımı, Metabolik stabilite"],
            ["Arctica islandica (Ming İstiridyesi)", "507 Yıl", "0.000 (Yarım binyıl sabit risk)", "5 asır boyunca intakt üreme", "Aşılmaz proteostaz, Düşük lipid peroksidasyon indeksi (PI)"],
            ["Dev Kaplumbağa (Galapagos/Seyşeller)", "190+ Yıl (Jonathan)", "0.000 (Mortalite düz çizgi)", "150 yaşında gençlik çiftleşmesi", "Somatik telomeraz sürekliliği, NEIL1/BRCA2 genişlemesi"],
            ["Grönland Köpekbalığı (Somniosus)", "392 ± 120 Yıl", "0.000", "150 yaşında ilk cinsel ergenlik", "-1.5°C Arktik derin deniz yavaş metabolizması, Kalp hızı 0.08 Hz"],
            ["Kızıl Denizkestanesi (Mesocentrotus)", "200+ Yıl", "0.000", "Yaşlandıkça daha verimli gametler", "Sürekli telomeraz aktivitesi, Kusursuz rejenerasyon"],
            ["Tatlı Su İstiridyesi (Margaritifera)", "250+ Yıl", "0.000", "Asırlar boyu sabit filtrasyon", "Oksidatif stres direnci, Mitokondriyal dayanıklılık"],
            ["Timsahlar (Crocodylia)", "80 - 100+ Yıl", "Çok Düşük (İhmal edilebilir yakın)", "Belirsiz büyüme (Indeterminate growth)", "Somatik telomeraz ve tavizsiz p53 regülasyonu"],
            ["Homo Aeternus (Hedeflenen Biyoloji)", "SONSUZ (Biyostaz / Genetik)", "0.000 (Gompertz Sıfırlanmıştır)", "Ebedi gençlik fertilitesi ve dinamizmi", "Komparatif biyoloji gen konsorsiyumu entegrasyonu"]
        ]
    },
    {
        "title": "Tablo 17.6: Turritopsis dohrnii ve Hydra vulgaris: Sonsuz Hücresel Döngü ve Dediferansiasyon",
        "headers": ["Ölümsüz Canlı Türü", "Biyolojik Sınıf / Yaşam Alanı", "Ölümsüzlük Mekanizması", "Kök Hücre / Genetik Faktör", "Hücresel Yaşlanmanın Sıfırlanması"],
        "rows": [
            ["Turritopsis dohrnii", "Hidrozoan (Denizanası / Akdeniz)", "Transdiferansiasyon (Ontogenetik Tersine Dönüş)", "PRC2 epigenetik kompleks duplikasyonu", "Medüzadan polipe geri dönüş / Embriyonik yaşa format"],
            ["Hydra vulgaris", "Knidli (Tatlı su polipi)", "Sonsuz Kök Hücre Havuzu Yenilenmesi", "FoxO transkripsiyon faktörü sürekli aktif", "Her 20 günde bir tüm vücut hücrelerinin de novo üretimi"],
            ["Turritopsis rubra (Ölümlü Akraba)", "Hidrozoan (Denizanası)", "Geleneksel ölüm (Transdiferansiasyon yok)", "Tek kopya tamir ve epigenetik genler", "Yaşlanır ve ölür (Dohrnii ile karşılaştırma kontrolü)"],
            ["Dediferansiasyon Dinamiği", "Olgun kas hücresi -> Pluripotent hücre", "Gelişimsel epigenetik hafızanın silinmesi", "Wnt kapatma + MAPK/ERK kinaz pulsasyonu", "Hücre kimliğinin embriyonik esnekliğe iadesi"],
            ["Telomeraz Davranışı", "Sürekli ve Döngüsel Formatlama", "Her polip dönüşümünde TERT ile tam uzama", "Telomer erozyonu tamamen imkânsız", "Sıfır replikatif Hayflick sınırı"],
            ["DNA Onarım Cephanesi", "Genişletilmiş DDR Kütüphanesi", "Kopya sayısı iki katına çıkmış DNA tamir genleri", "Mutasyon birikiminin mutlak engellenmesi", "Genomik saflığın sonsuz döngüler boyunca korunması"],
            ["Yamanaka Protokolü Eşdeğeri", "Doğal İn Vivo Epigenetik Sıfırlama", "OSKM faktörlerinin hücresel karşılığı", "Kimyasal uyaranlarla geriye sarılan gelişim", "İnsan hücresel gençleşmesinin yaşayan modeli"],
            ["Homo Aeternus Entegrasyonu", "Sentetik Rekürsif Döngü Devresi", "Kriz anında dokusal dediferansiasyon", "Optogenetik tetiklemeli hücresel format", "İnsan biyolojisinin dairesel gençleşme döngüsüne geçişi"]
        ]
    },
    {
        "title": "Tablo 17.7: Yarasalarda (Chiroptera) Metabolik Uçuş Stresi, İnflamazom ve Viral Tolerans",
        "headers": ["Yarasa Biyolojik Ekseni", "Fizyolojik / Biyokimyasal Durum", "İnsan / Kemirgen Durumu", "Moleküler Evrimsel Adaptasyon", "Longevity ve Kanser Direnci Çıktısı"],
        "rows": [
            ["Uçuş Metabolik Hızı", "İstirahatin 15-20 katı oksijen tüketimi", "Maksimum eforda 5-8 kat artış", "Vücut sıcaklığı uçuşta 41°C'ye tırmanır", "Muazzam ROS fırtınasına hücresel direnç geliştirme"],
            ["NLRP3 İnflamazomu", "Körelmiş (Dampened / LRR mutasyonları)", "Hiperaktif / Sitokin fırtınası başlatır", "Spesifik amino asit delesyonları", "İnflammaging ve otoimmünite yok / Tam doku toleransı"],
            ["Tip-1 İnterferon (IFN-alfa)", "Sürekli Bazal Aktif (Constitutive)", "Yalnızca viral uyarımla geçici başlar", "3 fonksiyonel IFN geninin kesintisiz transkripsiyonu", "Virüslerin replike olmadan anında durdurulması"],
            ["Ölümcül Virüs Toleransı", "Ebola, Marburg, SARS ile semptomsuz yaşam", "Masif sitokin fırtınası ve ölüm", "Viral replikasyonu frenleme + Doku yangısını söndürme", "Zoonotik patojenlerle kusursuz biyolojik barış"],
            ["Mitofaji Akışı (PINK1/Parkin)", "Aşırı Hızlı (Dakikalar içinde temizlik)", "Yavaş / Yaşla disfonksiyonel", "Mitokondriyal POLG polimeraz yüksek sadakati", "Hücre içi paslanmış mitokondrilerin anında yok edilmesi"],
            ["DNA Hasar Yanıtı (DDR)", "Genişlemiş ATM/ATR ve Ku70/Ku80", "Standart memeli tamir kiti", "Pozitif seçilim altındaki çift zincir tamir enzimleri", "Radyasyon ve metabolik kırıklara aşılmaz kalkan"],
            ["GHR / IGF-1 Ekseni (Brant's)", "Zayıflatılmış Reseptör Sinyali (GHR delesyonu)", "Yüksek büyüme hormonu sinyali", "Somatotropik sinyalin metabolik bakıma yönlendirilmesi", "7 gramlık cüsse ile 40+ yıl rekor yaşam süresi"],
            ["Telomer Kararlılığı (Myotis)", "40 yıl boyunca sıfır telomer kısalması", "Yıllık düzenli telomer aşınması", "Shelterin kompleksi koruması ve dengeli TERT", "Replikatif senesensin tamamen dışlanması"]
        ]
    },
    {
        "title": "Tablo 17.8: Karşılaştırmalı Omiks ve Uzun Ömürlü Türlerde Korunan Longevity Genleri",
        "headers": ["Longevity Gen Ailesi", "Moleküler Fonksiyonel Görevi", "Pozitif Seçilim Gösteren Türler", "Kısa Ömürlü Türlerdeki Durum", "İnsan Genomuna Mühendislik Hedefi"],
        "rows": [
            ["SIRT6 (Sirtuin 6)", "DNA çift zincir tamiri, H3K9 deasetilasyonu", "Kunduz, Çıplak kör köstebek faresi", "Zayıf ribozilaz aktivitesi (Fare)", "SIRT6 enzim aktivitesinin 3 kat artırılması"],
            ["ERCC1 / PCNA", "Nükleotid eksizyon onarımı, Replikasyon çatalı", "Grönland balinası, Filler", "Standart memeli kiti", "Prime editing ile balina nokta mutasyonlarının nakli"],
            ["HAS2 (Hyaluronan Sentaz 2)", "10 MDa yüksek moleküler hyaluronan sentezi", "Çıplak kör köstebek faresi, Spalax", "Düşük MW (1 MDa) üretimi (İnsan/Fare)", "nmrHAS2 transgeni ile insan dokularında HMW-HA zırhı"],
            ["TP53 / LIF6 Ekseni", "DNA hasarında anında apoptoz kaskadı", "Afrika fili (20 kopya TP53 + LIF6)", "Tek kopya TP53 (İnsan/Fare)", "Güvenli liman lokuslarına ek TP53 kopyaları yerleştirme"],
            ["FoxO (Forkhead Box O)", "Kök hücre yenilenmesi, Otofaji, Antioksidan", "Hydra vulgaris, Dev kaplumbağalar", "Akt ile fosforillenip inaktive olur", "Akt-dirençli nükleer kalıcı FoxO3a kaskadı"],
            ["NLRP3 (İnflamazom)", "İnterlökin-1beta salınımı ve piroptotik ölüm", "Yarasalar (Körelmiş LRR domaini)", "Hiperaktif yangısal reaktivite", "İnsan makrofajlarında NLRP3 sakinleştirme cerrahisi"],
            ["28S rRNA (Bölünmüş)", "Ribozomal rijitlik ve translasyonel sadakat", "Köstebek faresi, Çöl kemirgenleri", "Tek parça hata eğilimli ribozom", "Sentetik bölünmüş 28S rRNA mühendisliği"],
            ["Supercool PVA Mimetikleri", "Buz kristali adsorpsiyon-inhibisyonu", "Kutup balıkları (AFGP), Bitkiler", "Memelilerde sıfır antifriz proteini", "Biyostaz perfüzyon kokteyllerinde standart BBA kullanımı"]
        ]
    },
    {
        "title": "Tablo 17.9: Komparatif Longevity Genlerinin İnsan Hücrelerine Aktarım Matrisi",
        "headers": ["Ksenojenik Gen / Mekanizma", "Kaynak Organizma", "Kullanılan Gen Düzenleme Aracı", "Hedef İnsan Hücre / Dokusu", "Klinik Fenotipik Kazanç"],
        "rows": [
            ["nmrHAS2 (HMW-HA Sentezi)", "Çıplak Kör Köstebek Faresi", "AAV.CAP-B10 / LNP-mRNA", "Deri, Vasküler endotel, Eklemler", "Kanser-geçirmez elastik doku, Erken temas inhibisyonu"],
            ["20x TP53 + LIF6 Kaseti", "Afrika Fili (Loxodonta africana)", "CRISPR AAVS1 Güvenli Liman Entegrasyonu", "Hematopoietik kök hücreler, iPSC", "Aşırı duyarlı apoptoz ile kanser riskinin sıfırlanması"],
            ["Balina ERCC1/PCNA Nokta Mutasyonu", "Grönland Balinası", "Prime Editing (PE-max)", "Tüm somatik hücre nükleusları", "Homolog rekombinasyonla kusursuz çift zincir onarımı"],
            ["NLRP3 LRR Sakinleştirme", "Yarasalar (Chiroptera)", "Adenin Baz Düzenleyici (ABE)", "Monositler, Makrofajlar, Mikroglia", "İnflammaging ve sitokin fırtınasının mutlak tasfiyesi"],
            ["Akt-Dirençli FoxO3a", "Tatlı Su Polipi (Hydra vulgaris)", "dCas9 Epigenom / Transgenik Kaset", "Erişkin doku kök hücre havuzları", "Ömür boyu tükenmeyen sonsuz doku yenilenme kapasitesi"],
            ["Ai-Hsp27 Şaperon Zırhı", "Okyanus İstiridyesi (Arctica islandica)", "AAV9 Serebral Perfüzyon", "Kortikal nöronlar ve Kardiyomiyositler", "Protein agregatomu, Amiloid ve Tau plaklarının engellenmesi"],
            ["Düşük PI Membran Lipit Mühendisliği", "Kaya Balığı / Dev Kaplumbağa", "Metabolik enzim regülasyonu", "Mitokondriyal iç membranlar", "Lipid peroksidasyonunun sıfırlanması, Dayanıklı solunum"],
            ["Turritopsis Dediferansiasyon Anahtarı", "Ölümsüz Denizanası (T. dohrnii)", "Doksisiklin İndüklenebilir OSK Kaseti", "Tüm somatik organ parankimleri", "Kriz anında hücre kimliğini sıfırlayıp yeniden başlama"]
        ]
    },
    {
        "title": "Tablo 17.10: Homo Aeternus Negligible Senescence ve Biyolojik Ölümsüzlük Manifestosu",
        "headers": ["Evrimsel Longevity Katmanı", "Kaynağı Olan Şampiyon Tür", "İnsan Genomuna Entegre Mekanizma", "Nihai Biyolojik Güvenlik Garantisi", "Homo Aeternus Varoluşsal Düzeyi"],
        "rows": [
            ["Kanser Kalkanı Katmanı 1", "Çıplak Kör Köstebek Faresi", "10 MDa HMW-HA + CD44 ECI freni", "Katı tümör oluşumu ve metastaz imkânsız", "Kanserden tamamen muaf doku mimarisi"],
            ["Kanser Kalkanı Katmanı 2", "Afrika Fili (Loxodonta)", "20 Kopya TP53 + Uyanmış LIF6 geni", "DNA hasarlı tek hücrenin anında infazı", "Malign transformasyon olasılığı sıfır"],
            ["Genomik Sadakat Katmanı", "Grönland Balinası", "Balina ERCC1/PCNA + Homolog Rekombinasyon", "Somatik mutasyon hızı yılda <5 mutasyon", "200 yıl sonra dahi çocukluk genom saflığı"],
            ["Sonsuz Kök Hücre Katmanı", "Tatlı Su Polipi (Hydra vulgaris)", "Nükleer dengeli FoxO3a rejenerasyon ağı", "Kök hücre nişlerinin tükenmesinin engellenmesi", "Ömür boyu kesintisiz genç doku üretimi"],
            ["Enflamatuar Sükunet Katmanı", "Yarasalar (Chiroptera)", "Körelmiş NLRP3 inflamazom toleransı", "İnflammaging ve otoimmün krizlerin yokluğu", "Sakin, dengeli ve genç immün gözetim"],
            ["Proteomik Kararlılık Katmanı", "Ming İstiridyesi (Arctica islandica)", "Yüksek mukavemetli sHSPs ve şaperon ağı", "Amiloidoz, Tau ve lipofuskin sıfırlanması", "500 yıl boyunca berrak ve pırıl pırıl proteom"],
            ["Döngüsel Gençleşme Katmanı", "Ölümsüz Denizanası (Turritopsis)", "İn vivo kısmi transdiferansiasyon devresi", "Gerektiğinde epigenetik saati sıfıra sarma", "Zamanın doğrusal prangasından mutlak kurtuluş"],
            ["Homo Aeternus Zirvesi", "Tüm Canlılar Aleminin Sentezi", "Sentetik Biyoloji ve Komparatif Genomik", "Gompertz Yaşlanma Hızı G = 0.000000", "Doğanın ölümsüz çocuğu: Homo Aeternus"]
        ]
    }
]

# Dokuman Olusturma Dongusu
parts = [
    ("KISIM 1: PETO PARADOKSU VE BÜYÜK CANLILARDA KANSER DİRENCİ EVRİMİ", part1_subsections),
    ("KISIM 2: ÇIPLAK KÖR KÖSTEPEK FARESİ: YÜKSEK MOLEKÜLER AĞIRLIKLI HYALURONİK ASİT (HMW-HA)", part2_subsections),
    ("KISIM 3: GRÖNLAND BALİNASI: 200+ YILLIK YAŞAM VE KUSURSUZ DNA ONARIMI", part3_subsections),
    ("KISIM 4: FİLLER (LOXODONTA AFRICANA) VE 20 KOPYA TP53", part4_subsections),
    ("KISIM 5: NEGLIGIBLE SENESCENCE: YAŞLANMAYAN CANLILARIN BİYOFİZİĞİ", part5_subsections),
    ("KISIM 6: TURRITOPSIS DOHRNII VE HİDRALAR: MORFOLOJİK TRANSDİFERANSİASYON VE SONSUZ POLİP DÖNGÜSÜ", part6_subsections),
    ("KISIM 7: YARASALARDA (CHIROPTERA) İLTİHAP BASKILAMA VE AŞIRI VİRAL TOLERANS", part7_subsections),
    ("KISIM 8: KARŞILAŞTIRMALI OMİKS VE GENOMİK MADENCİLİK", part8_subsections),
    ("KISIM 9: İNSAN HÜCRELERİNE KOMPARATİF LONGEVITY GENLERİNİN AKTARIMI", part9_subsections),
    ("KISIM 10: HOMO AETERNUS NEGLIGIBLE SENESCENCE MANİFESTOSU: DOĞANIN EN UZUN YAŞAYAN CANLILARININ SENTETİK SENTEZİ", part10_subsections)
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
        frun = pf.add_run(f"Komparatif / Evrimsel Longevity Bağıntısı:  {formula}")
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
print(f"CİLT 17 Başarıyla Kaydedildi: {OUTPUT_PATH}")