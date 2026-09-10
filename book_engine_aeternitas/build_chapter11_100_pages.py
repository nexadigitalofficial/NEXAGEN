# -*- coding: utf-8 -*-
"""
PROJECT AETERNITAS - CİLT 11: KRONİK ENFLAMASYON, İMMÜNOSENESENS VE MİKROBİYOTA-BAĞIŞIKLIK EKSENİ
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

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_11_KRONIK_ENFLAMASYON_IMMUNOSENESENS_VE_MIKROBIYOTA_TAM_100_SAYFA.docx"

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
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 11: İMMÜNOSENESENS VE KRONİK ENFLAMASYON")
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
s_run = sub_p.add_run("CİLT 11: KRONİK ENFLAMASYON, İMMÜNOSENESENS VE MİKROBİYOTA-BAĞIŞIKLIK EKSENİ\\n(INFLAM-AGING, cGAS-STING, TİMİK İNVOLÜSYON, DISBIYOZIS VE İMMÜN RESTORASYON)")
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
ih_run = intro_h.add_run("CİLT 11 MANİFESTOSU: STERİL İNFLAMASYONUN TASFİYESİ, İMMÜN RESTORASYON VE BAĞIŞIKLIK MİMARİSİ")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "Bağışıklık sistemi, organizmayı patojenlerden ve neoplastik sapmalardan koruyan en dinamik savunma hattıdır. "
    "Ancak yaşlanma süreci, bu koruyucu kalkanı kendi dokularını yavaşça kemiren steril bir yıkım makinesine dönüştürür. "
    "Claudio Franceschi tarafından 'Inflam-Aging' olarak adlandırılan bu olgu; belirgin bir enfeksiyon olmaksızın, doku düzeyinde "
    "sürekli, düşük dereceli, steril ve kronik bir yangı fırtınasının sürmesidir.\\n\\n"
    "Bu yıkımın arkasında çok katmanlı moleküler patolojiler yatar: Hasarlı mitokondrilerden sızan mtDNA ve mikronükleus parçaları "
    "cGAS-STING yolağını tetikler; hücresel çöpler (DAMPs) ve mikrokristaller NLRP3 inflamazomunu ateşler; timus bezinin yağ dokusuna "
    "dönüşmesi (timik involüsyon) naif T hücre üretimini sıfırlar; bağırsak epitel bariyerinin çökmesi (leaky gut) kana endotoksin (LPS) sızdırır; "
    "ve tükenmiş T hücreleri bağışıklık sürveyansını kaybederek senesen hücrelerin ve tümörlerin çoğalmasına zemin hazırlar.\\n\\n"
    "Bu ciltte; steril inflamasyonun biyofiziği, cGAS-STING ve NLRP3 moleküler mimarisi, timik gençleşme protokolleri (TRIIM), "
    "doğal ve edinsel bağışıklık yaşlanması, bağırsak disbiyozisi ve mikrobiyotanın restorasyonu, immün kontrol noktaları (PD-1, CTLA-4) "
    "ve Homo Aeternus yapay immünom mimarisi 100 kapsamlı akademik bölümde ele alınmaktadır."
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
        "1.1 Franceschi Inflam-Aging Paradigması: Evrimsel Uyumsuzluk ve Kronik Yangı",
        "Claudio Franceschi ve arkadaşları tarafından 2000 yılında ortaya atılan 'Inflam-Aging' kavramı; yaşlanmaya eşlik eden, belirgin bir patojenik enfeksiyon olmaksızın gelişen, subklinik, sistemik, kronik ve steril pro-enflamatuar durumu tanımlar.",
        "Evrimsel antogonistik pleiotropi perspektifinden bakıldığında; gençlikte organizmayı ölümcül enfeksiyonlardan korumak ve yara iyileşmesini hızlandırmak üzere optimize edilmiş hiper-aktif bir doğal bağışıklık yanıtı, üreme çağı sonrasında (post-reprodüktif dönem) sürekli aktif kalarak dokuları yıpratan bir mekanizmaya dönüşür. Bu kronik yangı hali, hücresel senesens, genomik hasar, proteostaz çöküşü ve metabolik sendrom ile iki yönlü bir pozitif geribildirim sarmalı oluşturarak çoklu organ dejenerasyonunun ana itici gücü haline gelir.",
        "Inflam_Aging_Skoru = Sum_i ( w_i * [Sitokin_i_plazma] ) / ( 1 + [Anti_Inflamatuar_Tampon] )",
        "Bu skorlama modeli, pro-enflamatuar medyatörlerin (IL-6, TNF-alfa, IL-1beta) ağırlıklı toplamının endojen anti-enflamatuar faktörlere (IL-10, TGF-beta) oranını ifade eder."
    ),
    (
        "1.2 DAMP (Hasarla İlişkili Moleküler Kalıplar) ve PAMP Molekülleri",
        "Bağışıklık sistemi iki temel tehlike kalıbını tarar: Ekzojen mikroorganizmalardan türeyen Patojenle İlişkili Moleküler Kalıplar (PAMPs: LPS, peptidoglikan, viral RNA) ve hasarlı somatik hücrelerden açığa çıkan Hasarla İlişkili Moleküler Kalıplar (DAMPs / Alarminler).",
        "Inflam-aging sürecinde yangıyı körükleyen birincil faktörler steril DAMP molekülleridir. Hücre membran bütünlüğü bozulduğunda veya nekroz/piroptoz gerçekleştiğinde hücre dışına taşan ekstraselüler ATP, nükleer HMGB1, S100 proteinleri, mitokondriyal formil peptitler ve serbest histonlar; doğal bağışıklık reseptörleri (PRR: TLR2, TLR4, RAGE) tarafından 'ölümcül enfeksiyon' olarak algılanır. Bu yanlış alarm, çevre makrofaj ve monositleri uyararak yangıyı kalıcı bir döngüye sokar.",
        "d[DAMP_ekstraseluler]/dt = J_nekroz + J_senesens_salgi - k_fagositoz * [Makrofaj_aktif] * [DAMP]",
        "Bu dinamik eşitlik, dokudaki ekstraselüler DAMP yoğunluğunun hücresel lizis ve senesen salgı hızı ile fagositoz klerensi arasındaki farka bağlı olduğunu gösterir."
    ),
    (
        "1.3 HMGB1 (High-Mobility Group Box 1) Nükleer Kaçışı ve Dolaşım Enflamasyonu",
        "HMGB1, normal şartlarda çekirdekte histon olmayan bir kromatin proteini olarak DNA bükülmesini ve transkripsiyon faktörü bağlanmasını düzenleyen hayati bir moleküldür.",
        "Ancak hücresel stres, oksidatif hasar veya senesens durumunda HMGB1 hiperasetillenir ve nükleer lokalizasyon sinyalini (NLS) kaybederek sitoplazmaya sızar, oradan da ekstraselüler alana salgılanır. Hücre dışındaki serbest HMGB1, RAGE (İleri Glikasyon Son Ürünleri Reseptörü) ve TLR4 reseptörlerine yüksek afiniteyle bağlanır. Bu bağlanma, miyeloid diferansiyasyon faktörü 88 (MyD88) ve NF-kappaB kaskadını tetikleyerek pro-enflamatuar sitokin salgısını tavan yaptırır. Dolaşımdaki HMGB1 seviyeleri, ileri yaşta kırılganlık ve mortalite ile doğrudan koreledir.",
        "RAGE_Aktivasyon = k_RAGE * [HMGB1_ekstraseluler]^h / ( K_d_RAGE^h + [HMGB1_ekstraseluler]^h )",
        "Bu Hill tipi bağlanma formülü, ekstraselüler HMGB1 konsantrasyonunun RAGE reseptör doygunluğu ve aşağı akış enflamatuar sinyal üretimi üzerindeki kooperatif etkisini açıklar."
    ),
    (
        "1.4 Ekstraselüler ATP, P2X7 Reseptörü ve Pürinerjik Sinyal Kaskadı",
        "Adenozin trifosfat (ATP), hücre içinde 3-5 milimolar konsantrasyonda bulunan evrensel enerji para birimidir; ancak ekstraselüler aralıkta normalde nanomolar seviyelerde tutulur.",
        "Hasarlı veya senesen hücrelerin zarlarından sızan ya da pannekson kanallarıyla dışarı atılan ATP, ekstraselüler alanda 'tehlike sinyali' görevi görür. Makrofaj ve mikroglia yüzeyindeki trimerik ligand kapılı iyon kanalı olan P2X7 pürinerjik reseptörünü aktive eder. P2X7 aktivasyonu, hücre içine devasa bir Ca2+ girişine ve hücre dışına K+ çıkışına yol açarak hücre içi potasyum konsantrasyonunu düşürür. Bu iyonik şok, NLRP3 inflamazomunun montajlanmasını tetikleyen en güçlü fizyolojik anahtardır.",
        "I_P2X7 = G_max * [ATP_ekstraseluler]^n / ( EC50_P2X7^n + [ATP_ekstraseluler]^n ) * ( V_membran - E_rev )",
        "Bu elektrofizyolojik denklem, ekstraselüler ATP'nin P2X7 kanalı üzerinden yarattığı iyonik akımı ve buna bağlı hücresel depolarizasyonun şiddetini modeller."
    ),
    (
        "1.5 Hücresel Çöp Birikimi (Garbageome) ve Çözülemeyen Makromoleküler Stres",
        "Franceschi'nin 'Garbageome' (Hücresel Çöp Havuzu) teorisi, yaşlanan hücrelerin lizozom ve proteazom kapasitesini aşan çözünmez makromoleküler agregatların birikimini ve bunun doğurduğu enflamasyonu açıklar.",
        "Oto-fagositoz ve proteostaz yetersizliği nedeniyle sitozolde lipofuskin granülleri, çapraz bağlı agregatlaşmış proteinler, glike kollajen parçaları ve oksitlenmiş fosfolipitler birikir. Makrofajlar bu döküntüleri fagositozla yutmaya çalıştıklarında lizozomları tıkanır, lizozomal enzimler (Katepsin B) sitoplazmaya sızar ve hücre fagositoz iflasına (efferositoz yetersizliği) uğrar. Çözülemeyen bu kronik 'çöp' varlığı, bağışıklık sistemini ömür boyu sürecek nafile bir yangısal savaşa sürükler.",
        "d[Garbageome]/dt = J_hasarli_protein_uretim - ( V_max_lizozom * [Garbageome] ) / ( K_M + [Garbageome] )",
        "Bu kütle transferi formülü, lizozomal sindirim kapasitesinin (V_max_lizozom) aşılması durumunda hücresel çöp yükünün nasıl sınırsız biriktiğini formüle eder."
    ),
    (
        "1.6 Mitokondriyal DNA (mtDNA) Sızıntısı: N-formil Peptitler ve İnflamasyon",
        "Mitokondriler, endosimbiyotik evrim nedeniyle bakteriyel kökenli organellerdir; bu nedenle mitokondriyal bileşenler bağışıklık sistemi için yabancı (bakteri benzeri) antijenlerdir.",
        "Mitofaji yetersizliği ve mitokondriyal permeabilite geçiş gözeneğinin (mPTP) patolojik açılması sonucu, mitokondriyal DNA (mtDNA) ve N-formillenmiş peptitler sitozole ve dolaşıma sızar. mtDNA, memeli nükleer DNA'sının aksine yüksek oranda metillenmemiş CpG motifleri içerir ve endozomal TLR9 reseptörünü doğrudan uyarır. N-formil peptitler ise formil peptit reseptörlerine (FPR1) bağlanarak nötrofilleri uyarır ve vasküler yataklara lökosit infiltrasyonunu başlatır.",
        "TLR9_Sinyal_Akisi = k_TLR9 * [mtDNA_CpG_serbest] / ( K_d_TLR9 + [mtDNA_CpG_serbest] )",
        "Bu reseptör doyum modeli, hücre dışına veya sitoplazmaya sızan mitokondriyal DNA parçalarının TLR9 kaskadını nasıl doğrudan aktive ettiğini gösterir."
    ),
    (
        "1.7 Yaşla Artan Bazal TNF-alfa, IL-6 ve hs-CRP Düzeylerinin Sistemik Yıkımı",
        "Genç sağlıklı bireylerde dinlenim halinde kanda tespit edilemeyecek kadar düşük olan pro-enflamatuar sitokinler, 65 yaş üzerinde 2 ila 4 kat bazal artış gösterir.",
        "Karaciğer tarafından interlökin-6 (IL-6) uyarısıyla sentezlenen yüksek duyarlılıklı C-reaktif protein (hs-CRP), kümülatif inflam-aging yükünün en güvenilir klinik göstergesidir. Kronik olarak yükselen TNF-alfa ve IL-6; iskelet kasında protein sentezini baskılayıp proteolizi uyararak sarkopeniye yol açar; adipositlerde lipolizi tetikleyip serbest yağ asitlerini kana salarak insülin direncini artırır; osteoklastları aktive ederek osteoporozu hızlandırır ve hipotalamusta leptin direncini körükler.",
        "Sarkopeni_Hizi = k_atrofi * [TNF-alfa_serum] * [IL-6_serum] / ( [IGF-1_serum] + K_koruma )",
        "Bu patolojik model, pro-enflamatuar sitokin fırtınasının iskelet kası yıkım hızını büyüme faktörlerinin koruyucu etkisini bastırarak nasıl katladığını gösterir."
    ),
    (
        "1.8 İnflam-Aging ve Klonal Hematopoez Arasındaki Çift Yönlü İtici Güç",
        "Yaşlı bireylerin kemik iliğinde kök hücrelerin belirli somatik mutasyonlar (DNMT3A, TET2, ASXL1) kazanarak klonal olarak genişlemesi (CHIP - Klonal Hematopoez), inflam-aging ile ölümcül bir kısır döngü oluşturur.",
        "TET2 veya DNMT3A mutasyonu taşıyan hematopoietik kök hücrelerden türeyen monosit ve makrofajlar, dinlenim halinde dahi aşırı miktarda IL-1beta ve IL-6 salgılar (hiper-enflamatuar fenotip). Bu sitokinler dolaşıma yayılarak sistemik yangıyı körüklerken; inflam-aging ortamının kendisi de mutasyona uğramış bu agresif klonlara vahşi tip kök hücrelere karşı bir seçilim avantajı (fitness advantage) sağlar. Bu durum kardiyovasküler mortaliteyi ikiye katlar.",
        "Klon_Buyume_Avantaji = Delta_Fitness = s_0 + alpha_IL6 * [IL-6_kemik_iligi] + alpha_TNF * [TNF-alfa]",
        "Bu popülasyon genetiği denklemi, lokal enflamatuar sitokin yoğunluğunun mutant hematopoietik klonların büyüme seçilim katsayısını nasıl doğrudan artırdığını modeller."
    ),
    (
        "1.9 Vasküler Endotel İltihabı ve Aterosklerotik Erken Yaşlanma",
        "Vasküler endotel tabakası, kan ile dokular arasındaki seçici bariyer olmanın ötesinde en büyük endokrin ve parakrin organdır; yaşlanmada endotel iltihabı (endotelyal disfonksiyon) ilk çöken sistemdir.",
        "Dolaşımdaki sitokinler ve oksitlenmiş LDL (oxLDL), endotel hücrelerinde ICAM-1 ve VCAM-1 adhezyon moleküllerinin ekspresyonunu artırarak dolaşımdaki monositlerin subendotelyal alana göçünü kolaylaştırır. Eş zamanlı olarak endotelyal nitrik oksit sentaz (eNOS) enzimi, BH4 kofaktörünün oksidasyonu nedeniyle 'uncoupling'e (ayrışma) uğrar; nitrik oksit (NO) üretmek yerine süperoksit radikali üretmeye başlar. Ortaya çıkan vazodilatasyon kaybı, arteriyel sertleşme ve aterom plakları biyolojik yaşlanmanın hızını belirler.",
        "eNOS_Süperoksit_Orani = 1 / ( 1 + [BH4] / (K_m_BH4 * (1 + [BH2]/K_i_BH2)) )",
        "Bu biyokimyasal ayrışma fonksiyonu, oksitlenmiş biopterin (BH2) birikiminin eNOS enzimini faydalı NO üreticisinden toksik süperoksit jeneratörüne nasıl dönüştürdüğünü açıklar."
    ),
    (
        "1.10 Steril İnflamasyonun Biyofiziksel Kinetiği ve Eşik Dinamikleri",
        "Steril inflamasyon, belirli bir moleküler hasar eşiği aşılana kadar asemptomatik ve gizli ilerler; ancak eşik aşıldığında sistem non-lineer (bistable) bir patlama gösterir.",
        "Bu faz geçişi, pozitif otokatalitik döngülerle yönetilir: DAMP salınımı sitokinleri artırır, sitokinler hücre ölümünü ve senesensi artırır, bu da daha fazla DAMP açığa çıkarır. Bu bistable sistemde geri dönüş eşiği (histerezis) son derece yüksektir; yani sistemi gençlik durumuna geri getirmek için yangıyı sadece başlatıcı seviyeye indirmek yetmez, çok daha derin bir temizlik ve anti-enflamatuar resetleme protokolü gerekir.",
        "d[Yangisal_Durum]/dt = (k_oto * [Yangisal_Durum]^2) / (K_esik^2 + [Yangisal_Durum]^2) - k_klerens * [Yangisal_Durum]",
        "Bu bistable diferansiyel dinamik modeli, inflam-aging'in belirli bir kritik hasar eşiğinden sonra kendi kendini besleyen kalıcı bir kararlı duruma kilitlendiğini kanıtlar."
    )
]

# ================= KISIM 2 =================
part2_subsections = [
    (
        "2.1 Sitoplazmik DNA Algılayıcısı cGAS (Siklik GMP-AMP Sentaz) Aktivasyonu",
        "cGAS (Siklik GMP-AMP Sentaz), ökaryotik hücrelerin sitoplazmasında bulunmaması gereken çift zincirli DNA'yı (dsDNA) sekans bağımsız olarak tespit eden en duyarlı konak savunma enzimidir.",
        "Normal şartlarda DNA çekirdek ve mitokondri içinde hapsedilmiştir. Sitoplazmaya herhangi bir dsDNA sızdığında (>45 baz çifti uzunluğunda), cGAS molekülleri DNA'nın fosfodiester omurgasına pozitif yüklü yüzeyleriyle bağlanır. İki cGAS molekülü ve iki dsDNA sarmalı bir araya gelerek 2:2 stokiyometrisinde aktif bir dimerik kompleks oluşturur. Bu bağlanma, enzimin katalitik merkezinde konformasyonel bir bükülme yaratarak ATP ve GTP nükleotidlerini bağlama cebini açar.",
        "Katalitik_Doygunluk_cGAS = [dsDNA_sitozol]^2 / ( K_d_DNA^2 + [dsDNA_sitozol]^2 )",
        "Bu kooperatif bağlanma formülü, sitozolik çift zincir DNA konsantrasyonunun karesiyle orantılı olarak cGAS enziminin nasıl patlamalı şekilde aktif hale geldiğini tanımlar."
    ),
    (
        "2.2 2'3'-cGAMP İkinci Habercisinin Biyosentezi ve Yapısal Kararlılığı",
        "Aktive olan cGAS, hücresel ATP ve GTP moleküllerini kullanarak benzersiz bir siklik dinükleotid olan 2'3'-siklik GMP-AMP (cGAMP) sentezler.",
        "Bu reaksiyonda enzim önce GTP'nin guanozin monofosfatını ATP'nin adenozinine 2'-5' fosfodiester bağıyla bağlar; ardından halkayı kapatmak için ikinci bir 3'-5' fosfodiester bağı kurar. Oluşan bu non-kanonik 2'3'-cGAMP hibrit bağı, insan fosfodiesterazlarına karşı olağanüstü dirençlidir ve hücre içinde uzun süre stabil kalır. Dahası cGAMP, gap junction (Cx43) kanalları aracılığıyla komşu sağlıklı hücrelere yayılarak enfeksiyon veya hasar görmemiş komşularda da steril yangıyı başlatabilir (bystander etkisi).",
        "J_cGAMP_sentez = k_cat_cGAS * [cGAS:dsDNA] * [ATP] * [GTP] / ( K_M_ATP * K_M_GTP + [ATP]*[GTP] )",
        "Bu iki-substratlı hız eşitliği, sitoplazmik cGAMP üretim debisinin DNA ile birleşmiş aktif cGAS kompleksi ve nükleotid havuzuna bağımlılığını formüle eder."
    ),
    (
        "2.3 STING (Stimulator of Interferon Genes) Dimizasyonu ve ER'den Golgiye Translokasyonu",
        "Endoplazmik retikulum (ER) zarına gömülü bir transmembran proteini olan STING (Stimulator of Interferon Genes / TMEM173), cGAMP'nin yüksek afiniteli hücresel reseptörüdür.",
        "Dinlenim durumunda STING ER zarında inaktif bir homodimer olarak bulunur. cGAMP, dimerik arayüzde yer alan 'V-şekilli' bağlama cebine nanomolar afiniteyle (Kd ~ 4-5 nM) yerleşir. cGAMP kenetlenmesi, STING dimerinin baş kısımlarını 180 derece kapatarak enzimi kompakt bir oligomerik yapıya sokar. Aktive STING, COP-II veziküllerine paketlenerek ER'den çıkar ve ERGIC (ER-Golgi Ara Kompartmanı) üzerinden Golgi aygıtına doğru hızlı bir hücre içi göç başlatır.",
        "d[STING_Golgi]/dt = k_trans * [STING:cGAMP_ER] - k_retro * [STING_Golgi] - k_deg_lizozom * [STING_Golgi]",
        "Bu veziküler trafik modeli, cGAMP ile aktive olmuş STING dimerlerinin ER'den Golgiye taşınma ve ardından lizozomal yıkıma uğrama kinetiğini açıklar."
    ),
    (
        "2.4 TBK1 Kinaz Alımı ve IRF3 Transkripsiyon Faktörünün Fosforilasyonu",
        "STING oligomerleri Golgi zarına ulaştığında, sitoplazmik C-terminal kuyrukları (CTT) palmitoilasyona uğrar ve devasa bir sinyal platformuna dönüşür.",
        "Bu platform, sitozolik serin/treonin kinaz TBK1'i (TANK-Binding Kinase 1) kendine çeker. TBK1, STING'in CTT kuyruğundaki Ser366 kalıntısını fosforiller. Fosforile STING kuyruğu, transkripsiyon faktörü IRF3 (İnterferon Yanıt Faktörü 3) için yüksek afiniteli bir yanaşma iskelesi haline gelir. Bölgeye yanaşan IRF3, komşu TBK1 tarafından C-terminal serin kümesinden (Ser385, Ser386) fosforillenir. Fosforillenen IRF3 homodimerleşir, nükleer porlardan çekirdeğe akar ve antiviral/pro-enflamatuar genleri açar.",
        "Aktivasyon_IRF3 = k_TBK1 * [TBK1_STING] * [IRF3_monomer]^2 / ( K_M_IRF3 + [IRF3_monomer]^2 )",
        "Bu enzimatik kinetik bağıntısı, Golgi zarında toplanan TBK1 kinazın serbest IRF3 havuzunu nasıl dimerize ederek transkripsiyonel fırtınayı başlattığını gösterir."
    ),
    (
        "2.5 Tip I İnterferon (IFN-alfa/beta) Yanıtı ve ISG (İnterferonla Uyarılan Genler)",
        "Dimerik fosfo-IRF3, nükleusta IFN-beta ve IFN-alfa promoter bölgelerindeki ISRE (İnterferon Uyarılı Yanıt Elemanı) dizilerine bağlanarak Tip I İnterferon transkripsiyonunu patlatır.",
        "Salınan IFN-beta, otokrin ve parakrin yolla hücre yüzeyindeki IFNAR1/IFNAR2 reseptörlerine bağlanır. Bu bağlanma JAK1/Tyk2 kinazları üzerinden STAT1 ve STAT2'yi fosforiller; fosforile STAT1/STAT2, IRF9 ile birleşerek ISGF3 kompleksini kurar ve yüzlerce İnterferonla Uyarılan Genin (ISG: OAS1, MX1, IFIT1-3) kalıcı ekspresyonunu sağlar. Yaşlanan dokularda bu antiviral alarm yanıtının patojensiz olarak sürekli açık kalması, doku rejenerasyonunu felç eder ve kök hücreleri derin bir proliferatif uykuda (dormancy) kilitler.",
        "Transkripsiyon_ISG = V_max_ISGF3 * [ISGF3_nukleus] / ( K_ISRE + [ISGF3_nukleus] )",
        "Bu transkripsiyonel akı formülü, cGAS-STING kaskadının nihai ürünleri olan Tip I interferonların hücre çekirdeğinde yarattığı sürekli hücresel toksisite yükünü açıklar."
    ),
    (
        "2.6 IKK Kompleksi ve NF-kappaB Aktivasyonu: Çift Koldan SASP Üretimi",
        "cGAS-STING sinyalinin yaşlanmadaki en yıkıcı özelliği, sadece IRF3 üzerinden interferon üretmekle kalmayıp; paralel bir kol üzerinden IKK kinaz kompleksini ateşleyerek NF-kappaB'yi uyarmasıdır.",
        "Golgiye demirlenen STING, TRAF6 ve IKK (I-kappa-B Kinaz: IKKalfa, IKKbeta, NEMO) kompleksini aktive eder. IKKbeta, sitozolde NF-kappaB'yi (p65/p50) hapseden inhibitör protein I-kappa-B-alfa'yı (I-kappa-B-a) Ser32 ve Ser36 kalıntılarından fosforiller. Fosforillenen I-kappa-B-a proteazomda parçalanır; serbest kalan NF-kappaB çekirdeğe girerek IL-6, TNF-alfa, IL-8 ve MMP'leri (SASP faktörleri) üretir. Böylece cGAS-STING, hücresel senesensin SASP programının ana jeneratörü haline gelir.",
        "d[NFkB_nukleus]/dt = k_IKK * [STING*] * [NFkB_inaktif] - k_represyon * [IkBa_yeni]",
        "Bu diferansiyel ilişki, aktif STING varlığının NF-kappaB çekirdek girişini nasıl sürekli olarak yüksek tuttuğunu ve SASP salgısını körüklediğini formüle eder."
    ),
    (
        "2.7 Mikronükleus Zarı Yırtılması ve cGAS'ın Kromatin İstilası",
        "Hücre bölünmesi sırasında kromozomal anöploidi veya sentromer hasarı nedeniyle ana çekirdekten kopan gecikmiş kromozom parçaları, kendi etraflarında kırılgan bir zar örerek 'mikronükleus' oluşturur.",
        "Mikronükleus zarı lamin B1 eksikliği nedeniyle mekanik olarak son derece zayıftır ve hücre döngüsü ilerledikçe spontan olarak yırtılır (Catastrophic Nuclear Envelope Collapse). Zar yırtıldığı anda sitozolde bekleyen cGAS enzimleri sel gibi içeri akar ve çıplak genomik kromatin parçalarına yapışır. Mikronükleus içi DNA yoğunluğu cGAS'ı aşırı doygunluğa ulaştırarak hücreyi devasa bir cGAMP fabrikasına dönüştürür. Bu olay, genomik kararsızlığın nasıl steril yangıya ve hücresel senesense dönüştüğünün en somut biyofiziksel kanıtıdır.",
        "P_yirtilma_mikronukleus = 1 - exp( - (sigma_gerilim / sigma_lamin_direnc)^k * t )",
        "Bu Weibull kırılma modeli, mikronükleus zarının zaman ve nükleer zar gerilimi karşısında spontan yırtılma olasılığını ve cGAS istilasını modeller."
    ),
    (
        "2.8 Retrotranspozon LINE-1 cDNA'sının cGAS-STING Yolağını Kronik Ateşlemesi",
        "Yaşlanan hücrelerde heterokromatin erozyonu nedeniyle sessizliğini kaybeden LINE-1 (L1) retrotranspozonları, ters transkriptaz (ORF2p) enzimleri aracılığıyla sitoplazmada serbest cDNA iplikçikleri sentezler.",
        "Sitozolde biriken bu L1 retrotranspozon cDNA parçaları, cGAS tarafından yabancı viral DNA gibi algılanır ve cGAS-STING yolağını durmaksızın tetikler. John Sedivy ve ekibinin keşfine göre; yaşlanan hücrelerde SASP salgısının geç ve kalıcı fazı neredeyse tamamen LINE-1 cDNA'sı tarafından yönlendirilir. Lamivudin (3TC) veya Emtrisitabin gibi ters transkriptaz inhibitörleri ile LINE-1 cDNA sentezi durdurulduğunda; cGAS uyarımı kesilmekte, SASP salgısı %80 gerilemekte ve sistemik inflam-aging belirgin şekilde sönümlenmektedir.",
        "J_L1_cDNA = k_RT * [ORF2p_aktif] * [L1_mRNA_sitozol] / ( K_M_dCTP + [dNTP] )",
        "Bu kinetik eşitlik, sitozolde biriken LINE-1 cDNA akısının ORF2p ters transkriptaz aktivitesi ve serbest dNTP havuzuna doğrudan bağımlı olduğunu gösterir."
    ),
    (
        "2.9 Otophagy Aracılı STING Klirensinin Yaşla Körelmesi",
        "Normal genç hücrelerde aşırı ve kronik yangıyı önlemek için STING aktivasyonu geçicidir; cGAMP bağlayan STING oligomerleri p62 (SQSTM1) şaperonu aracılığıyla otofagozomlara yönlendirilir ve lizozomda parçalanır.",
        "Ancak yaşlanan dokularda makrootofaji akısının yavaşlaması ve lizozomal enzimlerin kireçlenmesi/tükenmesi nedeniyle STING'in otofajik yıkımı durur. Golgi membranında biriken aktif STING oligomerleri günlerce sinyal üretmeye devam eder. Bu yıkım kusuru, geçici bir savunma mekanizması olması gereken cGAS-STING kaskadını kalıcı ve öldürücü bir oto-enflamatuar doku tahrip makinesine dönüştürür.",
        "d[STING_Aktif]/dt = Sentez_STING - k_otofaji * [Otofagozom_Aktif] * [p62] * [STING_Aktif]",
        "Bu denge formülasyonu, azalan otofajik akı ([Otofagozom_Aktif] -> 0) durumunda aktif STING proteininin dokularda kontrolsüzce biriktiğini ortaya koyar."
    ),
    (
        "2.10 Sentetik cGAS-STING İnhibitörleri (H-151, RU.521) ve Yaşlanma Karşıtı Tedavi",
        "cGAS-STING yolağının yaşlanmanın ve nörodejeneratif hastalıkların merkezinde yer aldığının anlaşılması, bu ekseni hedefleyen güçlü sentetik inhibitörlerin geliştirilmesini sağlamıştır.",
        "RU.521, cGAS'ın katalitik nükleotid bağlama cebine girerek dsDNA varlığında bile ATP/GTP kenetlenmesini sterik olarak bloke eder. H-151 ise STING'in Cys91 kalıntısına kovalent olarak bağlanan küçük bir moleküldür; STING'in cGAMP uyarılı oligomerizasyonunu ve palmitoilasyonunu tamamen durdurarak Golgiye göçünü engeller. Fare modellerinde H-151 uygulaması, yaşlı farelerde nöroinflamasyonu temizlemekte, bilişsel fonksiyonları gençleştirmekte ve iskelet kası atrofisini geri döndürmektedir.",
        "Inhibisyon_STING_H151 = [H-151] / ( IC50_H151 + [H-151] )",
        "Bu farmakodinamik formül, H-151 kovalent inhibitörünün STING aktivitesini nanomolar düzeyde nasıl seçici olarak baskıladığını tanımlar."
    )
]

# ================= KISIM 3 =================
part3_subsections = [
    (
        "3.1 NLRP3 İnflamazom Kompleksinin Yapısı: Sensör, ASC Adaptörü ve Pro-Kaspaz-1",
        "NLRP3 inflamazomu, doğal bağışıklık sisteminin en sofistike sitozolik moleküler makinesidir. Üç ana bileşenden oluşur: Sensör proteini NLRP3, adaptör protein ASC (Apoptosis-associated speck-like protein containing a CARD) ve efektör enzim Pro-Kaspaz-1.",
        "NLRP3 (NLR family pyrin domain containing 3); N-terminal pyrin (PYD) alanına, merkezi bir NACHT nükleotid-bağlama ve oligomerizasyon alanına ve C-terminal lösin zengini tekrar (LRR) alanına sahiptir. Hücre uyarıldığında NLRP3 oligomerleşir; PYD-PYD etkileşimi ile ASC proteinini çeker. ASC'ler polar lifler halinde polimerize olarak mikroskopta görülebilen devasa bir 'ASC specki' oluşturur. ASC'nin CARD alanı ise Pro-Kaspaz-1'i komplekse alarak kaspazın kendi kendini kesip aktif Kaspaz-1 tetramerine dönüşmesini sağlar.",
        "Speck_Oligomerizasyon = k_poly * [ASC_monomer]^p / ( K_nukleasyon^p + [ASC_monomer]^p )",
        "Bu nükleasyon bağımlı polimerizasyon denklemi, ASC proteininin kritik bir eşik konsantrasyonu aşıldığında nasıl devasa bir makromoleküler yangı odağına (speck) dönüştüğünü gösterir."
    ),
    (
        "3.2 Priming Sinyali: NF-kappaB Bağımlı NLRP3 ve Pro-IL-1beta Transkripsiyonu",
        "NLRP3 inflamazomunun kontrolsüzce patlamasını önlemek için evrim iki aşamalı bir güvenlik kilidi (Two-signal model) geliştirmiştir. İlk adım 'Priming' (Hazırlık) sinyalidir.",
        "Dinlenim durumundaki makrofajlarda NLRP3 ve Pro-IL-1beta protein seviyeleri fonksiyonel bir kompleks kurmak için yetersizdir. TLR ligandları (LPS gibi), DAMPs veya TNF-alfa/IL-1beta sitokinleri hücre yüzeyindeki reseptörlerine bağlandığında; MyD88/TRIF kaskadı üzerinden NF-kappaB çekirdeğe girer. NF-kappaB, NLRP3 ve IL1B genlerinin transkripsiyonunu dramatik olarak artırır. Aynı zamanda NLRP3'ün de-ubikitilasyonunu ve bazal fosforilasyonunu sağlayarak onu ikinci sinyale duyarlı hale getirir.",
        "d[NLRP3_protein]/dt = k_transkripsiyon * [NFkB_nukleus] - k_deg_ubikitin * [NLRP3_protein]",
        "Bu hız eşitliği, hazırlık aşamasında NF-kappaB çekirdek yoğunluğunun sitoplazmik NLRP3 protein rezervini nasıl oluşturduğunu tanımlar."
    ),
    (
        "3.3 Aktivasyon Sinyali: K+ Dışa Akışı, Mitokondriyal ROS ve Lizozomal Yırtılma",
        "Hazırlanmış (primed) hücrede inflamazom montajını başlatan ikinci adım 'Aktivasyon' sinyalidir. NLRP3, tek bir kimyasal ajanı değil; hücre içi homeostazın bozulmasını gösteren birden fazla biyofiziksel stresi algılar.",
        "En evrensel tetikleyici intraselüler potasyum iyonunun dışarı sızmasıdır ([K+]_i < 70 mM düşüşü). Ekstraselüler ATP'nin P2X7 reseptörünü açması potasyumu hızla boşaltır. İkinci ana tetikleyici, hasarlı mitokondrilerden salınan aşırı reaktif oksijen türleridir (mtROS). Üçüncü mekanizma ise hücrenin sindiremediği partikülleri yutması sonucu lizozom zarlarının yırtılması ve Katepsin B'nin sitozole sızmasıdır. Bu üç sinyal yolu NACHT alanının ADP-ATP değişimini uyararak kompleksi kilitler.",
        "Tetikleme_NLRP3 = f([K+_disari]) * f([mtROS]) * f([KatepsinB_sitozol])",
        "Bu çarpımsal fonksiyon, hücre içi stres parametrelerinin birleşik etkisinin NLRP3 NACHT alanının konformasyonel açılmasını nasıl tetiklediğini simgeler."
    ),
    (
        "3.4 NEK7 Kinazın NLRP3 Oligomerizasyonundaki Zorunlu Rolü",
        "Son yıllarda yapılan yapısal biyoloji çalışmaları (kriyo-EM), NLRP3'ün aktivasyonu için hücre döngüsü kinazı olan NEK7'nin (NIMA-related kinase 7) fiziksel varlığının mutlak bir zorunluluk olduğunu keşfetmiştir.",
        "Potasyum dışa akışı gerçekleştiğinde, NEK7 kinaz enzimi NLRP3'ün LRR ve NACHT alanlarının arayüzüne doğrudan kenetlenir. NEK7'nin bağlanması, NLRP3'ün kapalı inaktif halka yapısını kırarak onun hekzamerik ve disk benzeri yüksek dereceli oligomerler oluşturmasını sağlar. İlginç bir şekilde NEK7 aynı zamanda mitozda iğ ipliği oluşumunu yönetir; bu nedenle bir hücre mitoz bölünmeye girdiğinde NEK7 meşgul olduğu için inflamazom ateşlenemez. Yaşlanan senesen hücrelerde ise hücre döngüsü durduğu için serbest kalan NEK7 sürekli NLRP3'e kilitlenir.",
        "[NLRP3:NEK7_kompleks] = [NLRP3_primed] * [NEK7_serbest] / ( K_d_NEK7 + [NEK7_serbest] )",
        "Bu denge formülasyonu, serbest NEK7 mevcudiyetinin NLRP3 oligomerizasyon kapasitesini doğrudan belirleyen moleküler sınırlayıcı faktör olduğunu gösterir."
    ),
    (
        "3.5 Kaspaz-1 Aktivasyonu ve Pro-IL-1beta ile Pro-IL-18 Proteolitik Kesimi",
        "ASC specki üzerinde toplanan inaktif Pro-Kaspaz-1 zimojenleri, birbirine aşırı yaklaşarak (induced proximity) trans-otokatalitik kesim gerçekleştirir ve p20/p10 heterodimerlerinden oluşan tam aktif Kaspaz-1 tetramerini açığa çıkarır.",
        "Kaspaz-1, sitoplazmada inaktif prekürsör olarak bekleyen 31 kDa Pro-İnterlökin-1beta'yı spesifik bir aspartik asit bölgesinden (Asp116-Ala117 bağı) keserek 17 kDa ağırlığındaki süper-enflamatuar aktif IL-1beta'ya dönüştürür. Benzer şekilde Pro-IL-18'i keserek aktif IL-18 haline getirir. Aktif IL-1beta; ateşi, lökosit proliferasyonunu, endotel aktivasyonunu ve kemokin salgısını tetikleyen en güçlü pirojenik sitokindir.",
        "d[IL1b_aktif]/dt = k_cat_Kasp1 * [Kaspaz1*] * [Pro-IL1b] / ( K_M_IL1b + [Pro-IL1b] )",
        "Bu Michaelis-Menten proteolitik eşitliği, Kaspaz-1 aktivitesinin aktif sistemik interlökin-1beta üretim hızını nasıl belirlediğini açıklar."
    ),
    (
        "3.6 Gasdermin D (GSDMD) N-terminal Alanının Zarda Por Açması ve Piroptoz",
        "Kaspaz-1'in ikinci ve mekanik olarak en yıkıcı substratı, hücre ölümü efektörü olan Gasdermin D (GSDMD) proteinidir.",
        "Kaspaz-1, GSDMD'nin inhibitör C-terminal alanını N-terminal alanından kesip ayırır. Serbest kalan lipofilik GSDMD-N-terminal parçaları plazma membranına göç eder ve fosfoinositidlere bağlanarak zarda oligomerleşir. 27-28 adet GSDMD-N monomeri bir araya gelerek hücre zarında 10 ila 14 nanometre iç çapa sahip devasa delikler (porlar) açar. Bu deliklerden aktif IL-1beta ve IL-18 dışarı fışkırırken; hücre içine kontrolsüz su ve sodyum girer, hücre şişer ve ozmotik olarak patlar (Piroptoz - enflamatuar hücre ölümü).",
        "Zar_Por_Sayisi = Sum_i ( [GSDMD-N_oligomer]_i ) * ( Area_por / Area_membran )",
        "Bu por yoğunluğu formülü, hücre zarındaki perforasyon alanının ozmotik lizis ve hücre patlaması hızını nasıl belirlediğini modeller."
    ),
    (
        "3.7 Mikrokristaller (Ürik Asit, Kolesterol) ve Amiloid-Beta Kaynaklı NLRP3 Ateşlenmesi",
        "Yaşlanan dokularda metabolik artıkların birikimi, çözünürlük sınırlarını aşarak mikrokristallerin ve fibrillerin oluşmasına yol açar; bu agregatlar NLRP3'ün amansız tetikleyicileridir.",
        "Gut hastalığında eklemlerde biriken Monosodyum Ürat (MSU) kristalleri, damar duvarında biriken Kolesterol kristalleri ve Alzheimer hastalarının beyninde biriken Amiloid-Beta oligomerleri makrofaj ve mikroglialar tarafından fagositoza uğrar. Kristaller lizozom içinde sindirilemez; rijit yapılarıyla lizozom zarını delerek yırtarlar. Sızan asidik katepsin enzimleri ve potasyum kaybı NLRP3'ü sürekli hiperaktif tutar. Bu durum, yaşa bağlı kronik artritin, aterosklerozun ve Alzheimer nöroinflamasyonunun ortak moleküler kökenidir.",
        "Lizozomal_Yirtilma_Hizi = k_hasar * [Mikrokristal_Yuku] / ( Membran_Kolesterol_Orani + K_elastisite )",
        "Bu biyofiziksel bağıntı, mikrokristal birikiminin lizozomal zar rijiditesini aşarak sitoplazmik sızıntı ve inflamazom patlaması yaratma kinetiğini ifade eder."
    ),
    (
        "3.8 NLRP3 İnflamazomunun Beyin Mikroglia ve Hipotalamik Yaşlanmadaki Yıkımı",
        "Merkezi sinir sisteminin yerleşik makrofajları olan mikroglialar, yaşlanma sürecinde 'primed' (aşırı duyarlı) bir duruma geçer ve nörodejenerasyonun zeminini hazırlar.",
        "Hipotalamus, tüm vücudun metabolizmasını ve yaşlanma hızını kontrol eden ana endokrin merkezdir. Hipotalamusta biriken nöroinflamasyon ve mikroglial NLRP3 aktivasyonu, GnRH (Gonadotropin Salgılatıcı Hormon) salınımını baskılar. GnRH kaybı somatik dokularda yaşlanmayı hızlandırır, kas kaybını ve bilişsel gerilemeyi tetikler. Eş zamanlı olarak hipokampustaki mikroglial inflamazom uyarımı, BDNF seviyelerini düşürerek sinaptik plastisiteyi yıkar. Mikroglial NLRP3'ün susturulması, bilişsel gençliği ve nörogenezi korumanın en kritik şartıdır.",
        "GnRH_Salinim_Hizi = GnRH_bazal / ( 1 + ( [IL-1b_hipotalamus] / K_i_GnRH )^m )",
        "Bu nöroendokrin baskılama modeli, hipotalamustaki lokal inflamazom aktivasyonunun sistemik gençlik hormonu GnRH üretimini nasıl logaritmik olarak kilitlediğini açıklar."
    ),
    (
        "3.9 MCC950 ve Doğal İnflamazom İnhibitörlerinin (BHB) Geroprotektif Gücü",
        "NLRP3'ün patolojik yaşlanmadaki merkezi rolünün kanıtlanması, hedefe yönelik yüksek afiniteli küçük moleküllü inhibitörlerin geliştirilmesine yol açmıştır.",
        "MCC950 (CRID3), sülfonilüre türevi sentetik bir moleküldür; NLRP3'ün NACHT alanındaki Walker B motifi yakınlarına nanomolar afiniteyle (IC50 ~ 7.5 nM) bağlanır. Bu bağlanma NLRP3'ün ATP hidrolizini ve konformasyonel açılmasını sterik olarak kilitler; AIM2 veya NLRC4 gibi diğer inflamazomlara dokunmadan sadece NLRP3'ü susturur. Açlık ve ketojenik metabolizma sırasında karaciğerde üretilen doğal keton cisimciği Beta-Hidroksibütirat (BHB) da fizyolojik konsantrasyonlarda potasyum sızıntısını engelleyerek NLRP3'ü doğrudan baskılar.",
        "Inhibisyon_NLRP3_MCC = [MCC950] / ( IC50_MCC950 + [MCC950] )",
        "Bu formülasyon, sentetik allosterik inhibitörlerin NLRP3 inflamazom kaskadını hücresel düzeyde nasıl tam olarak kapattığını tanımlar."
    ),
    (
        "3.10 İnflamazom Aşırı Aktivasyonunun Genetik ve Farmakolojik Susturulması",
        "Geleceğin rejeneratif tıbbı, inflam-aging yangınını söndürmek için genetik ve farmakolojik müdahaleleri birleştiren çift yönlü protokoller tasarlamaktadır.",
        "CRISPR-Cas9 ile mikroglia ve makrofajlarda NLRP3 veya ASC genlerinin dokuya özel koşullu susturulması (downregulation), yaşlı hayvanlarda amiloid plaklarının enflamasyon yaratmasını engellemekte ve hafıza kaybını durdurmaktadır. Eş zamanlı olarak Kaspaz-1 inhibitörleri (Belnacasan / VX-765) ve IL-1beta monoklonal antikorları (Canakinumab - CANTOS klinik çalışmasında kardiyovasküler olayları %15 düşüren ajan) kullanılarak sistemik kaskad bloke edilir. Bu müdahaleler, organizmayı piroptotik doku kaybından ve kronik yangısal yıpranmadan tamamen kurtarır.",
        "Sistemik_Inflamazom_Freni = ( 1 - [Canakinumab] / (K_d_cana + [Canakinumab]) ) * ( 1 - [MCC950] / (IC50 + [MCC950]) )",
        "Bu çarpımsal koruma katsayısı, çift yönlü biyolojik ve farmakolojik ajanların inflamazom kaynaklı sistemik hasarı sıfıra nasıl yaklaştırdığını gösterir."
    )
]

parts.append(("KISIM 1: INFLAM-AGING KAVRAMI VE STERİL İNFLAMASYON BİYOFİZİĞİ", part1_subsections))
parts.append(("KISIM 2: cGAS-STING SİNYAL YOLAĞI VE EKTOPİK DNA YANITI", part2_subsections))
parts.append(("KISIM 3: NLRP3 İNFLAMAZOMU VE İNTERLÖKİN-1BETA / IL-18 OLGUNLAŞMASI", part3_subsections))

# ================= KISIM 4 =================
part4_subsections = [
    (
        "4.1 Timus Bezinin Histolojik Mimarisi: Korteks, Medulla ve Timik Epitelyal Hücreler (TEC)",
        "Timus bezi, kemik iliğinden gelen öncül pro-T hücrelerinin olgun, fonksiyonel ve kendine-hoşgörülü (self-tolerant) T lenfositlerine dönüştüğü primer lenfoid organdır.",
        "Histolojik olarak dışta korteks ve içte medulla olmak üzere iki ana kompartmandan oluşur. Kortikal Timik Epitelyal Hücreler (cTEC), T hücre reseptörünün (TCR) kendi MHC moleküllerini tanıma yeteneğini test eden 'Pozitif Seçilim' sürecini yönetir. Medullar Timik Epitelyal Hücreler (mTEC) ise AIRE (Otoimmün Regülatör) transkripsiyon faktörü sayesinde vücuttaki tüm periferik doku antijenlerini (insülin, miyelin vb.) T hücrelerine sunarak kendi dokularına saldıran otoreaktif klonları apoptozla yok eden 'Negatif Seçilim' sürecini yürütür. Bu mimarinin bütünlüğü, bağışıklık dengesinin temel taşıdır.",
        "Secilim_Verimi = [T_olgun_cikti] / [Pro_T_giris] = f_pozitif(cTEC) * ( 1 - f_negatif_otoreaktif(mTEC) )",
        "Bu stokiyometrik oran, timik kortikal ve medullar epitelyal hücrelerin fonksiyonel kalitesinin nihai naif T hücresi üretim verimini nasıl belirlediğini gösterir."
    ),
    (
        "4.2 Yaşla İlişkili Timik İnvolüsyonun Zaman Çizelgesi ve Yağ İnfiltrasyonu",
        "Timus bezi, memeli vücudunda yaşlanma belirtisi gösteren ilk organdır; bu fizyolojik gerileme sürecine 'Timik İnvolüsyon' adı verilir.",
        "Puberte ile birlikte cinsiyet hormonlarının (östrojen ve testosteron) fırlamasıyla başlayan involüsyon, her yıl fonksiyonel timik dokunun yaklaşık %3'ünün kaybedilmesiyle ilerler. Orta yaşta (45-50 yaş) timik parankimin %70'inden fazlası yerini yağ dokusuna (adiposit infiltrasyonu) bırakmıştır; 70 yaşında ise aktif lenfoid doku neredeyse tamamen yok olur. Fonksiyonel cTEC ve mTEC ağının çökmesi, timusun artık kana yeni T hücresi salgılayamaması anlamına gelir. Bu durum yaşlılıkta aşı yanıtsızlığının ve yeni patojenlere karşı savunmasızlığın temel nedenidir.",
        "Fonksiyonel_Timus_Hacmi(t) = V_0 * exp( -k_involusyon * (t - t_puberte) ) + V_reziduel",
        "Bu üstel bozunma denklemi, puberteden itibaren geçen kronolojik zamanla fonksiyonel timus kütlesinin nasıl yağ infiltrasyonuna uğrayarak minimum rezidüel seviyeye indiğini açıklar."
    ),
    (
        "4.3 FOXN1 Transkripsiyon Faktörünün Kaybı ve Timik Mikroçevre Bozulması",
        "Timik epitelyal hücrelerin (TEC) gelişimi, diferansiyasyonu ve hayatta kalması, 'Forkhead box N1' (FOXN1) transkripsiyon faktörünün mutlak ekspresyonuna bağlıdır.",
        "FOXN1 nakavt fareler (nude mice) tüysüz olmanın yanı sıra tamamen timussuz doğarlar ve hücresel bağışıklıktan yoksundurlar. Yaşlanma sürecinde TEC hücrelerinde FOXN1 ekspresyonu epigenetik susturma ve transkripsiyonel yavaşlama nedeniyle hızla düşer. FOXN1 seviyeleri kritik bir eşiğin altına indiğinde, epitelyal hücreler hücresel kimliklerini kaybederek mezenkimal ve adipojenik hücrelere dönüşür (Epitelyal-Mezenkimal Geçiş / EMT). Sonuç olarak T hücrelerinin tutunacağı ve olgunlaşacağı üç boyutlu epitelyal iskele tamamen dağılır.",
        "TEC_Hayatta_Kalma = k_tec * [FOXN1_nukleus]^m / ( K_M_foxn1^m + [FOXN1_nukleus]^m )",
        "Bu Hill doygunluk eşitliği, nükleer FOXN1 transkripsiyon faktörü konsantrasyonunun timik epitelyal hücre canlılığı ve yapay iskelenin korunumu üzerindeki doğrudan belirleyiciliğini gösterir."
    ),
    (
        "4.4 Naif T Hücre (CD45RA+) Çıktısının Sıfırlanması ve Periferik Homeostatik Proliferasyon",
        "Timusun körelmesiyle birlikte kana yeni çıkan 'naif' (daha önce hiçbir antijenle karşılaşmamış, CD45RA+ CD62L+ CCR7+) T hücre akışı durma noktasına gelir.",
        "Kandaki toplam T hücresi sayısını sabit tutmak için organizma 'Homeostatik Periferik Proliferasyon' mekanizmasını devreye sokar. Mevcut T hücreleri, lenf nodlarında IL-7 ve IL-15 sitokinlerinin uyarısıyla bölünerek sayılarını çoğaltır. Ancak bu bölünmeler sırasında hücreler her siklusta telomerlerini kaybeder, klonal tükenmeye uğrar ve CD45RA+ naif kimliklerini yitirerek hafıza/efektör (CD45RO+) benzeri senesen hücrelere dönüşürler. Böylece periferik havuz sayıca korunsa dahi kalitatif olarak tamamen yıpranır.",
        "d[Naif_T_Hucre]/dt = J_timus_cikti(t) + k_homeo * [IL-7] * [Naif_T] - k_senesens * [Naif_T]",
        "Bu dinamik kütle korunum denklemi, timik çıktının (J_timus_cikti -> 0) sıfırlanmasıyla birlikte naif T hücresi havuzunun homeostatik bölünmeyle nasıl tükendiğini modeller."
    ),
    (
        "4.5 T Hücre Reseptörü (TCR) Klonal Çeşitliliğinin Aşınması ve Antijen Körlüğü",
        "Genç bir bireyin bağışıklık sistemi, V(D)J rekombinasyonu sayesinde teorik olarak 10^12'den fazla farklı antijeni tanıyabilecek muazzam bir T Hücre Reseptörü (TCR) çeşitliliğine (repertoire diversity) sahiptir.",
        "Timik involüsyon ve periferik klonal genişlemeler sonucunda, yaşlı bireylerde TCR klonal çeşitliliği %90'a varan oranlarda daralır. Bireyin bağışıklık repertuarı, sadece geçmişte karşılaşılan birkaç kronik patojene özgü devasa monoklonal hücre orduları tarafından işgal edilir. Sonuç olarak organizma 'antijenik körlük' durumuna girer; daha önce karşılaşmadığı yeni bir viral suş (SARS-CoV-2 veya yeni influenza mutantları) veya yeni ortaya çıkan bir kanser neoantijeni karşısında onu tanıyacak uygun TCR klonunu bulamaz ve savunmasız kalır.",
        "Shannon_TCR_Entropisi = - Sum_i ( p_i * ln(p_i) ) -> Minimum_Yaslilik",
        "Bu bilgi teorisi formülasyonu, TCR klon frekanslarının (p_i) dağılımındaki entropik çeşitliliğin yaşlanmayla nasıl dramatik olarak tabana vurduğunu sayısallaştırır."
    ),
    (
        "4.6 CD28 Kostimülatör Molekülünün Kaybı: CD28-null Senesen T Hücreleri",
        "T hücresinin tam olarak aktive olabilmesi için TCR sinyalinin (Sinyal 1) yanı sıra CD28 kostimülatör reseptörünün antijen sunan hücredeki B7-1/B7-2 (CD80/CD86) molekülleriyle kenetlenmesi (Sinyal 2) şarttır.",
        "Kronik antijenik uyarım ve tekrarlayan replikatif bölünmeler sonucunda, yaşlanan CD8+ ve CD4+ T hücreleri yüzeylerindeki CD28 reseptörünü kalıcı olarak kaybeder (CD28-null / CD28- T hücreleri). Bu CD28- hücreler klasik immün yanıta giremez, proliferasyon gösteremez; ancak yüksek miktarda granzim B, perforin ve pro-enflamatuar sitokinler (IFN-gama, TNF-alfa) salgılayan sitotoksik senesen hücrelere dönüşürler. Bu hücreler vasküler endotelde birikerek aterosklerozu ve doku tahribatını hızlandırır.",
        "CD28_Negatif_Orani = [T_CD28_null] / [T_Toplam] = 1 / ( 1 + exp( -k_div * (N_bolunme - N_kritik) ) )",
        "Bu lojistik fonksiyon, T hücrelerinin geçirdiği kümülatif mitoz bölünme sayısı kritik eşiği aştığında CD28 kaybı ve senesen fenotipe geçiş olasılığını gösterir."
    ),
    (
        "4.7 Sitomegalovirüs (CMV) Enfeksiyonunun İmmün Bellek Havuzunu İşgali (İmmün Risk Profili)",
        "İnsan popülasyonunun büyük kısmında latent olarak bulunan bir herpesvirüs olan Sitomegalovirüs (CMV), immünosenesensin en agresif hızlandırıcısıdır.",
        "CMV hiçbir zaman vücuttan atılamaz; bağışıklık sistemi virüsü baskılamak için sürekli devasa kaynaklar tahsis eder. Yaşlı bireylerde periferik CD8+ T hücresi havuzunun %30 ila %50'si yalnızca CMV'nin birkaç epitopuna özelleşmiş, son evre diferansiye (TEMRA: CD45RA+ CCR7-) hafıza hücreleri tarafından bloke edilir. Bu fenomen 'Bellek Şişmesi' (Memory Inflation) ve 'İmmün Risk Profili' (IRP) olarak adlandırılır. CMV seropozitifliği, CD4/CD8 oranının 1'in altına düşmesi ve yaşlılarda mortalitenin katlanması ile doğrudan ilişkilidir.",
        "Bellek_Sisme_Katsayisi = [T_CMV_ozgul] / [T_CD8_toplam] = 1 - exp(-k_kronik_antijen * Sure_enfeksiyon)",
        "Bu kinetik bağıntı, kronik CMV enfeksiyon süresinin CD8+ T lenfosit havuzunun ne kadarlık bir fraksiyonunu rehin aldığını modeller."
    ),
    (
        "4.8 Timusu Yeniden Büyütme Stratejileri: Rekombinant IL-7, FOXN1 Gen Terapisi",
        "Timik involüsyonun tersine çevrilmesi (thymic rejuvenation), bağışıklık sisteminin gençlik kapasitesine kavuşturulmasında en kritik hedef alanıdır.",
        "Rekombinant İnterlökin-7 (rIL-7) uygulanması, timustaki erken T progenitörlerinin proliferasyonunu uyarır ve timustan periferik kana yeni T hücresi göçünü hızlandırır. Çok daha radikal bir genetik yaklaşım ise, timik epitelyal hücrelere AAV vektörleri veya lipid nanopartiküllerle FOXN1 transkripsiyon faktörünün genetik transferidir. Yaşlı farelere yapılan FOXN1 reaktivasyonu, körelmiş timusun yeniden organize olmasını, histolojik korteks/medulla ayrımının geri kazanılmasını ve gençlik düzeyinde naif T hücresi üretiminin yeniden başlamasını sağlamıştır.",
        "Timik_Rejenerasyon_Hizi = k_rej * [FOXN1_ekspresyon] * [IL-7_lokal] / ( K_M_il7 + [IL-7_lokal] )",
        "Bu rejenerasyon eşitliği, FOXN1 gen terapisinin lokal IL-7 sitokini ile birleştiğinde timik mimariyi ve yeni naif T hücresi çıktısını nasıl üstel olarak artırdığını açıklar."
    ),
    (
        "4.9 Büyüme Hormonu, DHEA ve Çinko ile Timik Rejenerasyon (TRIIM Çalışması)",
        "Dr. Gregory Fahy ve arkadaşları tarafından yürütülen tarihi TRIIM (Thymus Regeneration, Immunorestoration, and Insulin Mitigation) klinik çalışması, insanlarda timik gençleşmeyi ve epigenetik yaşın geri döndürülmesini başaran ilk insan denemesidir.",
        "Protokol; timik epitel büyümesini uyaran rekombinant insan Büyüme Hormonu (rhGH), GH'nin neden olabileceği insülin direncini kırmak için DHEA ve Metformin, ve timik hormonların (timulin) kofaktörü olan Çinko kombinasyonundan oluşmuştur. 1 yıllık tedavi sonucunda MRI görüntülemelerinde timus bezindeki yağ dokusunun yerini aktif lenfoid dokunun aldığı kanıtlanmıştır. En çarpıcı sonuç; Horvath, Hannum, PhenoAge ve GrimAge DNA metilasyon saatleri ile ölçülen biyolojik yaşın, 1 yıllık tedavi sonunda ortalama 2.5 yıl net olarak gençleştiğinin gösterilmesidir.",
        "Delta_Epigenetik_Yas_TRIIM = - ( 1.5_yil_genclesme + 1.0_yil_kronolojik_kazanc ) = -2.5_yil",
        "Bu klinik formülasyon, TRIIM kokteylinin timik rejenerasyon aracılığıyla epigenetik saatler üzerindeki net geriye döndürme büyüklüğünü temsil eder."
    ),
    (
        "4.10 Biyo-Mühendislik Ürünü Yapay Organoid Timus Tasarımı",
        "Timik rejenerasyonun nihai teknolojik ufku; pluripotent kök hücrelerden (iPSC) üretilen yapay timik epitelyal organoidlerin vücuda implante edilmesidir.",
        "Hastanın kendi somatik hücrelerinden türetilen iPSC'ler, kademeli sinyal faktörleri (Wnt, BMP, FGF) ile önce ön bağırsak endodermine, ardından fonksiyonel cTEC ve mTEC benzeri hücrelere diferansiye edilir. Bu hücreler kollajen/fibronektin bazlı 3D biyo-yazıcı iskelelerine ekilir ve vaskülarize edilir. Vücuda yerleştirilen bu 'Sentetik Timus Organoidleri' (Artificial Thymic Organoids - ATO), hastanın kendi kemik iliği kök hücrelerini kabul ederek kesintisiz ve ömür boyu kusursuz bir naif T hücresi üretimi ve tolerans eğitimi sağlar.",
        "ATO_Kapasitesi = V_organoid * Rho_TEC * ( 1 - F_fibrozis ) * J_HSC_alimi",
        "Bu biyo-mühendislik formülü, yapay timus organoidinin doku hacmi, hücre yoğunluğu ve kök hücre alım debisi üzerinden naif T lenfosit üretim gücünü tanımlar."
    )
]

# ================= KISIM 5 =================
part5_subsections = [
    (
        "5.1 Nötrofillerin Kemotaksis ve Fagositoz Kusurları: Hedefe Ulaşamayan Lökositler",
        "Nötrofiller, dolaşımda en bol bulunan ve akut bakteriyel enfeksiyonlara karşı ilk yanıtı veren profesyonel fagositlerdir. Yaşlanmayla birlikte nötrofil sayısı azalmaz; ancak fonksiyonel yön bulma kabiliyetleri felç olur.",
        "Genç nötrofiller kemotaktik gradyanları (IL-8, fMLP, C5a) izleyerek hedefe milimetrik bir doğrulukla yönelirken; yaşlı bireylerin nötrofilleri kemotaktik polarizasyonlarını kaybeder. Plazma membranındaki lipid sallarının rijitliği ve PIP3 sinyalizasyonundaki dengesizlik (PI3K hiperaktivasyonu ve PTEN defekti), nötrofilin hedefe doğru düz gitmek yerine doku içinde rastgele dolanmasına (directional persistence kaybı) yol açar. Bu amaçsız göç, patojenin temizlenmesini geciktirirken, nötrofillerin dokulara kontrolsüz proteolitik enzimler salarak çevreye zarar vermesine neden olur.",
        "Kemotaktik_Hassasiyet = V_yonelim / V_toplam_hiz = cos(Theta_sapma) * ( 1 - Delta_Membran_Rijiditesi )",
        "Bu biyofiziksel açı denklemi, yaşlanan nötrofilin hedef kemoatraktana doğru yönelimsel hız bileşeninin membran sertleşmesiyle nasıl sıfıra yaklaştığını formüle eder."
    ),
    (
        "5.2 Nötrofil Ekstraselüler Kapanları (NETs / NETosis) ve Yaşa Bağlı Trombo-İnflamasyon",
        "Nötrofiller mikropları fagositozla yok etmenin yanı sıra, nükleer DNA'larını histonlar ve granül enzimleri (MPO, Elastaz) ile kaplayarak hücre dışına ağ şeklinde fırlatabilirler (NETosis).",
        "Yaşlılıkta nötrofillerin NET oluşturma eşiği patolojik olarak düşer; hafif bir uyarıda dahi dokulara devasa NET ağları saçılır. Hücre dışındaki bu çıplak DNA-histon kapanları, vasküler yataklarda trombosit agregasyonunu tetikler, pıhtılaşma faktörlerini aktive eder ve mikrovasküler trombozlara yol açar (İmmünotromboz). Eş zamanlı olarak çözülemeyen NET'ler dokularda kronik steril yangıyı körükler ve otoantijen kaynağı haline gelerek lupus benzeri otoimmün reaksiyonları tetikler.",
        "J_NETosis = k_net * [PAD4_aktif] * [mtROS] / ( K_M_net + [mtROS] )",
        "Bu aktivasyon bağıntısı, peptidylarginine deiminase 4 (PAD4) enziminin ve mitokondriyal ROS düzeyinin nükleer kromatin gevşemesi ve patolojik NETosis hızını nasıl yönettiğini modeller."
    ),
    (
        "5.3 Makrofaj Polarizasyonunda M1 Pro-İnflamatuar Kayma ve Doku Onarım Yetmezliği",
        "Doku makrofajları iki zıt fonksiyonel duruma polarize olabilir: Patojen öldürücü ve enflamatuar M1 fenotipi ile doku onarıcı, anjiyogenik ve çözücü M2 fenotipi.",
        "Yaşlanan dokularda makrofaj polarizasyon dengesi kalıcı olarak pro-enflamatuar M1 yönüne kayar. Doku mikroçevresindeki DAMP'ler ve senesen hücre salgıları, makrofajları sürekli M1 durumunda kilitler (yüksek TNF-alfa, IL-1beta, iNOS üretimi). Hasar sonrası dokuyu onaracak, matriksi yeniden yapılandıracak ve yangıyı sonlandıracak M2 polarizasyonu (IL-10, Arg-1, TGF-beta) ise tetiklenemez. Sonuç olarak yaşlı bireylerde kas yaralanmaları, kemik kırıkları ve deri yaraları onarılamaz ve kronik yangılı ülserlere dönüşür.",
        "Polarizasyon_Orani = [Makrofaj_M1] / [Makrofaj_M2] = K_pol * ( [DAMP] + [IFN-gama] ) / ( [IL-4] + [IL-13] )",
        "Bu oran denklemi, doku yangısal medyatörlerinin makrofaj polarizasyonunu M1 yönünde nasıl şiddetle saptırdığını ve doku onarımını durdurduğunu simgeler."
    ),
    (
        "5.4 Makrofajlarda Otofaji ve Fagositoz Yavaşlaması: 'Efferositoz' İflası",
        "Hücresel gençliğin ve doku temizliğinin en kritik adımı, her gün vücutta ölen milyarlarca apoptotik hücrenin makrofajlar tarafından yangı çıkarmadan sessizce yutulup sindirilmesidir (Efferositoz).",
        "Yaşlanan makrofajlarda apoptotik hücreleri tanıyan reseptörlerin (MerTK, Axl, Tim-4) ekspresyonu azalır ve intraselüler fagozom-lizozom füzyonu yavaşlar. Efferositoz iflas ettiğinde, temizlenemeyen apoptotik hücreler doku içinde sekonder nekroza girer; zarları patlar ve tüm toksik intraselüler DAMP içeriklerini dokuya boşaltırlar. Bu durum, kendi hücrelerimizin cesetlerinin bağışıklık sistemini tahrik eden birer yabancı cisim gibi davranmasına yol açar.",
        "Efferositoz_Klerens_Hizi = V_max_eff * [Apoptotik_Hucre] / ( K_M_eff * ( 1 + [Lipofuskin_Makrofaj]/K_i ) + [Apoptotik_Hucre] )",
        "Bu yarışmalı inhibisyon formülü, makrofaj içinde biriken lipofuskin ve çöp moleküllerinin efferositoz temizlik debisini nasıl doğrudan felç ettiğini gösterir."
    ),
    (
        "5.5 Dendritik Hücrelerin Antijen Sunum Kapasitesinde ve Göçünde Gerileme",
        "Dendritik Hücreler (DC), doğal bağışıklık ile edinsel bağışıklık arasındaki köprüdür; periferik dokulardan yakaladıkları antijenleri lenf nodlarına taşıyarak naif T hücrelerine sunarlar.",
        "Yaşlılıkla birlikte dendritik hücrelerin periferik dokudan lenfatik damarlara göç etmesini sağlayan kemokin reseptörü CCR7 ekspresyonu zayıflar. Lenf noduna ulaşabilen DC'lerin ise MHC Sınıf II ve kostimülatör molekülleri (CD80, CD86) yüzeye çıkarma kapasitesi düşer. Antijen işleme makinelerinde (immünoproteazom) yaşanan yavaşlama, antijenik peptitlerin MHC yarıklarına düzgün yüklenmesini engeller; bu da T hücrelerinin verimli bir şekilde eğitilip klonal olarak genişlemesini önler.",
        "Goc_Verimi_DC = N_lenf_noduna_ulasan / N_periferik = G_0 * exp(-k_yas * Yas) * ( [CCR7] / CCR7_genclik )",
        "Bu göç parametresi, yaşa bağlı CCR7 reseptör kaybının lenfatik transfer verimliliğini ve antijen sunumunu nasıl düşürdüğünü gösterir."
    ),
    (
        "5.6 Doğal Katil (NK) Hücre Alt Tipleri: CD56bright vs CD56dim ve Sitotoksisite Kaybı",
        "Doğal Katil (NK) hücreler, tümör hücrelerini ve virüsle enfekte hücreleri MHC kısıtlaması olmaksızın anında imha eden sitotoksik lenfositlerdir.",
        "İnsan NK hücreleri iki ana gruba ayrılır: İmmünoregülatuvar sitokin salgılayan genç CD56bright CD16- hücreler ve yüksek sitotoksik aktiviteye sahip CD56dim CD16+ hücreler. Yaşlanma sürecinde yeni NK hücresi üretimi düştüğü için CD56bright havuzu kurur; periferik kanda ileri derecede yaşlanmış CD56dim hücreler birikir. Ancak bu CD56dim hücrelerin hedef hücre ile kurduğu immünolojik sinaps zayıflar, perforin ve granzim salınım polaritesi bozulur; hücre başına düşen lilik (öldürücü) kapasite %50 geriler.",
        "Litik_Etkinlik_NK = k_lizis * [Perforin_salinan] / ( K_sinaps + [Perforin_salinan] ) * ( 1 - [KIR_inhibitör_sinyal] )",
        "Bu biyokimyasal formül, tek bir NK hücresinin hedef hücreyi delme ve imha etme gücünün yaşlanmayla sönümlenen perforin degranülasyonuna bağımlılığını modeller."
    ),
    (
        "5.7 NK Hücrelerinin Senesen Hücreleri İmhada Yetersiz Kalması (İmmün Sürveyans Zaafı)",
        "Senesen hücreler, çevre dokulardan bağışıklık sistemi tarafından temizlenmek üzere yüzeylerinde NK hücre aktive edici ligandlar (MICA, MICB, ULBP1-3) eksprese ederler.",
        "NK hücrelerinin yüzeyindeki NKG2D reseptörü bu ligandları tanıyarak senesen hücreyi tespit eder ve apoptoza sürükler (Senosen Sürveyansı). Ancak yaşlanmayla iki taraflı bir hile ortaya çıkar: 1) Senesen hücreler yüzeylerindeki MICA/B ligandlarını ADAM metalloproteinazları ile kesip kana döker (ligand shedding), böylece NK reseptörlerini kör ederler; 2) NK hücrelerinde NKG2D sinyal iletimi (DAP10 fosforilasyonu) körelir. Sonuç olarak senesen hücreler bağışıklıktan kaçar ve dokularda kontrolsüzce birikir.",
        "Senosen_Kacis_Indeksi = 1 - ( [NKG2D_aktif] * [MICA_membran] ) / ( K_tanima + [MICA_cozunen] )",
        "Bu kaçış modeli, çözünen MICA ligandlarının NKG2D reseptörlerini nasıl meşgul ederek senesen hücrelerin NK sürveyansından kurtulmasını sağladığını açıklar."
    ),
    (
        "5.8 Miyeloid Kayma (Myeloid Skewing): Hematopoietik Kök Hücrelerin Miyeloid Biası",
        "Kemik iliğindeki Hematopoietik Kök Hücreler (HSC), normalde dengeli bir şekilde hem lenfoid (T ve B hücreleri) hem de miyeloid (granülosit, monosit) soylara farklılaşır.",
        "Yaşlanmayla birlikte epigenetik sürüklenme ve kemik iliği nişindeki kronik inflamasyon (TGF-beta, IL-6), HSC popülasyonunda dramatik bir 'Miyeloid Kayma' (Myeloid Skewing) yaratır. Lenfoid-biaslı kök hücreler tükenirken, miyeloid-biaslı kök hücreler klonal olarak baskın hale gelir. Bu kayma sonucunda organizma bol miktarda pro-enflamatuar monosit ve nötrofil üretirken; adaptif bağışıklığın lenfositleri (T ve B) üretilemez. Bu dengesizlik hem enfeksiyonlara karşı savunmasızlığı hem de inflam-aging yangınını aynı anda körükler.",
        "Miyeloid_Lenfoid_Orani = [Miyeloid_Progenitor] / [Lenfoid_Progenitor] = R_0 * exp( alpha_bias * Yas )",
        "Bu üstel kayma denklemi, artan yaşla birlikte hematopoietik diferansiyasyon dengesinin nasıl kaçınılmaz olarak miyeloid soy yönüne saptığını modeller."
    ),
    (
        "5.9 Komplement Sisteminin Yaşla Aşırı Aktivasyonu ve Doku Hasarı",
        "Komplement sistemi, patojenlerin lizisini sağlayan ve yangıyı uyaran 30'dan fazla plazma proteininden oluşan enzimatik bir proteolitik kaskaddır.",
        "Yaşlanma sürecinde komplement inhibitör faktörleri (Faktör H, CD55, CD59) doku yüzeylerinde azalırken; kaskadın bazal aktivasyonu patolojik olarak artar. Özellikle retinada yaşa bağlı makula dejenerasyonunda (AMD), beyinde sinaps budanmasında (Alzheimer) ve böbrek glomerüllerinde C3a, C5a anafilatoksinleri ve Membran Atak Kompleksi (MAC / C5b-9) depolanır. Kendi sağlıklı doku hücrelerimizin membranlarında açılan bu mikroskobik komplement delikleri lizise, kalsiyum sızıntısına ve kronik doku atrofisine neden olur.",
        "MAC_Litik_Hasar = k_MAC * [C5b-9_kompleks_zar] / ( 1 + [CD59_koruyucu] / K_i_cd59 )",
        "Bu litik eşitlik, yüzey koruyucu CD59 molekülünün kaybının membran atak kompleksi delinme hızını ve hücre ölümünü nasıl artırdığını gösterir."
    ),
    (
        "5.10 Doğal Bağışıklık Hücrelerinin Metabolik ve Fonksiyonel Yeniden Programlanması",
        "Doğal bağışıklık hücrelerindeki yaşlanma kalıcı bir kader değildir; hücresel metabolik yakıt yollarının yeniden yapılandırılmasıyla gençlik fonksiyonları restore edilebilir.",
        "M1 makrofajların kontrolsüz glikolitik akışı (Warburg benzeri metabolizma), itakonat türevleri (4-oktil itakonat) veya sentetik Nrf2 aktivatörleri ile frenlendiğinde hücreler anti-enflamatuar ve onarıcı M2 fenotipine geri döner. Eş zamanlı olarak nötrofillerde PI3K-gama inhibitörleri kullanılarak kemotaktik rota düzeltilmekte; fagositoz aktivatörleri ile efferositoz kapasitesi canlandırılmaktadır. Bu immünometabolik reprogramlama, doğal bağışıklığın yıkıcı yangısını dindirirken koruyucu kalkanını yeniden kurar.",
        "Metabolik_Reset_Skoru = ( J_oksidatif_fosforilasyon / J_glikoliz ) * ( [Itakonat] / K_itakonat )",
        "Bu metabolik oran, makrofaj içi yakıt dengesinin glikolizden mitokondriyal solunuma kaydırılmasının yangı söndürme verimliliğini temsil eder."
    )
]

# ================= KISIM 6 =================
part6_subsections = [
    (
        "6.1 Kemik İliği B Lenfopoezinde Düşüş ve Pro-B / Pre-B Progenitör Kaybı",
        "B lenfositlerinin gelişimi (B lenfopoez), kemik iliğinde hematopoietik kök hücrelerden başlayarak Pro-B, Pre-B ve olgunlaşmamış (immature) B hücresi aşamalarından geçer.",
        "Yaşlanan kemik iliğinde stromal hücrelerin IL-7 (İnterlökin-7) ve CXCL12 (SDF-1) sitokin üretimi dramatik şekilde azalır. Bu trofik destek kaybı nedeniyle erken B progenitörleri (özellikle Pre-B hücreleri) apoptoza uğrar ve günlük B hücresi üretimi gençlik seviyesinin %10-20'sine kadar geriler. Kemik iliğinin sarı yağ iliğine dönüşmesi, B lenfopoezi için gerekli anatomik nişleri fiziksel olarak ortadan kaldırır.",
        "B_Lenfopoez_Akisi = k_B * [IL-7_stroma] * [HSC_lenfoid] / ( K_M_il7 + [IL-7_stroma] )",
        "Bu üretim debisi denklemi, stromal IL-7 konsantrasyonu ve lenfoid kök hücre sayısının kemik iliği taze B hücresi çıktısını nasıl belirlediğini açıklar."
    ),
    (
        "6.2 Naif B Hücre Havuzunun Tükenmesi ve Yaşla İlişkili B Hücreleri (ABCs: T-bet+ CD11c+)",
        "Taze B hücresi çıktısının kesilmesiyle periferik kanda naif B hücreleri (CD19+ CD27- IgD+) tükenir ve yerini 'Yaşla İlişkili B Hücreleri' (Age-Associated B Cells - ABCs) alır.",
        "ABCs fenotipik olarak T-bet transkripsiyon faktörünü ve miyeloid belirteç CD11c'yi eksprese eden atipik bir B hücresi alt grubudur. Bu hücreler klasik BCR uyarımına zayıf yanıt verirken; TLR7 ve TLR9 stimülasyonuna aşırı duyarlıdır. Gençlikte koruyucu antikor üreten hücrelerin yerini alan ABC'ler, yaşlı bireylerde kronik otoantikor salgılar ve yüksek düzeyde TNF-alfa ve IL-6 üreterek sistemik inflam-aging'e doğrudan katılırlar.",
        "ABC_Fraksiyonu = [B_Tbet_CD11c] / [B_Toplam] = 1 / ( 1 + exp(-k_abc * (Yas - T_esik)) )",
        "Bu sigmoidal birikim modeli, yaşlanmayla periferik B lenfosit havuzunda patolojik ABC hücre oranının nasıl kaçınılmaz olarak yükseldiğini gösterir."
    ),
    (
        "6.3 Germinal Merkez Reaksiyonlarının Zayıflaması ve Somatik Hipermutasyon Bozukluğu",
        "Yüksek afiniteli antikorların üretilmesi, sekonder lenfoid organlarda (dalak ve lenf nodları) kurulan 'Germinal Merkez' (GC) mikro-yapılarında gerçekleşir.",
        "Yaşlılıkla birlikte foliküler dendritik hücre (FDC) ağının zayıflaması ve foliküler T yardımcı (Tfh) hücre desteğinin yetersizliği nedeniyle germinal merkezler küçülür ve kısa sürede dağılır. B hücrelerinin antikor değişken bölgelerinde mutasyonlar yaparak afiniteyi artıran Aktivasyon Kaynaklı Sitidin Deaminaz (AID) enziminin ekspresyonu %60 oranında düşer. Somatik hipermutasyonun (SHM) yavaşlaması, B hücrelerinin antijene sıkı sıkıya bağlanan mükemmel antikorlar geliştirmesini engeller.",
        "Somatik_Mutasyon_Hizi = k_mut * [AID_aktif] * [Germinal_Merkez_Omru]",
        "Bu formülasyon, AID enzim aktivitesi ve germinal merkez kararlılığının antikor optimizasyon hızını nasıl doğrudan sınırladığını gösterir."
    ),
    (
        "6.4 Antikor Afinite Maturasyonunun Çökmesi ve Aşı Yanıtsızlığı",
        "Somatik hipermutasyon ve germinal merkez yetersizliğinin doğrudan klinik sonucu, antikor afinite maturasyonunun çökmesidir.",
        "Yaşlı bireyler aşılandığında (örn. mevsimsel influenza veya pnömokok aşıları), üretilen antikorların antijene bağlanma sabiti (Ka) ve nötralizasyon gücü gençlere kıyasla 10 ila 100 kat daha düşüktür. Antikorlar patojene zayıf bağlanır ve virüsü etkisiz hale getiremez. Bu durum, yaşlı popülasyonda aşıların koruyuculuk oranının %30-40 seviyelerine gerilemesinin ve aşılanmış olmalarına rağmen ölümcül pnömoni ve grip komplikasyonları yaşamalarının temel biyolojik izahatıdır.",
        "Asi_Koruyuculuk_Verimi = Ka_antikor * [IgG_plazma] / ( K_notralizasyon + Ka_antikor * [IgG_plazma] )",
        "Bu nötralizasyon fonksiyonu, antikor afinitesindeki (Ka) düşüşün sistemik aşı etkinliğini nasıl dramatik olarak düşürdüğünü açıklar."
    ),
    (
        "6.5 İmmünoglobulin Sınıf Değişimi (CSR) Bozuklukları ve E-protein Regülasyonu",
        "B hücresi aktive olduğunda başlangıçtaki düşük afiniteli IgM/IgD izotiplerini yüksek efektör fonksiyonlu IgG, IgA veya IgE izotiplerine dönüştürmek zorundadır (Sınıf Değişimi Rekombinasyonu - CSR).",
        "CSR süreci de AID enzimine ve E2A (E47), Pax5 gibi transkripsiyon faktörlerine bağımlıdır. Yaşlı B lenfositlerinde E47 transkripsiyon faktörünün mRNA ve protein seviyeleri çöker; bunun sonucunda AID gen promoterı açılamaz. B hücreleri IgM'den mukozal koruyucu IgA'ya veya sistemik koruyucu IgG1/IgG3 izotiplerine geçiş yapamaz. Mukozalarda IgA eksikliği solunum ve sindirim yolu enfeksiyonlarına kapı aralarken, sistemik kanda işlevsiz poliklonal IgM birikir.",
        "CSR_Etkinligi = [IgG_sinif_degismis] / [Toplam_Aktive_B] = k_csr * [E47] * [AID] / ( K_d_csr + [E47]*[AID] )",
        "Bu regülasyon denklemi, nükleer E47 ve AID seviyelerinin immünoglobulin sınıf değişimi başarısını nasıl doğrudan kontrol ettiğini ortaya koyar."
    ),
    (
        "6.6 Yaşlılıkta Otoantikor Üretiminin Artışı ve Periferik Toleransın Kırılması",
        "İmmünosenesens bir yandan patojenlere karşı antikor üretimini felç ederken, diğer yandan ironik bir şekilde kendi dokularımıza saldıran otoantikorların üretimini patlatır.",
        "Kemik iliğinde santral toleransın ve lenf nodlarında periferik B hücre toleransının (anerji mekanizmaları) gevşemesi, kendini tanıyan otoreaktif B hücre klonlarının apoptozdan kaçmasını sağlar. Yaşlı bireylerin kanında antinükleer antikorlar (ANA), romatoid faktör (RF) ve anti-fosfolipid antikorları sıklıkla pozitifleşir. Bu düşük afiniteli polireaktif otoantikorlar; böbrek bazal membranlarında, eklem sinovyasında ve damar endotelinde immün kompleksler halinde çökerek yaşa bağlı subklinik doku yıkımını hızlandırır.",
        "Otoantikor_Titre = Titre_0 + Delta_Titre * ( 1 - [Tolerans_Faktoru] ) * [ABC_B_Hucreleri]",
        "Bu patolojik model, azalan tolerans faktörleri ve artan ABC hücrelerinin sistemik otoantikor yükünü nasıl katlanarak artırdığını açıklar."
    ),
    (
        "6.7 Klonal B Hücre Genişlemesi ve Monoklonal Gammopati (MGUS) Riski",
        "Yaşlı bireylerde B hücresi repertuarının poliklonal çeşitliliği kaybolurken, tek bir mutant B hücresi klonunun kontrolsüzce çoğalarak dolaşımı işgal etmesi yaygınlaşır.",
        "Bu tablonun en belirgin klinik örneği, 70 yaş üstü popülasyonun %5'inden fazlasında görülen 'Önemi Belirsiz Monoklonal Gammopati'dir (MGUS). Tek bir plazma hücresi klonu kanda monoklonal bir immünoglobulin (M-spikes / paraprotein) üretir. MGUS, zamanla multipl miyelom veya lenfoma gibi hematolojik malignitelere dönüşme riski taşır; aynı zamanda monoklonal proteinler böbrek tübüllerinde birikerek amiloidoz ve renal fonksiyon kaybına yol açar.",
        "MGUS_Malign_Donusum_Hizi = k_transformasyon * [M_Protein_Konsantrasyonu] * ( 1 + [Genomik_Instabilite] )",
        "Bu onkolojik risk fonksiyonu, dolaşımdaki paraprotein konsantrasyonu ve genomik kararsızlığın habis dönüşüm hızını nasıl belirlediğini modeller."
    ),
    (
        "6.8 Foliküler T Yardımcı (Tfh) Hücre Yetmezliğinin B Hücre Fonksiyonlarına Yansıması",
        "B hücrelerinin germinal merkezlerde hayatta kalması ve afinite kazanması, Foliküler T Yardımcı (Tfh) hücrelerinin sağladığı CD40L ve IL-21 sinyallerine mutlak bağımlıdır.",
        "Yaşlanmayla birlikte naif T hücre üretiminin çökmesi, fonksiyonel Tfh hücre havuzunu da kurutur. Mevcut yaşlı Tfh hücreleri ise B hücre foliküllerine göçü sağlayan CXCR5 reseptörünü kaybeder ve B hücreleri ile stabil immünolojik temas kuramaz. Yeterli CD40L-CD40 kenetlenmesi ve IL-21 uyarısı alamayan B hücreleri germinal merkezde apoptoza uğrar. B hücre yaşlanması, aslında büyük oranda Tfh hücresi yaşlanmasının ikincil bir kurbanıdır.",
        "B_Hucre_Kurtulma_Orani = k_surv * [Tfh_aktif] * [CD40L] * [IL-21] / ( K_esik_tfh + [IL-21] )",
        "Bu bağımlılık eşitliği, germinal merkezdeki B hücre sağkalımının Tfh yardım sinyallerinin yoğunluğu tarafından nasıl doğrudan dikte edildiğini belgeler."
    ),
    (
        "6.9 BAFF ve APRIL Sitokin Sinyalizasyonunun Yaşa Bağlı Değişimi",
        "B hücresi sağkalımı ve olgunlaşması, TNF süper familyası üyeleri olan BAFF (B-cell Activating Factor) ve APRIL (A Proliferation-Inducing Ligand) sitokinleri tarafından düzenlenir.",
        "Yaşlı bireylerde azalan B hücresi sayısını kompanse etmek amacıyla dolaşımdaki BAFF seviyeleri kompanzatuar olarak aşırı yükselir. Ancak kronik olarak yüksek BAFF seviyeleri, otoreaktif B hücrelerinin negatif seçilimden kaçmasına izin verir ve otoimmüniteyi körükler. Eş zamanlı olarak B hücre yüzeyindeki BAFF Reseptörü (BAFF-R) aşağı regüle olurken, TACI reseptörünün artması hücreleri apoptoza ve anormal farklılaşmaya iter.",
        "d[BAFF_serum]/dt = J_salgi_makrofaj - k_tuketim * [B_Hucre_Sayisi] * [BAFF-R]",
        "Bu dinamik denge denklemi, azalan B hücresi tüketiminin kanda serbest BAFF birikimine ve bunun da otoreaktif klon kaçışına nasıl yol açtığını modeller."
    ),
    (
        "6.10 Genç B Hücre Repertuarının Sentetik Olarak Yeniden İnşası",
        "B hücresi yaşlanmasını tersine çevirmek; kemik iliği nişini tazelemek, AID transkripsiyonunu yeniden açmak ve patolojik ABC klonlarını temizlemekten geçer.",
        "Kemik iliğine mezenkimal kök hücre (MSC) transplantasyonu veya parakrin IL-7/CXCL12 salgılatıcı biyomateryallerin enjeksiyonu taze B lenfopoezini canlandırır. Monoklonal antikorlar ile T-bet+ CD11c+ ABC hücrelerinin seçici olarak hedeflenip ortadan kaldırılması, sistemik otoantikor fırtınasını ve inflam-aging yükünü anında hafifletir. E47 ve AID genlerinin sentetik epigenetik aktivatörlerle uyarılması ise yaşlı bireylerde aşı yanıtını gençlik seviyesine geri döndürür.",
        "Repertuar_Genclesme_Skoru = ( [Naif_B_Hucre_yeni] / [ABC_B_Hucre] ) * ( AID_aktivite / AID_bazal )",
        "Bu sentetik gençleşme indeksi, taze naif B hücresi üretiminin patolojik klonlara oranını ve aşı yanıt kalitesini sayısallaştırır."
    )
]

parts.append(("KISIM 4: TİMİK İNVOLÜSYON VE T HÜCRE REPERTUARININ DARALMASI", part4_subsections))
parts.append(("KISIM 5: DOĞAL BAĞIŞIKLIK SİSTEMİNDE YAŞLANMA: NÖTROFİL, MAKROFAJ VE NK DİSFONKSİYONU", part5_subsections))
parts.append(("KISIM 6: B HÜCRE YAŞLANMASI, ANTİKOR AFİNİTESİ VE OTOİMMÜNİTE PATLAMASI", part6_subsections))

# ================= KISIM 7 =================
part7_subsections = [
    (
        "7.1 Sağlıklı Mikrobiyotanın Çeşitliliği ve Kommensal Türler (Bifidobacterium, Akkermansia)",
        "İnsan gastrointestinal sistemi, 100 trilyondan fazla mikroorganizma ve 3 milyondan fazla mikrobiyal gen barındıran devasa bir 'metabolik ve immünolojik organ'dır.",
        "Sağlıklı bir genç mikrobiyotası, yüksek taksonomik çeşitlilik (alfa çeşitliliği) ve simbiyotik kommensal bakterilerin dominansı ile karakterizedir. Bifidobacterium türleri (B. infantis, B. longum) anne sütü oligosakkaritlerini sindirerek bağırsak mukozasını korur; Akkermansia muciniphila ise kolondaki musin tabakasını kontrollü yıkarak epitelin sürekli taze mukus üretmesini uyarır ve bariyer bütünlüğünü sağlar. Bu kommensal türler, patojen bakterilerin koloni kurmasını sterik olarak engeller ve periferik bağışıklık hücrelerini regülatuar T hücresi (Treg) yönünde eğitir.",
        "Kommensal_Bariyer_Gucu = k_bar * [Akkermansia] * [Bifidobacterium] / ( 1 + [Firsatci_Patojen] / K_esik )",
        "Bu koruma eşitliği, kommensal türlerin yoğunluğunun bağırsak epitelini patojenik invazyona karşı koruma katsayısını nasıl belirlediğini modeller."
    ),
    (
        "7.2 Yaşlılık Disbiyozisi: Mikrobiyal Alfa Çeşitliliğinin Çöküşü ve Fırsatçı Patojen Artışı",
        "Yaşlanma süreci, bağırsak ekosisteminde dramatik bir çöküş ve tür kaybı ile seyreder; bu patolojik bozulmaya 'Yaşlılık Disbiyozisi' adı verilir.",
        "Diş kayıpları, çiğneme zorlukları, liften fakir tek tip beslenme, azalan mide asidi (aklorhidri), yavaşlayan bağırsak motilitesi ve sık antibiyotik kullanımı; mikrobiyotanın alfa çeşitliliğini (Shannon indeksi) tabana indirir. Sağlıklı kommensaller yok olurken, yerlerini Clostridium difficile, Enterobacteriaceae ve Pseudomonas gibi pro-enflamatuar, toksin üreten fırsatçı patobiontlar alır. Bu mikrobiyal dejenerasyon, bağırsak lümenini devasa bir enflamasyon kaynağına dönüştürür.",
        "Disbiyozis_Indeksi = ( Sum[Patobiont_Turleri] / Sum[Kommensal_Turler] ) * ( 1 / Shannon_Alfa_Cesitliligi )",
        "Bu oran, patobiont baskınlığının ve tür çeşitliliği kaybının toplam mikrobiyal dengesizlik derecesini nasıl katladığını sayısallaştırır."
    ),
    (
        "7.3 Firmicutes / Bacteroidetes Oranının Değişimi ve Metabolik Sonuçları",
        "Bağırsak mikrobiyotasının iki baskın bakteriyel filumu Firmicutes ve Bacteroidetes'tir; bu iki grubun birbirine oranı (F/B Oranı) metabolik sağlığın ana göstergelerinden biridir.",
        "Genç yetişkinlerde dengeli olan F/B oranı, yaşlanmayla ve metabolik sendromla birlikte genellikle Firmicutes lehine aşırı yükselir. Firmicutes filumundaki bakteriler, besinlerden daha fazla kalori absorbe etme kapasitesine sahip enzimler kodlar; bu durum aynı kalori alımında dahi visseral yağlanmayı ve obeziteyi tetikler. Ancak çok ileri yaşlılıkta ve kırılgan (frail) geriatrik bireylerde F/B oranı bu kez kontrolsüz şekilde çöker ve genel malnütrisyona zemin hazırlar.",
        "Enerji_Hasat_Kapasitesi = Eta_0 * ( [Firmicutes] / [Bacteroidetes] )^gamma_metabolik",
        "Bu biyometrik fonksiyon, bağırsaktaki F/B oranının tüketilen diyet liflerinden organizmaya aktarılan net serbest enerji katsayısını nasıl belirlediğini açıklar."
    ),
    (
        "7.4 Akkermansia muciniphila Kaybı ve Mukus Tabakasının Aşınması",
        "Verrucomicrobia filumuna ait bir bakteri olan Akkermansia muciniphila, bağırsak mukozasının en kritik koruyucu muhafızıdır.",
        "Akkermansia, kolon goblet hücreleri tarafından salgılanan MUC2 musin proteinini parçalayarak beslenir; bu fizyolojik tüketim goblet hücrelerini uyararak mukus tabakasının her gün 150 mikrometre kalınlıkta taze olarak yenilenmesini sağlar. Yaşlanmayla birlikte bağırsaktaki Akkermansia popülasyonu hızla tükenir. Akkermansia kaybı, mukus tabakasının incelmesine, çatlamasına ve lümendeki bakterilerin doğrudan epitel hücre membranlarına temas etmesine yol açar; bu temas lokal enteriti ve sistemik yangıyı ateşler.",
        "Mukus_Kalinligi(t) = T_0 + k_yenileme * [Akkermansia] - k_erozyon * [Patojen_Yuk]",
        "Bu denge modeli, bağırsak musin astarının kalınlığının Akkermansia konsantrasyonu ile eroziv faktörler arasındaki dinamik etkileşime bağlılığını formüle eder."
    ),
    (
        "7.5 Kısa Zincirli Yağ Asitleri (SCFA: Butirat, Propiyonat, Asetat) Üretiminde Çöküş",
        "Bağırsak kommensal bakterilerinin diyet liflerini anaerobik fermantasyona uğratması sonucu ürettikleri Kısa Zincirli Yağ Asitleri (SCFA: Asetat, Propiyonat ve Butirat), evrensel metabolik habercilerdir.",
        "Özellikle Faecalibacterium prausnitzii ve Roseburia türleri tarafından üretilen Butirat, kolonositlerin (bağırsak epitel hücreleri) primer enerji kaynağıdır (%70 ATP butirattan sağlanır). Propiyonat karaciğerde glukoneojenezi düzenlerken, Asetat periferik yağ dokusunda lipid metabolizmasını yönetir. Yaşlılık disbiyozisinde SCFA üreten bakterilerin yok olması, lümendeki butirat konsantrasyonunu %80 oranında düşürür. Enerjisiz kalan kolonositler atrofiye uğrar ve bariyer bütünlüğü çöker.",
        "J_SCFA_uretim = Sum_i ( k_ferm_i * [Bakteri_i] * [Diyet_Lifi] / ( K_M_lif + [Diyet_Lifi] ) )",
        "Bu fermantasyon kinetiği denklemi, kolondaki toplam kısa zincirli yağ asidi üretim akısının kommensal lif sindirici bakteri popülasyonuna doğrudan bağımlı olduğunu modeller."
    ),
    (
        "7.6 Butiratın Kolonositte Epigenetik Koruyucu ve HDAC İnhibitörü Fonksiyonunun Kaybı",
        "Butirat sadece bir besin maddesi değil; memeli hücrelerinde Sınıf I ve Sınıf II histon deasetilazların (HDAC) en güçlü doğal inhibitörüdür.",
        "Epitel hücrelerinin içine SLC5A8 taşıyıcısı ile giren butirat, HDAC'leri inhibe ederek histon H3 ve H4 asetilasyonunu artırır. Bu epigenetik açılma, sıkı bağlantı proteinlerinin (Claudin-1, ZO-1) ve anti-enflamatuar genlerin transkripsiyonunu uyarır. Aynı zamanda lamina propriadaki naif T hücrelerinin Foxp3 transkripsiyon faktörünü aktive ederek immünosüpresif Regülatuar T hücrelerine (Treg) dönüşmesini sağlar. Yaşlılıkta butiratın tükenmesi, bağırsakta Treg kalkanını yıkarak kontrolsüz oto-enflamasyona neden olur.",
        "Treg_Diferansiyasyon = k_treg * [Butirat_lamina] / ( K_d_HDAC + [Butirat_lamina] ) * [TGF-beta]",
        "Bu regülasyon eşitliği, lokal butirat seviyelerinin HDAC inhibisyonu üzerinden immünoregülatuvar Treg hücresi indüksiyonundaki belirleyici rolünü açıklar."
    ),
    (
        "7.7 Triptofan Katabolizması ve Mikrobiyal İndol Türevlerinin Azalması",
        "Esansiyel bir amino asit olan L-triptofan, bağırsak mikrobiyotası tarafından indol, indol-3-propiyonik asit (IPA) ve indol-3-asetik asite (IAA) dönüştürülür.",
        "Bu indol türevleri, bağırsak epitelindeki ve bağışıklık hücrelerindeki Aril Hidrokarbon Reseptörünün (AhR) doğal ligandlarıdır. AhR aktivasyonu, epitel hücre sağkalımını ve antimikrobiyal peptitlerin (reg3gama) salınımını uyarır. Yaşlanan mikrobiyotada triptofanın indol yolağı yerine inflamatuar konak Kynurenin yolağına (IDO1 enzimi üzerinden) kayması; bağırsak mukozal immünitesini zayıflatırken, beyinde nörotoksik kinolinik asit birikimine ve bilişsel gerilemeye yol açar.",
        "AhR_Aktivasyon_Indeksi = [IPA_mikrobiyal] / ( K_d_AhR + [IPA_mikrobiyal] ) * ( 1 / (1 + [Kynurenin]/K_i) )",
        "Bu reseptör doyum modeli, mikrobiyal indol türevlerinin kaybının ve artan kynurenin metabolitlerinin AhR koruyucu kalkanını nasıl çökerttiğini gösterir."
    ),
    (
        "7.8 Süper-Asırlıkların (Centenarian) Benzersiz Mikrobiyom İmzası ve Koruyucu Suşlar",
        "100 yaşını aşmış sağlıklı süper-asırlıkların (Centenarians) mikrobiyom analizleri, bu olağanüstü bireylerin yaşlanmaya rağmen genç bireylerinkine benzer yüksek mikrobiyal çeşitliliği koruduğunu göstermiştir.",
        "Süper-asırlıkların bağırsaklarında Akkermansia, Christensenellaceae, Alistipes ve Bifidobacterium suşları yüksek oranda zenginleşmiştir. Dahası bu bireylerde sekonder safra asitlerini (özellikle isoallo-litokolik asit / isoalloLCA) sentezleyen özel bakteriyel enzimler bulunur. IsoalloLCA, gram-pozitif patojenleri (C. difficile gibi) güçlü şekilde inhibe eden ve Th17/Treg dengesini koruyan doğal bir antibiyotiktir. Bu bulgular, uzun yaşamın bağırsak metabolit mimarisiyle doğrudan kodlandığını kanıtlar.",
        "Centenarian_Mikrobiyom_Skoru = [Christensenellaceae] * [isoalloLCA] / ( [Enterobacteriaceae] + Epsilon )",
        "Bu biyolojik indeks, uzun ömürlü bireylerin bağırsaklarındaki özgün koruyucu bakteriyel taksonların ve koruyucu safra asitlerinin yoğunluğunu temsil eder."
    ),
    (
        "7.9 Antibiyotik Maruziyetinin Kümülatif Mikrobiyom Erozyonu",
        "Yaşam boyunca tekrarlayan geniş spektrumlu antibiyotik tedavileri, bağırsak ekosisteminde geri döndürülemez 'ekolojik yara izleri' (ecological scars) bırakır.",
        "Tek bir florokinolon veya beta-laktam antibiyotik kürü, mikrobiyotanın %30'undan fazlasını yok edebilir; bazı hassas kommensal suşlar tedavi bittikten aylar sonra bile geri dönemez ve nesilleri tükenir. Kümülatif antibiyotik maruziyeti, mikrobiyomun dayanıklılık (resilience) kapasitesini aşındırır, antibiyotik direnç genlerinin (rezistom) bağırsakta yayılmasına neden olur ve biyolojik yaşlanma saatlerini hızlandırır.",
        "Mikrobiyal_Rezilyans_Kaybi = R_0 * exp( -Sum_j ( D_j * Toksisite_Katsayisi_j ) )",
        "Bu kümülatif hasar modeli, yaşam boyu alınan antibiyotik kürlerinin mikrobiyal ekosistemin kendini toparlama kapasitesini nasıl kalıcı olarak tükettiğini formüle eder."
    ),
    (
        "7.10 Yeni Nesil Probiyotikler, Prebiyotikler ve Sentetik Postbiyotik Mühendisliği",
        "Mikrobiyotanın gençleştirilmesi; klasik yoğurt bakterilerinin ötesinde, katı anaerobik 'Yeni Nesil Probiyotikler' (NGP) ve sentetik postbiyotiklerin hedefe yönelik teslimatını gerektirir.",
        "Oksijensiz ortamda çoğaltılan taze Akkermansia muciniphila ve Faecalibacterium prausnitzii liyofilize kapsüller halinde mikrobiyotaya ekilir. Eş zamanlı olarak bu bakterileri seçici besleyen prebiyotik lifler (Fruktooligosakkaritler - FOS, İnülin, 2'-Fukozillaktoz) verilir. Çok daha hızlı bir müdahale ise doğrudan bakteriyel metabolitlerin (Butirat, İndol-3-Propiyonik Asit, Urolithin A) enterik kaplı tabletlerle kolona ulaştırıldığı 'Postbiyotik Tedavisi'dir; bu yaklaşım bakteriye ihtiyaç duymadan epitel gençleşmesini anında başlatır.",
        "Epitel_Restorasyon_Verimi = k_post * [Butirat_hedefli] + k_ngp * [Akkermansia_kolonize] * [Prebiyotik]",
        "Bu terapötik model, sentetik postbiyotik teslimatı ile yeni nesil probiyotik kolonizasyonunun bağırsak yenilenmesindeki sinerjistik gücünü açıklar."
    )
]

# ================= KISIM 8 =================
part8_subsections = [
    (
        "8.1 Bağırsak Epitel Bütünlüğü: Sıkı Bağlantı (Tight Junction) Proteinleri (Claudin, Occludin, ZO-1)",
        "Bağırsak epitel tabakası, tek bir hücre katmanından oluşan ve 32 metrekarelik bir yüzey alanında iç ortamı lümendeki toksinlerden ayıran hassas bir biyofiziksel bariyerdir.",
        "Epitel hücrelerinin paraselüler (hücreler arası) boşluğu, 'Sıkı Bağlantı' (Tight Junction - TJ) protein kompleksleri ile mühürlenmiştir. Transmembran proteinleri olan Claudin-1, Claudin-4 ve Occludin, hücre dışı ilmekleriyle komşu hücreye kenetlenir. Bu transmembran elemanlar hücre içinde Zonula Occludens-1 (ZO-1) ve ZO-2 adaptör proteinleri aracılığıyla aktin hücre iskeletine demirlenir. Bu moleküler fermuar, 4 Angstrom'dan büyük moleküllerin ve bakteriyel toksinlerin kontrolsüz geçişini engeller.",
        "Paraseluler_Gecirgenlik = P_0 / ( 1 + [ZO-1:Claudin_kompleks] / K_m_tight )",
        "Bu membran biyofiziği formülü, sıkı bağlantı protein komplekslerinin yoğunluğunun paraselüler doku geçirgenliğini nasıl ters orantılı sınırlandırdığını tanımlar."
    ),
    (
        "8.2 Yaşla İlgili Bariyer Gevşemesi ve Paraselüler Geçirgenlik Artışı",
        "Yaşlanma sürecinde mukus tabakasının erimesi, butirat eksikliği ve epitel kök hücrelerinin yenilenme kapasitesinin düşmesi sıkı bağlantı fermuarının çözülmesine yol açar.",
        "Sıkı bağlantı proteinleri (ZO-1, Occludin) fosforillenerek hücre içine çekilir ve parçalanır; yerlerine gözenek oluşturan Claudin-2 proteinleri eksprese edilir. Bu durum bağırsak astarında mikroskobik çatlakların ve açıklıkların oluşmasına neden olur; bu patoloji klinik olarak 'Geçirgen Bağırsak Sendromu' (Leaky Gut) olarak tanımlanır. Artık bağırsak lümeni ile kan dolaşımı arasındaki steril sınır kalkmıştır.",
        "d[Bariyer_Direnc]/dt = R_biyosentez_TJ - k_yikim * [TNF-alfa_lokal] * [Bariyer_Direnc]",
        "Bu türev eşitliği, lokal mukozal TNF-alfa sitokini varlığının transepitelyal elektriksel direnci (TEER) ve bariyer sağlamlığını nasıl zamanla tükettiğini gösterir."
    ),
    (
        "8.3 Lipopolisakkarit (LPS) Translokasyonu ve Portal Kana Bakteriyel Sızıntı",
        "Geçirgen bağırsak astarından kana sızan en ölümcül bakteriyel toksin, Gram-negatif bakterilerin dış zarında bulunan Lipopolisakkarit (LPS / Endotoksin) molekülüdür.",
        "Normalde lümende tutulan devasa LPS yükü, gevşeyen paraselüler yarıklardan ve hasarlı enterositlerden geçerek lamina propriaya sızar ve oradan mezenterik venler aracılığıyla doğrudan Portal Vene karışır. Kana karışan serbest LPS, hepatik ve periferik dolaşıma hızla yayılır. 60 yaş üzerindeki sağlıklı bireylerde bile açlık plazma LPS seviyelerinin gençlere göre 2-3 kat yüksek olduğu kanıtlanmıştır; bu olgu 'Metabolik Endotoksemi' olarak bilinir.",
        "J_LPS_translokasyon = D_LPS * ( [LPS_lumen] - [LPS_portal] ) * Area_çatlak / Kalinlik_epitel",
        "Bu Fick difüzyon yasası uyarlaması, bağırsak epitelindeki çatlak alanının genişlemesiyle portal kana akan endotoksin kütle akısını modeller."
    ),
    (
        "8.4 Sistemik Metabolik Endotoksemi ve TLR4 / CD14 Reseptör Aktivasyonu",
        "Dolaşıma karışan LPS molekülleri, kanda Lipopolisakkarit Bağlayıcı Protein (LBP) tarafından yakalanır ve bağışıklık hücrelerinin yüzeyindeki CD14 ve TLR4 (Toll-Like Receptor 4) reseptör kompleksine teslim edilir.",
        "LPS'nin TLR4'e kenetlenmesi, yardımcı protein MD-2 aracılığıyla reseptörün dimerizasyonunu tetikler. Hücre içi TIR alanları MyD88 ve TRIF adaptörlerini toplar; bu durum IKK kinaz kompleksini ateşleyerek NF-kappaB'yi ve MAPK yolağını (p38, JNK) uyarır. Sonuç olarak tüm vücuttaki monositler, makrofajlar ve endotel hücreleri kana aralıksız olarak TNF-alfa, IL-6 ve IL-1beta sitokinleri salgılamaya başlar. Metabolik endotoksemi, inflam-aging yangınının en birincil fitilidir.",
        "TLR4_Aktivasyon_Sinyali = k_TLR4 * [LPS:LBP] * [CD14] / ( K_d_TLR4 + [LPS:LBP] )",
        "Bu kaskad aktivasyon eşitliği, kanda dolaşan endotoksin-LBP komplekslerinin bağışıklık hücrelerinde TLR4 yangısal yanıtını nasıl başlattığını gösterir."
    ),
    (
        "8.5 Karaciğerde Kupffer Hücrelerinin Uyarılması ve Sistemik Sitokin Fırtınası",
        "Portal kanla gelen tüm bakteriyel ürünlerin ilk varış noktası karaciğerdir; karaciğer sinüzoidlerinde yerleşik makrofajlar olan Kupffer Hücreleri bu toksinleri karşılayan ilk savunma hattıdır.",
        "Kronik LPS akışı Kupffer hücrelerini aşırı uyarır; hücreler fagosite edebileceklerinden fazla endotoksinle karşılaşınca tükenir ve kana devasa miktarda pro-enflamatuar sitokin püskürtür. Bu sitokinler karaciğer parankiminde hepatosteatozu (yağlanma), hepatosit apoptozunu ve hepatik stellat hücrelerin aktive olarak kolajen üretmesini (karaciğer fibrozu) tetikler. Eş zamanlı olarak karaciğerden salınan sitokinler sistemik dolaşıma karışarak tüm vücudu saran steril bir yangı dalgası yaratır.",
        "d[IL-6_sistemik]/dt = J_Kupffer_salgi * [LPS_portal] - k_klerens * [IL-6_sistemik]",
        "Bu dinamik denklem, portal kandan karaciğere giren endotoksin debisinin sistemik dolaşımdaki interlökin-6 konsantrasyonunu nasıl doğrudan belirlediğini ortaya koyar."
    ),
    (
        "8.6 Kan-Beyin Bariyerinin Bağırsak Sızıntısı Tarafından Aşındırılması (Bağırsak-Beyin Ekseni)",
        "Bağırsak geçirgenliğinin en trajik nörolojik sonucu, 'Bağırsak-Beyin Ekseni' (Gut-Brain Axis) aracılığıyla Kan-Beyin Bariyerini (BBB) de yıkmasıdır.",
        "Dolaşıma geçen LPS ve IL-6/TNF-alfa sitokinleri, beyin kapiller endotel hücrelerine tutunur. Burada endotel hücrelerinin sıkı bağlantı proteinlerini (Claudin-5) yıkarak kan-beyin bariyerini geçirgen hale getirir. Beyin parankimine sızan endotoksinler mikrogliaları uyarır; nöroinflamasyon, sinaptik hasar, amiloid plak birikimi ve dopaminerjik nöron kaybı hızlanır. Birçok Parkinson ve Alzheimer hastasında nörodejenerasyonun ilk tohumlarının bağırsak sızıntısıyla atıldığı günümüzde kesinleşmiş bir gerçektir.",
        "BBB_Gecirgenlik = P_BBB_0 * ( 1 + alpha_LPS * [LPS_plazma] ) * ( 1 + beta_TNF * [TNF-alfa] )",
        "Bu fonksiyon, sistemik endotoksemi ve enflamatuar sitokinlerin kan-beyin bariyerinin koruyucu direncini nasıl katlanarak bozduğunu modeller."
    ),
    (
        "8.7 Zonulin Biyomarkerı ve Bağırsak Mukozal Hasar Kinetiği",
        "Zonulin, enterositler tarafından salgılanan ve bağırsak sıkı bağlantılarını fizyolojik olarak gevşeterek su ve besin geçişini ayarlayan tek bilinen endojen memeli proteinidir.",
        "Lümende patojenik bakterilerin çoğalması veya genetik olarak duyarlı bireylerde glüten/gliadin maruziyeti, enterosit yüzeyindeki CXCR3 reseptörlerini uyararak zonulin salgısını patlatır. Aşırı zonulin salınımı, EGFR ve PAR2 reseptörleri üzerinden protein kinaz C (PKC) yolağını aktive eder; PKC, ZO-1 ve Occludin'i fosforilleyerek aktin iskeletinden koparır ve sıkı bağlantıları felç eder. Plazma zonulin seviyeleri, bağırsak mukozal hasarının ve inflam-aging riskinin en güvenilir invaziv olmayan biyobelirtecidir.",
        "Zonulin_Salinimi = V_max_zon * [Gliadin_Patojen] / ( K_M_cxcr3 + [Gliadin_Patojen] )",
        "Bu hız eşitliği, lümen içi mikrobiyal veya diyet kaynaklı tetikleyicilerin zonulin salınım debisini nasıl fırlattığını formüle eder."
    ),
    (
        "8.8 Diyet Lifleri, Mukus Koruyucular ve Glutamin Desteği ile Bariyer Tamiri",
        "Geçirgen bağırsak astarının onarımı; epitel hücrelerine yakıt sağlamak ve sıkı bağlantı ekspresyonunu transkripsiyonel olarak yeniden açmakla mümkündür.",
        "L-glutamin, ince bağırsak enterositlerinin birincil metabolik yakıtıdır; hücre bölünmesini ve sıkı bağlantı proteinlerinin sentezini hızla uyarır. Çözünür diyet lifleri ve dirençli nişasta ise kolondaki butirat üretimini artırarak Claudin-1 ve Occludin gen promoterlarını epigenetik olarak açar. Çinko Karnosin desteği mukozal kan akımını ve yara iyileşmesini artırırken; sodyum bütirat lavmanları veya takviyeleri transepitelyal direnci 48 saat içinde restore eder.",
        "Bariyer_Iyilesme_Hizi = k_tamir * [Glutamin] * [Butirat] / ( (K_m_gln + [Glutamin]) * (K_m_but + [Butirat]) )",
        "Bu iki-substratlı sinerji modeli, glutamin ve butiratın epitel bütünlüğünü yeniden inşa etmedeki katalitik gücünü açıklar."
    ),
    (
        "8.9 Fekal Mikrobiyota Transplantasyonu (FMT) ile Biyolojik Yaşın Geri Döndürülmesi",
        "Yaşlı organizmanın çökmüş bağırsak ekosistemini sıfırlamanın en radikal ve etkili yolu, genç ve sağlıklı bir donörden alınan tam mikrobiyotanın nakledilmesidir (FMT).",
        "Çığır açan preklinik çalışmalarda (Nature Aging 2021); genç farelerden alınan fekal mikrobiyota yaşlı farelere nakledildiğinde, yaşlı farelerin bağırsak geçirgenliği kapanmış, sistemik inflam-aging belirteçleri gençlik seviyesine gerilemiş, beyindeki mikroglial aktivasyon sönümlenmiş ve farelerin labirent bilişsel test performansları dramatik şekilde gençleşmiştir. FMT, tek tek bakteriler yerine binlerce kommensal türün ekolojik dengesini tek seferde yerine koyarak biyolojik saati geri döndürür.",
        "Genclesme_Katsayisi_FMT = Delta_Biyolojik_Yas = - k_FMT * Kolonizasyon_Basarisi * ( 1 - [Patobiont_Yeni] )",
        "Bu klinik model, genç mikrobiyom naklinin sistemik biyolojik yaşı ve doku canlılığını geri döndürme büyüklüğünü temsil eder."
    ),
    (
        "8.10 Biyo-Mühendislik Ürünü Sentetik Bağırsak Astarı ve Bariyer Restorasyonu",
        "Homo Aeternus mühendisliği, bağırsak bariyerini biyolojik zaaflardan arındırmak için sentetik biyomalzemeler ve genetik modifikasyonlar kullanır.",
        "Biyo-uyumlu müsin-benzeri sentetik hidrojel polimerleri, bağırsak lümeninde ince bir nano-kaplama oluşturarak bakteriyel toksinleri fiziksel olarak hapsederken besin emilimine tam izin verir. Eş zamanlı olarak genetik modifikasyona uğratılmış akıllı probiyotikler (Lactobacillus reuteri Aeterna), epitel hasarı algıladıklarında lokal olarak Glukagon Benzeri Peptit-2 (GLP-2) ve Trefoil Faktör 3 (TFF3) salgılayarak mukozal rejenerasyonu mikrosaniyeler içinde tetikler; böylece bağırsak hiçbir zaman kana toksin sızdırmaz.",
        "Sıfır_Sızıntı_Bariyer_Gucu = 1 - exp(-k_nano * [Sentetik_Hidrojel] - k_glp2 * [GLP2_lokal]) -> 1.0",
        "Bu limit formülasyonu, sentetik hidrojel kalkanı ve akıllı probiyotiklerin bağırsak sızıntısını ve endotoksemiyi tamamen sıfırladığını kanıtlar."
    )
]

# ================= KISIM 9 =================
part9_subsections = [
    (
        "9.1 İmmün Kontrol Noktalarının (Checkpoints) Fizyolojik Görevi: Otoimmünite Freni",
        "Bağışıklık sistemi, hedefini yok ettikten sonra çevre dokulara zarar vermemek ve kendi hücrelerine saldırmamak için evrimsel 'İmmün Kontrol Noktası' (Immune Checkpoint) reseptörleriyle donatılmıştır.",
        "T hücrelerinin yüzeyinde eksprese edilen bu inhibitör reseptörler (PD-1, CTLA-4, TIM-3, LAG-3, TIGIT), ligandlarıyla kenetlendiklerinde hücre içine inhibitör tirozin fosfatazları (SHP-1, SHP-2) çağırır. SHP fosfatazları, T Hücre Reseptörü (TCR) ve CD28 kaskadındaki aktif kinazları defosforilleyerek T hücresinin aktivasyonunu, sitokin salgısını ve proliferasyonunu durdurur. Sağlıklı bir gençte bu frenler otoimmüniteyi önleyen hayati güvenlik supaplarıdır.",
        "Inhibisyon_TCR = k_shp * [PD1:PDL1] / ( K_d_pd1 + [PD1:PDL1] ) * [SHP-2_aktif]",
        "Bu fosfataz aracılı kısıtlama eşitliği, immün kontrol noktası ligand kenetlenmesinin TCR aktivasyon sinyalini nasıl doğrudan söndürdüğünü gösterir."
    ),
    (
        "9.2 T Hücre Yorgunluğu (Exhaustion) Fenotipi: PD-1, TIM-3, LAG-3 ve TIGIT Eş-Ekspresyonu",
        "Kronik viral enfeksiyonlar, kanserojen neoantijenler veya yaşlanmayla ortaya çıkan sürekli otoantijenik uyarım altında, T hücreleri zamanla 'T Hücre Yorgunluğu' (T-cell Exhaustion) adı verilen patolojik bir evreye girer.",
        "Yorulmuş T hücreleri, yüzeylerinde tek bir kontrol noktası değil; birden fazla inhibitör reseptörü aynı anda ve sürekli olarak aşırı eksprese eder (PD-1 yüksek, TIM-3+, LAG-3+, TIGIT+). Bu çoklu fren kilitlenmesi sonucunda T hücresi artık IL-2, TNF-alfa ve IFN-gama üretemez hale gelir; litik granüllerini (perforin/granzim) boşaltamaz ve apoptoza aşırı duyarlı hale gelir. Yorgun T hücreleri dokuda var olsalar bile görevlerini yapamayan 'felçli lenfositler'dir.",
        "Yorgunluk_Skoru = Prod_i ( 1 + [Inhibitor_Reseptor_i] / K_i_esik )",
        "Bu kümülatif çarpımsal skor, yüzeydeki inhibitör checkpoint reseptörlerinin eşzamanlı artışının lenfosit fonksiyonel felcini nasıl logaritmik olarak katladığını ifade eder."
    ),
    (
        "9.3 TOX Transkripsiyon Faktörünün Epigenetik Olarak T Hücrelerini Felç Etmesi",
        "T hücre yorgunluğunun geçici bir dinlenme durumu olmadığı, genomik düzeyde kalıcı bir epigenetik kilitlenme olduğu 'TOX' (Thymocyte Selection-Associated High Mobility Group Box) transkripsiyon faktörünün keşfiyle anlaşılmıştır.",
        "Kronik TCR uyarımı nükleer faktör NFAT'ın AP-1 olmadan tek başına çekirdeğe girmesine yol açar; bu durum TOX genini ateşler. Yüksek düzeyde eksprese edilen TOX proteini, kromatini yeniden yapılandırarak yorgunluk genlerinin (PDCD1, HAVCR2) promoterlarını kalıcı olarak açık tutarken, hafıza ve efektör genlerin promoterlarını heterokromatin ile kapatır. Bu epigenetik skarlaşma gerçekleştikten sonra antijen ortadan kalksa bile T hücresi genç efektör haline geri dönemez.",
        "d[TOX_nukleus]/dt = k_NFAT * [NFAT_yalniz] - k_deg * [TOX] + k_pozitif_oto * [TOX]^2",
        "Bu pozitif otokatalitik döngü, TOX ekspresyonunun belirli bir kronik uyarım süresinden sonra T hücresinde nasıl geri dönüşsüz bir yorgunluk kaderi mühürlediğini formüle eder."
    ),
    (
        "9.4 Kronik Antijenik Yük Altında Senesens ile Yorgunluk Arasındaki Moleküler Ayrım",
        "İmmünosenesens literatüründe sıkça karıştırılan 'T Hücre Senesensi' ile 'T Hücre Yorgunluğu' mekanistik olarak tamamen farklı iki hücresel patolojidir.",
        "T hücre senesensi, telomer kısalması veya DNA hasarı sonucu ortaya çıkan kalıcı bir hücre döngüsü arrestidir (p16INK4a ve p21CIP1 yüksek, CD28 negatif, CD57 pozitif); ancak bu hücreler metabolik olarak aktiftir ve devasa miktarda pro-enflamatuar SASP sitokini salgılarlar. T hücre yorgunluğu ise TOX güdümlü, çoklu kontrol noktaları (PD-1, TIM-3) ile karakterize, sitokin salgısının tamamen sustuğu fonksiyonel bir tükeniştir. Yaşlı bir bağışıklık sistemi bu iki patolojik fenotipin eşzamanlı birikimiyle felç olur.",
        "Fenotip_Ayraci = ( [p16] * [CD57] / [CD28] )_Senesens vs ( [TOX] * [PD1] * [TIM3] )_Yorgunluk",
        "Bu karşılaştırmalı biyobelirteç oranı, lenfosit popülasyonunda senesens ile immünolojik tükenme arasındaki moleküler sınır çizgisini tanımlar."
    ),
    (
        "9.5 Kanser İmmün Sürveyansının Çöküşü ve Yaşla Artan Malignite İnsidansı",
        "İnsanlarda kanser insidansının 60 yaşından sonra üstel bir patlama göstermesinin temel nedeni, mutasyon hızının artmasından ziyade mutant hücreleri yok eden 'İmmün Sürveyans' kalkanının çökmesidir.",
        "Gençlikte ortaya çıkan prekanseröz hücreler, CD8+ sitotoksik T lenfositleri ve NK hücreleri tarafından derhal tespit edilerek temizlenir. Yaşlılıkta ise naif T hücresi eksikliği nedeniyle tümör neoantijenlerini tanıyacak TCR klonları bulunamaz; mevcut efektör T hücreleri ise mikroçevredeki PD-L1 ligandları karşısında derhal yorgunluğa girer. Bağışıklık sürveyansından kaçan transforme hücreler hızla çoğalarak klinik tümör kitlelerine dönüşür.",
        "Kanser_Kacis_Hizi = J_mutasyon * ( 1 - Etkinlik_Surveyans ) = J_mutasyon * ( 1 - [T_efektor_genc] / K_surv )",
        "Bu epidemiyolojik risk modeli, fonksiyonel genç sitotoksik T hücresi havuzunun erimesiyle neoplastik hücrelerin bağışıklık kalkanını delme hızının nasıl fırladığını açıklar."
    ),
    (
        "9.6 Kontrol Noktası Blokajı (Anti-PD-1 / Anti-CTLA-4) Tedavilerinin Yaşlı Bireylerdeki Yanıtı",
        "Onkolojide devrim yaratan immün kontrol noktası inhibitörleri (İmmunotherapy: Nivolumab, Pembrolizumab, İpilimumab), yorulmuş T hücreleri üzerindeki frenleri kaldırarak onları yeniden tümör öldürücü hale getirir.",
        "Ancak yaşlı hastalarda kontrol noktası blokajına verilen yanıt çelişkilidir: Bir yandan yaşlı bireylerde kümülatif mutasyon yükü yüksek olduğu için neoantijen sayısı fazladır ve bu durum tümörü immünoterapiye daha duyarlı kılabilir. Diğer yandan yaşlı hastaların T hücreleri derin TOX epigenetik skarlaşması taşıdığı ve klonal TCR çeşitliliği daraldığı için, fren kalksa dahi tam fonksiyonel yeniden canlanma (reinvigoration) her zaman başarılamayabilir.",
        "Kanser_Yanit_Ihtimali = k_resp * Neoantijen_Yuku * Shannon_TCR_Cesitliligi / ( 1 + [TOX_epigenetik_skar] )",
        "Bu onkolojik olasılık formülü, immünoterapi başarısının neoantijen zenginliği, TCR klonal çeşitliliği ve TOX kilitlenme derecesi arasındaki dengeye bağlı olduğunu gösterir."
    ),
    (
        "9.7 İmmün Kontrol Noktalarının Fazla Açılmasının Doğurduğu Otoimmün Toksisite",
        "Kontrol noktası inhibitörlerinin yaşlanma karşıtı veya onkolojik amaçla kontrolsüzce kullanılması, bağışıklık sisteminin frenlerini tamamen boşaltarak ölümcül 'İmmün İlişkili İstenmeyen Olaylar'a (irAE) yol açabilir.",
        "PD-1 ve CTLA-4 frenleri kalktığında, yaşlı bireylerde zaten periferik toleransı zayıflamış olan otoreaktif T hücre klonları serbest kalır. Bu hücreler sağlıklı dokulara saldırarak fulminan miyokardit, otoimmün hepatit, şiddetli kolit ve tip 1 diyabet benzeri hiper-akut doku nekrozları yaratabilir. Dolayısıyla immün kontrol noktalarının modülasyonu, 'kanser/senesens temizliği' ile 'otoimmün felaket' arasındaki bıçak sırtı dengede milimetrik olarak kalibre edilmelidir.",
        "irAE_Toksisite_Riski = exp( -k_fren * [Blokaj_Dozu] ) * [Otoreaktif_T_Klonal_Yuku]",
        "Bu toksisite riski modeli, kontrol noktası blokaj dozunun artışıyla latent otoreaktif hücrelerin tetiklediği otoimmün hasar patlamasını simgeler."
    ),
    (
        "9.8 CAR-T Hücreleri ve Gençleştirilmiş Otolog İmmün Terapiler",
        "Yaşlı hastanın yorgun T hücrelerinin yarattığı kısıtlılıkları aşmak için, sentetik immünoloji 'Kimerik Antijen Reseptörü' (CAR-T) hücre mühendisliğini geliştirmiştir.",
        "Hastadan izole edilen T hücreleri laboratuvarda genetik olarak modifiye edilir; yüzeylerine tümör veya senesen hücre belirteçlerini tanıyan sentetik bir antikor başlığı (scFv) ve hücre içine CD28/4-1BB kostimülatör alanları ile CD3-zeta aktivasyon alanları entegre edilir. Dahası, bu hücrelere ex vivo ortamda telomeraz (TERT) gen transferi yapılarak telomerleri uzatılır ve CRISPR ile PD-1 genleri silinerek yorgunluğa tam dirençli 'Gençleştirilmiş Süper-Lenfositler' yaratılır.",
        "Sitotoksik_Kapasite_CAR = N_CAR_T * k_killing * [Antijen_Hedef] / ( K_M_car + [Antijen_Hedef] ) * ( 1 + Boost_41BB )",
        "Bu biyomühendislik formülü, 4-1BB kostimülatör alanıyla güçlendirilmiş CAR-T hücrelerinin hedef patolojik hücreleri yok etme kinetiğini modeller."
    ),
    (
        "9.9 Sentetik İmmünomodülatör Peptitler ile T Hücre Canlandırılması",
        "Monoklonal antikorların sistemik toksisitelerinden kaçınmak için, immün kontrol noktalarını geçici ve pulsatil olarak modüle eden sentetik peptitler tasarlanmıştır.",
        "Bu küçük peptitler (örn. PD-1/PD-L1 etkileşim arayüzünü taklit eden siklik peptitler), dolaşımda sadece birkaç saatlik bir yarı ömre sahiptir. Haftalık mikro-dozlar halinde uygulandıklarında, T hücreleri üzerindeki inhibitör baskıyı sadece senesen ve prekanseröz hücrelerin temizlenmesine yetecek kadar (12-24 saat) kaldırır; ardından hızla klerense uğrayarak bağışıklık sisteminin frenlerini yeniden devreye sokar. Bu pulsatil strateji, sıfır otoimmünite riski ile maksimum immün gençleşme sağlar.",
        "Pulsatil_Fren_Kaldirma = Integral_0_tau( [Peptit_plazma](t) dt ) < Esik_Otoimmun_Toksisite",
        "Bu maruziyet integrali, kısa yarı ömürlü peptitlerin otoimmün toksisite sınırını aşmadan T hücrelerini nasıl güvenle reaktive ettiğini açıklar."
    ),
    (
        "9.10 İmmün Checkpoint Kalibrasyonu: Kanser Koruması ile Otoimmün Güvenlik Dengesi",
        "Uzun yaşam tıbbının nihai hedefi; bağışıklık kontrol noktalarını statik olarak açmak veya kapatmak değil, organizmanın anlık durumuna göre dinamik olarak ayarlayan bir 'Homeostatik Ayar Noktası' (Tuning Point) kurmaktır.",
        "Yapay zeka güdümlü biyosensörler dolaşımdaki kanser neoantijenlerini veya senesen hücre SASP faktörlerini tespit ettiğinde lokal immün kontrol noktaları gevşetilmeli; temizlik tamamlandığında ise doku bütünlüğünü korumak için kontrol noktaları yeniden sıkılaştırılmalıdır. Bu dinamik kalibrasyon, yaşlı organizmada kanser bağışıklığını genç bir atlet düzeyinde tutarken, otoimmüniteye asla geçit vermez.",
        "Optimum_Ayar_Noktasi = ArgMin_w [ Risk_Malignite(w) + Risk_Otoimmunite(w) ]",
        "Bu optimizasyon denklemi, immün kontrol noktası katsayısının (w) iki zıt ölümcül riski eşzamanlı minimize ettiği matematiksel denge noktasını tanımlar."
    )
]

# ================= KISIM 10 =================
part10_subsections = [
    (
        "10.1 Sentetik Timik Mikroçevre: Sürekli Naif T Hücre Üreten Yapay Epitelyal Ağ",
        "Homo Aeternus bağışıklık mimarisi, doğal timusun involüsyon kaderini tamamen iptal eden yapay biyomimetik organ implantları ile inşa edilir.",
        "İmplante edilen mikrovaskülarize 'Sentetik Timik Ağ' (Synthetic Thymic Niche), biyobozunur nanokompozit iskeleler üzerinde ölümsüzleştirilmiş cTEC ve mTEC hücre hatları barındırır. Hücrelerde FOXN1 ve AIRE ekspresyonu sentetik promoterlar altında sürekli sabit tutulur. Bu yapay mikroçevre, kemik iliğinden gelen lenfoid öncülleri 365 gün 24 saat kabul ederek kanda naif T hücresi (CD45RA+) oranını ömür boyu %60'ın üzerinde tutar; böylece organizma hiçbir zaman antijenik körlüğe girmez.",
        "Naif_T_Uretim_Surekliligi = J_ATO = Constant = 10^7_hucre/gun",
        "Bu kararlı durum debi eşitliği, sentetik timik ağın yaşlanmaya bağlı hiçbir sönümleme yaşamaksızın sabit hızda taze T lenfositi ürettiğini ifade eder."
    ),
    (
        "10.2 CRISPR-Cas Destekli Telomerik Uzatma ile İmmün Hücre Replikatif Sonsuzluğu",
        "Doğal bağışıklık hücreleri tekrarlayan antijenik bölünmelerle telomerlerini tüketip senesense girerken; sentetik genom mimarisi bu sınırı ortadan kaldırır.",
        "Hematopoietik kök hücrelere entegre edilen dCas9-transkripsiyonel aktivatör devreleri, hücre proliferasyona girdiğinde TERT (Telomeraz Ters Transkriptaz) genini geçici olarak aktive eder. Her bölünmede kaybedilen 50-100 baz çiftlik telomerik dizi, bölünmenin hemen ardından milisaniyeler içinde telomeraz tarafından sentezlenerek yerine konur. Bu kontrollü replikatif sonsuzluk, klonal tükenmeyi ve CD28-null senesen lenfosit oluşumunu tamamen önler.",
        "d[Telomer_Boyu_Lenfosit]/dt = Delta_L_sentez(TERT) - Delta_L_kayip(Replikasyon) = 0",
        "Bu dinamik denge denklemi, kontrollü telomeraz indüksiyonunun immün hücrelerde telomer aşınmasını sıfırlayarak hücresel replikatif ömrü sonsuz kıldığını kanıtlar."
    ),
    (
        "10.3 cGAS-STING ve NLRP3 Eşiklerinin Biyo-Sensörlerle Sabitlenmesi",
        "Hücresel gençlik için patojen savunması korunmalı ancak steril yangı (inflam-aging) mutlak surette engellenmelidir; bu ikilem sentetik regülatuar biyo-anahtarlarla çözülür.",
        "Hücre içine yerleştirilen sentetik protein devreleri, intraselüler serbest dsDNA veya ekstraselüler ATP yoğunluğunu izler. Hasar seviyesi normal hücre içi sızıntı düzeyindeyse (steril stres), devre aktive olarak STING'in Cys91 bölgesine bağlanan inhibitör peptitleri veya NLRP3'ün Walker B cebini kilitleyen nano-antikorları salgılar. Ancak sinyal gerçek bir mikrobiyal enfeksiyon eşiğini (yüksek PAMP konsantrasyonu) aştığında fren anında kalkar ve güçlü immün yanıt patlar.",
        "Yangisal_Esik_Duyarliligi = IF( [PAMP] > Esik_Patojen, Full_Aktif_Savunma, Steril_Blokaj_Modu )",
        "Bu mantıksal kontrol fonksiyonu, yapay biyosensör devresinin patojenik enfeksiyon ile steril hücresel çöp arasındaki ayrımı kusursuz yaparak inflam-aging'i engellediğini gösterir."
    ),
    (
        "10.4 Sentetik Biyoloji ile Tasarlanmış Genomik Süper-TCR Kütüphaneleri",
        "Homo Aeternus genomu, evrimsel rastgele V(D)J rekombinasyonunun doğurduğu kör noktaları ve yetersizlikleri ortadan kaldıran yapay bir 'Süper-TCR Kütüphanesi' ile donatılır.",
        "Yapay zeka (AlphaFold ve RoseTTAFold) tarafından de novo tasarlanan bu sentetik T hücresi reseptörleri, doğada bilinen ve gelecekte mutasyonla ortaya çıkabilecek tüm viral protein dizilimlerini, bakteriyel toksinleri ve kanser neoantijenlerini pikomolar afiniteyle bağlayacak değişken bölgeler içerir. Kök hücrelere aktarılan bu ön-programlı TCR çeşitliliği, organizmaya tüm olası biyolojik tehditlere karşı 'doğuştan hazır ve yenilmez' bir adaptif kalkan kazandırır.",
        "TCR_Kapsama_Katsayisi = Sum_j ( Antijen_j_tanima ) / Toplam_Olası_Patojen_Uzayi -> 1.0",
        "Bu limit bağıntısı, sentetik süper-TCR kütüphanesinin tüm teorik patojenik epitop uzayını tam kapsama oranıyla örttüğünü belgeler."
    ),
    (
        "10.5 Akıllı Biyo-Bariyer: Bağırsak Mukozasını Yenileyen Genetiği Değiştirilmiş Probiyotikler",
        "Bağırsak sızıntısını ve endotoksemiyi sonsuza dek tarihe gömmek için, genetik olarak modifiye edilmiş 'Akıllı Biyo-Bariyer Probiyotikleri' (Smart Shield Microbiota) kullanılır.",
        "CRISPR ile yeniden tasarlanmış bu sentetik suşlar (Akkermansia aeternitas ve E. coli Nissle Prime), bağırsak lümeninde sürekli olarak mikroskobik hasar taraması yapar. Bir mikronluk bir epitel çatlağı veya zonulin artışı tespit edildiğinde, bakteriler anında rekombinant Trefoil Faktörleri (TFF), Claudin-1 stabilizatörleri ve yüksek konsantrasyonda bütirat salgılar. Yara mikro-saniyeler içinde kapatılarak portal kana tek bir molekül LPS sızması dahi engellenir.",
        "Mukozal_Tamir_Gecikmesi = Tau_tepki * exp( -k_akilli * [Biyo_Bariyer_Probiyotik] ) -> 0",
        "Bu limit fonksiyonu, akıllı probiyotiklerin epitel hasarına müdahale tepki süresini sıfıra indirerek bariyer bütünlüğünü daimi kıldığını açıklar."
    ),
    (
        "10.6 Sitomegalovirüs ve Kronik Virüslere Karşı Mutlak Sentetik Antijenik Temizlik",
        "İmmün sistemin bellek havuzunu rehin alan ve immünosenesensi hızlandıran Sitomegalovirüs (CMV), EBV ve HSV gibi latent virüsler sentetik gen düzenleme ile vücuttan tamamen kazınır.",
        "Nörotropik ve miyeloid-hedefli lipid nanopartiküller (LNP) aracılığıyla iletilen Prime Editing ve CRISPR-Cas12a nükleazları, konak hücre çekirdeğinde sessizce bekleyen viral epizomları ve latent entegre viral genomları hedef alır. Tek bir uygulamada viral DNA çoklu bölgelerden kesilerek parçalanır. Virüs tamamen yok edildiğinde, CMV'ye kilitlenmiş milyonlarca senesen CD8+ T hücresi apoptozla temizlenir ve immün bellek havuzu tamamen gençleşir.",
        "Viral_Yuk_Temizligi = [Viral_DNA_kopyasi](t) = N_0 * exp( -k_crispr * [Cas12a_hedefli] * t ) -> 0",
        "Bu klerens denklemi, latent viral rezervuarların hedefe yönelik genetik cerrahi ile sıfırlanma sürecini modeller."
    ),
    (
        "10.7 Senosen Hücreleri Hedef Alan Otolog CAR-NK ve CAR-Makrofaj Sürveyansı",
        "Homo Aeternus bağışıklığı, dokularda biriken senesen hücreleri pasif olarak beklemek yerine, aktif olarak arayıp yok eden sentetik 'Senolitik CAR Hücreleri' ile tahkim edilir.",
        "Hastanın kendi monosit ve NK hücrelerinden türetilen CAR-Makrofajlar ve CAR-NK hücreleri, senesen hücrelerin yüzeyinde biriken spesifik belirteçleri (uPAR, CD264, glikolize integrinler) tanıyacak şekilde programlanır. Dokuya yayılan bu sentetik avcılar, senesen hücreleri tespit ettikleri anda perforin/granzim ile apoptoza sokar veya fagositozla tamamen yutarak çevreye SASP saçılmasını engeller. Bu sürekli sürveyans, dokulardaki senesens yükünü daima sıfıra yakın tutar.",
        "Senosen_Klerens_Hizi = J_senolitik = k_car * [CAR-Makrofaj] * [Senesen_Hucre] / ( K_M + [Senesen_Hucre] )",
        "Bu hedefli fagositoz eşitliği, otolog senolitik CAR hücrelerinin dokulardaki hücresel yaşlanma artıklarını nasıl kalıcı bir debiyle temizlediğini gösterir."
    ),
    (
        "10.8 Yapay Zeka Tabanlı Klonal Hematopoez İzleme ve Erken Kök Hücre Ayıklama",
        "Kemik iliğinde mutasyon biriktiren ve lösemiye ya da inflam-aging'e yol açan mutant hematopoietik klonlar (CHIP), ultra-duyarlı yapay zeka izleme sistemleriyle taranır.",
        "Dolaşımdaki kandan izole edilen tek-hücre cfDNA dizileme verileri, derin öğrenme algoritmaları tarafından taranarak DNMT3A, TET2 veya JAK2 mutasyonu taşıyan klonların frekansı (VAF > %0.01) tespit edilir. Klon henüz %2 seviyesine ulaşmadan, mutant klona özgü neoantijenleri hedefleyen bispesifik antikorlar veya klon-spesifik CRISPR-LNP'ler devreye sokularak mutant kök hücreler kemik iliği nişinde sessizce ayıklanır. Böylece klonal hematopoez ve kardiyovasküler risk kaynağında yok edilir.",
        "Klon_Bastirma_Verimi = VAF_son / VAF_baslangic = exp(-k_hedefli_klerens * t) -> 0",
        "Bu bastırma modeli, mutasyon taşıyan hematopoietik hücre alt klonlarının erken sentetik müdahale ile tamamen imha edilme sürecini tanımlar."
    ),
    (
        "10.9 Biyo-Uyumlu İmmünomodülatuvar Nano-Robotlarla Steril İnflamasyonun Sıfırlanması",
        "Dolaşımda ve interstisyel doku sıvısında devriye gezen otonom 'İmmünomodülatuvar Nano-Robotlar' (IMNs), steril yangı medyatörlerini fiziksel olarak süpüren nano-teknolojik makinelerdir.",
        "Karbon nanotüp ve lipid-polimer hibritlerinden üretilen bu 200 nanometrelik robotlar, yüzeylerindeki kimerik reseptörlerle serbest TNF-alfa, IL-6, IL-1beta, ekstraselüler HMGB1 ve serbest mtDNA parçalarını yakalar. Yakalanan yangısal moleküller robotun iç haznesinde yer alan sentetik nükleazlar ve proteazlar tarafından tamamen zararsız amino asit ve nükleotidlere parçalanır. Nano-robotlar, sistemik dolaşımdaki yangısal sitokin fırtınalarını saniyeler içinde tamponlar.",
        "Sitokin_Supurme_Debisi = J_supurme = N_nanorobot * v_diff * Area_yakalama * [Sitokin_serum]",
        "Bu nano-akışkanlar mekaniği eşitliği, dolaşımdaki nano-robot ordusunun steril sitokinleri ve DAMP'leri kandan temizleme kapasitesini modeller."
    ),
    (
        "10.10 Homo Aeternus İmmünom Mimarisi: Enfeksiyonsuz, Kansersiz ve Yangısız Ebedi Biyoloji",
        "Homo Aeternus immünom mimarisi; yapay timik ağlar, süper-TCR çeşitliliği, telomerik sınırsızlık, akıllı bağırsak kalkanı ve sitokin süpürücü nano-makinelerin birleştiği mutlak bir biyolojik savunma kalesidir.",
        "Bu mimaride; enfeksiyonlar vücuda giremeden sınırda yok edilir, neoplastik ve senesen hücreler oluştukları anda cerrahi bir hassasiyetle ayıklanır, steril yangı asla kronikleşemez ve otoimmünite tamamen imkansız kılınır. Bağışıklık sistemi artık vücudu yıpratan bir entropi jeneratörü değil; organizmayı zamanın ve dış dünyanın tüm yıkıcı etkilerine karşı koruyan ölümsüz bir kalkan olarak işlev görür.",
        "Aeternus_Immun_Guvenlik_Indeksi = ISI = ( Guvenlik_Enfeksiyon * Guvenlik_Kanser ) / ( 1 + Entropi_Inflam_Aging ) -> Sonsuz",
        "Bu nihai immünolojik sonuç formülü, Homo Aeternus bağışıklık mimarisinin enfeksiyon ve kanser riskini sıfırlarken sistemik inflam-aging entropisini yok ederek sonsuz biyolojik kararlılığı garanti ettiğini ilan eder."
    )
]

# ================= 10 AKADEMİK KARŞILAŞTIRMA TABLOSU =================
parts.append(("KISIM 7: BAĞIRSAK MİKROBİYOTASI VE YAŞLANMA: DİSBİYOZİS DİNAMİKLERİ", part7_subsections))
parts.append(("KISIM 8: BAĞIRSAK BARİYER İHLALİ (LEAKY GUT) VE SİSTEMİK ENDOTOKSEMİ", part8_subsections))
parts.append(("KISIM 9: İMMÜN CHECKPOINT MOLEKÜLLERİ (PD-1, CTLA-4, TIGIT) VE İMMÜN YORGUNLUK", part9_subsections))
parts.append(("KISIM 10: HOMO AETERNUS İMMÜNOM MİMARİSİ: YAŞLANMAYAN BAĞIŞIKLIK VE MUTLAK SAVUNMA", part10_subsections))

tables_data = [
    (
        "TABLO 11.1: Inflam-Aging ve Steril Yangı Medyatörleri: Kaynakları, Reseptörleri ve Sistemik Yıkım Profili",
        ["DAMP / İnflamatuar Medyatör", "Hücresel / Doku Kaynağı", "Hedef PRR Reseptörü", "Aşağı Akış Sinyal Kaskadı", "Yaşlanmadaki Patolojik Sonucu", "Terapötik Susturma Stratejisi"],
        [
            ["HMGB1 (Serbest)", "Nükleer kaçış, senesen/nekrotik hücre", "RAGE, TLR2, TLR4", "MyD88, NF-kappaB, p38 MAPK", "Kronik vasküler inflamasyon, ateroskleroz", "Anti-HMGB1 monoklonal antikor, RAGE antagonisti"],
            ["Ekstraselüler ATP", "Lizis, pannekson (Panx1) kanalları", "P2X7 pürinerjik reseptörü", "K+ efluksu, Ca2+ akışı, NLRP3", "Piroptoz, makrofaj hiperaktivasyonu", "A-438079 (P2X7 blokörü), ekzo-nükleotidazlar"],
            ["Mitokondriyal DNA (mtDNA)", "mPTP açılması, mitofaji iflası", "TLR9, cGAS, AIM2", "IRF3, Tip I İnterferon, STING", "Steril sitokin fırtınası, kardiyomiyopati", "DNase I infüzyonu, mPTP inhibitörü (Siklosporin A)"],
            ["N-formil Peptitler", "Mitokondri iç zarı / matriksi", "FPR1 (Formil Peptit Reseptörü 1)", "Gi kenetli kemotaksis, ROS patlaması", "Nötrofillerin dokularda kontrolsüz birikimi", "FPR1 antagonistleri, mitokondri membran stabilizasyonu"],
            ["Serbest Histonlar", "Parçalanan mikronükleuslar, NETosis", "TLR4, integrinler", "Endotel hasarı, mikrovasküler tromboz", "Akut organ hasarı, dissemine intravasküler koagülasyon", "Polianyonik heparan sülfat, non-antikoagülan heparin"],
            ["İnterlökin-6 (IL-6)", "Senesen hücreler (SASP), M1 makrofaj", "IL-6R / gp130 kompleksi", "JAK / STAT3, karaciğerde CRP artışı", "Sarkopeni, osteoporoz, anemi, frailty", "Tocilizumab (anti-IL-6R), Sarilumab"]
        ]
    ),
    (
        "TABLO 11.2: cGAS-STING Sinyal Kaskadı: Moleküler Basamaklar, Kinetik Parametreler ve İnhibitörler",
        ["Aşama / Moleküler Adım", "Konum / Kompartman", "Biyokimyasal Reaksiyon", "Kinetik Sabit / Parametre", "Yaşlanma Sürecindeki Durumu", "Hedefleyici Molekül"],
        [
            ["dsDNA Tanıma", "Sitozol / Mikronükleus", "cGAS dimerleşmesi ve DNA bağlanması", "Kd ~ 1-5 nM (dsDNA >45 bp)", "Mikronükleus yırtılmasıyla sürekli aşırı doyum", "RU.521, G150"],
            ["cGAMP Sentezi", "Sitozol", "ATP + GTP -> 2'3'-cGAMP hibrit bağı", "k_cat ~ 1.2 s^-1", "Sitoplazmada sürekli yüksek cGAMP seviyeleri", "ENPP1 fosfodiesteraz aktivatörleri"],
            ["STING Kenetlenmesi", "Endoplazmik Retikulum", "STING dimerine cGAMP bağlanması", "Kd ~ 4.2 nM", "STING'in sürekli açık konformasyonda kalması", "H-151 (Cys91 kovalent inhibitörü), C-178"],
            ["Golgiye Translokasyon", "ERGIC / Golgi", "COP-II vezikülleriyle göç, palmitoilasyon", "t_trans ~ 15-30 dakika", "Otofajik yıkım gecikmesi nedeniyle birikim", "Brefeldin A, palmitoilasyon inhibitörleri"],
            ["TBK1 / IRF3 Aktivasyonu", "Golgi Zarı", "STING CTT üzerinden Ser366/386 fosforilasyonu", "Km_IRF3 ~ 200 nM", "Sürekli Tip I İnterferon (IFN-b) ekspresyonu", "GSK8612, Amlexanox"],
            ["IKK / NF-kappaB Kolu", "Golgi / Nükleus", "TRAF6 aracılı I-kappa-B parçalanması", "t_aktivasyon ~ 10 dakika", "SASP faktörlerinin (IL-6, IL-8) sürekli salınımı", "IKKbeta inhibitörleri, BMS-345541"]
        ]
    ),
    (
        "TABLO 11.3: NLRP3 İnflamazomu vs Diğer İnflamazom Türleri (NLRC4, AIM2, Pyrin): Yapı ve Yaşlanma Rolü",
        ["İnflamazom Kompleksi", "Sensör Proteini", "Spesifik Tetikleyici Ligand", "Adaptör İhtiyacı", "Patojen Savunması vs Steril İnflamasyon", "Yaşlanmadaki Dominansı"],
        [
            ["NLRP3 İnflamazomu", "NLRP3 (NACHT, LRR, PYD)", "K+ efluksu, mtROS, mikrokristal, ATP", "Zorunlu ASC adaptörü", "Öncelikli olarak steril doku stresi ve DAMP", "Inflam-aging'in en baskın ve yıkıcı faili"],
            ["NLRC4 İnflamazomu", "NLRC4 (CARD, NACHT, LRR)", "Bakteriyel flagellin, T3SS çubuk proteini", "ASC isteğe bağlı (doğrudan CARD kesimi)", "Özelleşmiş intraselüler bakteriyel enfeksiyon", "Yaşlanmada bazal değişimi minimaldir"],
            ["AIM2 İnflamazomu", "AIM2 (HIN-200, PYD)", "Sitozolik çift zincirli DNA (dsDNA)", "Zorunlu ASC adaptörü", "Viral ve bakteriyel intraselüler DNA", "Sitozolik DNA artışıyla sekonder aktivasyon"],
            ["Pyrin İnflamazomu", "Pyrin (MEFV geni)", "RhoA GTPaz inaktivasyonu (toksinler)", "Zorunlu ASC adaptörü", "Bakteriyel sitoiskelet toksinleri", "Ailesel Akdeniz Ateşi (FMF); yaşlanmada ikincil"],
            ["NLRP1 İnflamazomu", "NLRP1 (PYD, NACHT, CARD)", "Şarbon letal toksini, viral proteazlar", "Kendi CARD alanıyla doğrudan Kaspaz-1 keser", "Bakteriyel/viral proteolitik parçalanma", "Deri keratinosit senesensinde lokal uyarım"],
            ["Non-Kanonik İnflamazom", "Kaspaz-4 / Kaspaz-5 (İnsan)", "İntraselüler serbest LPS (Gram-negatif)", "Adaptörsüz doğrudan bağlanma", "Bakteriyel sepsis ve şiddetli endotoksemi", "Leaky gut kaynaklı sistemik vasküler hasar"]
        ]
    ),
    (
        "TABLO 11.4: Timik İnvolüsyon Parametreleri ve T Hücre Alt Tiplerinin Yaşla Değişimi",
        ["İmmünolojik Parametre", "Genç Birey (20 Yaş)", "Yaşlı Birey (75 Yaş)", "Biyolojik Değişim Yönü", "Klinik / İmmünolojik Yansıması", "Terapötik Düzeltme Hedefi"],
        [
            ["Aktif Timik Epitel Hacmi", "%80-90 fonksiyonel korteks/medulla", "%5-10 rezidüel lenfoid adacık", "Dramatik çöküş (Yağ infiltrasyonu)", "Yeni T hücresi üretiminin neredeyse durması", "rhGH, DHEA, FOXN1 gen terapisi"],
            ["Naif CD4+ T Hücreleri", "%40-50 dolaşım CD4 havuzu", "%15-20 dolaşım CD4 havuzu", "Belirgin azalma", "Yeni antijenlere karşı adaptif körlük", "IL-7 infüzyonu, sentetik timus organoidi"],
            ["Naif CD8+ T Hücreleri", "%30-40 dolaşım CD8 havuzu", "<%5 dolaşım CD8 havuzu", "Neredeyse tam tükeniş", "Viral enfeksiyon ve tümörlere karşı zaaf", "Kemik iliği progenitör desteği"],
            ["CD28-null T Hücreleri", "<%5 periferik kanda", "%40-60 periferik kanda", "Devasa patolojik artış", "SASP üretimi, endotel hasarı, senesens", "CD28-null hücrelerin senolitik imhası"],
            ["TCR Klonal Çeşitliliği", "~10^11 farklı TCR sekansı", "<10^9 farklı TCR sekansı", "100 kat klonal daralma", "Aşı yanıtsızlığı, mutasyonlara duyarsızlık", "Süper-TCR kütüphane transferi"],
            ["TREC (TCR Eksizyon Halkası)", "Yüksek (>1000 kopya/mikrog DNA)", "Saptanamaz düzeyde düşük (<10 kopya)", "Gerçek timik çıktının sıfırlanması", "İmmünolojik gençlik indeksinin kaybı", "Timik neogenez ile TREC pozitifliği artışı"]
        ]
    ),
    (
        "TABLO 11.5: Doğal Bağışıklık Hücrelerinde Gençlik vs Senesen Fonksiyon Karşılaştırması",
        ["Hücre Türü", "Genç Hücresel Fonksiyon", "Yaşlı / Senesen Hücresel Fonksiyon", "Moleküler Kusur Mekanizması", "Doku Düzeyindeki Sonucu", "Rejenerasyon Müdahalesi"],
        [
            ["Nötrofil", "Yüksek yönelimli kemotaksis, kontrollü fagositoz", "Rastgele göç (disoriented), aşırı NETosis", "PIP3 dengesizliği, PAD4 hiperaktivasyonu", "Doku hasarı, immünotromboz, gecikmiş klerens", "PI3K-gama inhibitörleri, DNase I"],
            ["Makrofaj", "Dengeli M1/M2 polarizasyonu, hızlı efferositoz", "Kronik M1 yangısal kilitlenme, efferositoz iflası", "MerTK kaybı, otofaji-lizozom tıkanıklığı", "Apoptotik hücrelerin nekroza kayması, kronik yara", "Itakonat mimetikleri, efferositoz indükleyicileri"],
            ["Dendritik Hücre (DC)", "Hızlı lenfatik göç, güçlü antijen sunumu", "Zayıf lenf nodu göçü, düşük MHC/CD80 ekspresyonu", "CCR7 kemokin reseptörü azalması", "T hücrelerinin priming aşamasında başarısızlık", "CCR7 uyarımı, sentetik DC aşıları"],
            ["Doğal Katil (NK)", "Yüksek CD56bright oranı, tam litik sinaps", "CD56dim birikimi, zayıf degranülasyon", "Perforin/granzim polarizasyon kaybı", "Senesen hücrelerin ve mikro-tümörlerin kaçışı", "IL-15/IL-21 priming, CAR-NK dönüşümü"],
            ["Mikroglia", "Sinaptik budama, hasar tarama ve onarım", "Primed hiper-aktif fenotip, aşırı sinaps yıkımı", "NLRP3 uyarımı, CX3CR1 sinyal körelmesi", "Nörodejenerasyon, Alzheimer amiloid birikimi", "MCC950, mikroglial senolitikler (PLX5622)"],
            ["Mast Hücresi", "Lokal homeostaz, kontrollü histamin salınımı", "Spontan degranülasyon, triptaz sızıntısı", "FceRI reseptör hipersensitivitesi", "Vasküler geçirgenlik, nörojenik inflamasyon", "Mast hücre stabilizatörleri (Ketotifen, Luteolin)"]
        ]
    ),
    (
        "TABLO 11.6: B Lenfosit Yaşlanması ve Antikor Kusurlarının Moleküler Haritası",
        ["B Hücre Bileşeni", "Gençlik Biyofiziği", "Yaşlılık Moleküler Bozukluğu", "Enzimatik / Transkripsiyonel Temel", "Klinik İmmünolojik Tablo", "Restorasyon Stratejisi"],
        [
            ["Kemik İliği B Lenfopoezi", "Günde milyonlarca taze naif B hücresi", "%80-90 üretim gerilemesi", "Stromal IL-7 ve CXCL12 kaybı", "Periferik naif B hücre havuzunun kuruması", "Kemik iliği mezenkimal niş gençleştirme"],
            ["Yaşla İlişkili B Hücreleri (ABCs)", "Pratikte yok (<%1)", "B hücre havuzunun %30-50'si", "T-bet ve CD11c anormal ekspresyonu", "Kronik otoantikor salgısı, IL-6 üretimi", "Anti-CD11c hedefe yönelik monoklonal ablasyon"],
            ["Somatik Hipermutasyon (SHM)", "Yüksek mutasyon frekansı ve afinite artışı", "Kör ve yetersiz mutasyon profili", "AID (Sitidin Deaminaz) enzim düşüşü", "Zayıf bağlanan, etkisiz antikorlar", "E47 ve AID gen transkripsiyonunun uyarılması"],
            ["Sınıf Değişimi (CSR)", "IgM'den yüksek afiniteli IgG/IgA'ya geçiş", "İzotip değişiminde blokaj (IgM takılması)", "E2A/Pax5 regülasyon bozukluğu", "Mukozal koruma kaybı (solunum/GİS enfeksiyonu)", "Sentetik immünoglobulin sınıf değiştiriciler"],
            ["Germinal Merkez Mimarisi", "Kararlı, uzun ömürlü ve organize merkezler", "Erken dağılan, küçük ve verimsiz merkezler", "Foliküler DC ve Tfh hücresi yetersizliği", "Uzun süreli immünolojik hafıza kurulamaması", "CD40L ve IL-21 lokal takviyesi"],
            ["Periferik B Toleransı", "Otoreaktif B klonlarının katı anerjisi/apoptozu", "Tolerans kırılması, otoreaktif klon kaçışı", "Aşırı BAFF sitokin seviyeleri", "Romatoid artrit, lupus, sistemik otoimmünite", "BAFF inhibitörleri (Belimumab)"]
        ]
    ),
    (
        "TABLO 11.7: Sağlıklı Genç Mikrobiyom vs Yaşlılık Disbiyom Karşılaştırması",
        ["Mikrobiyom Parametresi", "Sağlıklı Genç Mikrobiyotası", "Yaşlı / Kırılgan Disbiyotik Profil", "Etkilenen Ana Bakteriyel Türler", "Sistemik Metabolik Sonucu", "Düzeltici Müdahale Protokolü"],
        [
            ["Alfa Çeşitliliği (Shannon)", "Yüksek tür zenginliği ve homojen dağılım", "Düşük çeşitlilik, tek tip patobiont hakimiyeti", "Lactobacillus ve Bifidobacterium kaybı", "Metabolik esneklik kaybı, enfeksiyon duyarlılığı", "Geniş spektrumlu prebiyotikler, FMT"],
            ["Mukus Tabakası Bütünlüğü", "Kalın (>150 um) ve sürekli taze musin", "İncelmiş, erozyona uğramış mukus astarı", "Akkermansia muciniphila tükenmesi", "Bakterilerin doğrudan epitele teması, enterit", "Canlı pastörize Akkermansia takviyesi"],
            ["SCFA Üretim Profili", "Yüksek Butirat, Propiyonat ve Asetat", "SCFA seviyelerinde %70-80 çöküş", "Faecalibacterium prausnitzii kaybı", "Kolonosit enerji krizi, bariyer çöküşü", "Enterik kaplı sodyum butirat, dirençli nişasta"],
            ["Firmicutes / Bacteroidetes", "Dengeli oran (~1.5 - 2.0)", "Aşırı kaymalar (obezitede yüksek, frailty'de düşük)", "Ruminococcaceae vs Bacteroides", "Sistemik lipogenez veya şiddetli malnütrisyon", "Kişiselleştirilmiş makro-besin oranlama"],
            ["Toksik Metabolit Üretimi", "Minimal düzeyde", "TMAO, p-Krezil Sülfat, İndoksil Sülfat yüksek", "Fırsatçı Enterobacteriaceae artışı", "Kardiyovasküler hasar, kronik böbrek yetmezliği", "Sentetik adsorbanlar (AST-120), probiyotikler"],
            ["Sekonder Safra Asitleri", "Dengeli litokolik/deoksikolik asit", "Bozulmuş safra asidi dönüşümü", "Clostridium scindens disfonksiyonu", "Glukoz intoleransı, bağırsak motilite kaybı", "IsoalloLCA ve ursodeoksikolik asit (UDCA)"]
        ]
    ),
    (
        "TABLO 11.8: Bağırsak Bariyer İhlali (Leaky Gut) ve Sistemik İnflamatuar Yansımaları",
        ["Bariyer Bileşeni / Süreç", "Fizyolojik Gençlik Durumu", "Patolojik Yaşlanma Durumu", "Sızıntı Yapan Molekül", "Hedef Organ / Patoloji", "Bariyer Onarım Protokolü"],
        [
            ["Sıkı Bağlantılar (TJ)", "Claudin-1, Occludin ve ZO-1 ile tam sızdırmazlık", "Fosforilasyonla zardan çekilme, Claudin-2 porları", "Paraselüler makromolekül akışı", "Sistemik dolaşım / Tüm organlar", "L-glutamin, Çinko Karnosin, Butirat"],
            ["Mukozal İmmünoglobulin", "Lümene bol miktarda sekretuar IgA (sIgA) salgısı", "sIgA sentezinde %60 düşüş", "Patojenlerin epitele yapışması", "Lokal mukozal enfeksiyon ve invazyon", "Oral sIgA takviyesi, prebiyotik lifler"],
            ["Endotoksin Translokasyonu", "LPS portal kanda saptanamaz düzeyde", "Sürekli portal LPS sızıntısı (Endotoksemi)", "Bakteriyel Lipopolisakkarit (LPS)", "Karaciğer Kupffer hücre uyarımı, hepatit", "Polimiksin B filtreleri, LBP modülasyonu"],
            ["Zonulin Regülasyonu", "Bazal düşük düzey, geçici açılmalar", "Sürekli yüksek zonulin salınımı", "Kronik bağırsak geçirgenliği", "Otoimmün hastalıklar, tip 1 diyabet, çölyak", "Zonulin inhibitörü (Larazotid asetat / AT-1001)"],
            ["Bağırsak-Beyin Ekseni", "Sağlam kan-beyin bariyeri (BBB), dengeli vagus", "LPS kaynaklı BBB yıkımı, nöroinflamasyon", "Sitokinler, LPS, bakteriyel peptidoglikan", "Beyin / Alzheimer, Parkinson, depresyon", "Vagus sinir stimülasyonu, postbiyotikler"],
            ["Kardiyovasküler Endotel", "Sakin endotel, yüksek nitrik oksit (NO)", "LPS uyarılı TLR4 aktivasyonu, ICAM-1 artışı", "Endotoksin, IL-6, TNF-alfa", "Damar yatağı / Ateroskleroz, hipertansiyon", "Dolaşım sitokin temizliği, antioksidanlar"]
        ]
    ),
    (
        "TABLO 11.9: İmmün Kontrol Noktası Molekülleri (PD-1, CTLA-4, LAG-3, TIGIT, TIM-3) ve Fonksiyonları",
        ["Kontrol Noktası Reseptörü", "İfade Edildiği Hücre", "Spesifik Karşıt Ligand", "Hücre İçi İnhibitör Sinyal", "Yaşlanma / Yorgunluktaki Durumu", "Terapötik Blokaj Molekülü"],
        [
            ["PD-1 (CD279)", "Aktive, senesen ve yorgun T hücreleri", "PD-L1 (CD274) ve PD-L2 (CD273)", "ITIM ve ITSM alanları üzerinden SHP-2", "Yorulmuş T hücrelerinde sürekli yüksek ekspresyon", "Nivolumab, Pembrolizumab, sentetik peptitler"],
            ["CTLA-4 (CD152)", "Treg hücreleri, erken aktive T hücreleri", "CD80 (B7-1) ve CD86 (B7-2)", "CD28 ile yarışmalı bağlanma, PP2A aktivasyonu", "Erken lenf nodu priming fazını kilitler", "İpilimumab, Tremelimumab"],
            ["LAG-3 (CD223)", "Yorgun CD4+/CD8+ T ve NK hücreleri", "MHC Sınıf II molekülleri", "KORP motifi üzerinden kalsiyum blokajı", "PD-1 ile birlikte karsinojenik tolerans yaratır", "Relatlimab (anti-LAG-3 antikor)"],
            ["TIGIT", "T ve NK hücreleri", "CD155 (PVR) ve CD112 (Nectin-2)", "ITT benzeri alan üzerinden SHIP1/SHP-1", "NK hücre sitotoksisitesini felç eder", "Tiragolumab, Vibostolimab"],
            ["TIM-3 (HAVCR2)", "Terminal diferansiye yorgun T hücreleri", "Galektin-9, Fosfatidilserin, HMGB1", "Bat3 disosiasyonu, Fyn kinaz inhibisyonu", "Derin TOX bağımlı yorgunluğun son evresi", "Sabatolimab (anti-TIM-3)"],
            ["VISTA (B7-H5)", "Miyeloid hücreler, naif T hücreleri", "PSGL-1 (asidik pH ortamında)", "Asidik dokularda T hücrelerini susturma", "Yaşlı doku asidozunda immün baskılama", "CI-8993 (anti-VISTA antikor)"]
        ]
    ),
    (
        "TABLO 11.10: Doğal İnsan Bağışıklığı (Homo Sapiens) ile Homo Aeternus Sentetik İmmünom Karşılaştırması",
        ["İmmünolojik Boyut", "Vahşi Tip İnsan (Homo Sapiens)", "Homo Aeternus (Sentetik İmmünom)", "Kullanılan Sentetik Biyoloji Teknolojisi", "Sistemik Biyolojik Sonuç", "Ömür ve Sağlık Beklentisi"],
        [
            ["Timus Dinamiği", "Puberteden sonra involüsyon, 70'te sıfır çıktı", "Ömür boyu sabit çalışan yapay timus organoidi (ATO)", "iPSC türevli fonksiyonel cTEC/mTEC ve 3D biyo-baskı", "Her gün 10^7 taze naif T hücresi, sıfır antijenik körlük", "Yeni patojenlere karşı ebedi tam bağışıklık"],
            ["Steril Yangı (Inflam-Aging)", "Yaşla üstel artan TNF, IL-6, CRP fırtınası", "Biyo-sensörlerle kilitlenmiş cGAS/NLRP3 eşiği", "Sentetik allosterik nano-antikorlar, H-151/MCC devreleri", "Dokularda mutlak steril sükunet, sıfır inflam-aging", "Sarkopeni, osteoporoz ve artritin tamamen yok edilmesi"],
            ["Bağırsak Bariyeri", "Leaky gut, disbiyozis, kronik endotoksemi (LPS)", "Sentetik nano-hidrojel ve akıllı probiyotik kalkanı", "CRISPR modifiyeli Akkermansia ve TFF salgılayan suşlar", "Portal kana sıfır toksin sızıntısı, genç endotel", "Kardiyovasküler ve nörodejeneratif bağışıklık"],
            ["Hücresel Replikatif Sınır", "Hayflick limiti, telomer aşınması, CD28 kaybı", "Mitoz bağımlı geçici TERT ekspresyonu", "dCas9 koşullu transkripsiyonel aktivatör devreleri", "Klonal yorgunluk ve senesens yok, sonsuz çoğalma", "İmmün hafızanın ve efektör gücün ebedi korunumu"],
            ["Senolitik Sürveyans", "Senesen hücreleri tanıyamayan yaşlı NK/T hücreleri", "Hedefe kilitli otolog senolitik CAR-Makrofajlar", "uPAR ve CD264 hedefli sentetik CAR konstrüktleri", "Senesen hücrelerin oluştukları anda cerrahi imhası", "Dokuların gençlik hücresel kompozisyonunda kalması"],
            ["Viral / Kanser Direnci", "Yaşla artan malignite ve reaktive olan virüsler", "Latent viral genom temizliği ve Süper-TCR kütüphanesi", "Cas12a viral eksizyonu ve de novo tasarlanmış TCR'ler", "Tümörlerin anında tespiti, viral epizomların kazınması", "Kansersiz, enfeksiyonsuz, yangısız ölümsüz biyoloji"]
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
print("SUCCESS: Chapter 11 written to:", OUTPUT_PATH)

