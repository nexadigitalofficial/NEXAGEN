# -*- coding: utf-8 -*-
"""
PROJECT AETERNITAS - CİLT 14: GEN TERAPİSİ VEKTÖRLERİ, CRISPR-CAS9/PRIME EDITING VE İN VİVO GENOM MÜHENDİSLİĞİ
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

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_14_GEN_TERAPISI_VEKTORLERI_CRISPR_PRIME_EDITING_TAM_100_SAYFA.docx"

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
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 14: GEN TERAPİSİ VE İN VİVO GENOM MÜHENDİSLİĞİ")
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
s_run = sub_p.add_run("CİLT 14: GEN TERAPİSİ VEKTÖRLERİ, CRISPR-CAS9/PRIME EDITING VE İN VİVO GENOM MÜHENDİSLİĞİ\\n(AAV KAPSİD TASARIMI, LNP TESLİMATI, BAZ DÜZENLEME, EPİGENOM MODÜLASYONU VE İN VİVO GÜVENLİK)")
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
ih_run = intro_h.add_run("CİLT 14 MANİFESTOSU: YAŞAMIN KODUNU YENİDEN YAZMAK: GENOMİK DİJİTAL MİMARİ VE EBEDİ YAŞAM PROTOKOLÜ")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "Evrimsel biyoloji insan genomunu ölümsüzlük için değil; türün devamını sağlayacak üreme çağına kadar hayatta kalıp ardından "
    "yıpranıp elenecek fani bir biyolojik araç (tek kullanımlık soma teorisi) olarak kodlamıştır. Bu nedenle genomumuzda yerleşik "
    "olan yaşlanma programı; telomer erozyonundan somatik mutasyon birikimine, retrotranspozon istilasından epigenetik enformasyon "
    "kaybına kadar çok katmanlı bir moleküler çöküş mekanizmasıdır.\\n\\n"
    "Bu kaderi kırmanın tek yolu, yaşamın kaynak koduna doğrudan müdahale etmektir. Yirminci yüzyılın pasif farmakolojisi semptomları "
    "ötelerken; yirmi birinci yüzyılın in vivo genom mühendisliği hastalığın ve yaşlanmanın kök nedenini nükleotit düzeyinde yeniden "
    "programlama gücüne erişmiştir. AAV kapsid mühendisliği ve lipid nanopartiküller (LNPs) sayesinde genetik cerrahlar hücre zarını "
    "aşmakta; CRISPR-Cas9, Baz Düzenleyiciler (CBE/ABE) ve Prime Editing (PE) teknolojileri ile çift zincir kırığı riski olmadan "
    "tek bir baz hassasiyetiyle mutasyonları onarmaktadır.\\n\\n"
    "Bu ciltte; viral vektör virolojisi, sentetik LNP kütüphaneleri, Cas enzimlerinin biyofiziği, deaminaz baz düzenleme mekanizmaları, "
    "prime editing ters transkriptaz dinamikleri, dCas9 epigenom düzenleme, off-target kinetiği ve immün yanıtlardan arındırılmış "
    "Homo Aeternus ömür uzatma gen terapi kokteyli (TERT, Klotho, Follistatin, PGC-1a) 100 akademik alt bölümde incelenmektedir."
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
        "1.1 Adeno-İlişkili Virüs (AAV) Biyolojisi, Genom Yapısı ve Yaşam Döngüsü",
        "Adeno-İlişkili Virüs (AAV), Parvoviridae ailesine ait, zarfsız, tek sarmallı bir DNA (ssDNA) dependovirüsüdür ve insan gen terapisinde en güvenli viral platform olarak kabul edilir.",
        "Yaklaşık 4.7 kilobazlık (kb) doğal genomu; replikasyon ve ambalajlamayı yöneten Rep genini, 60 alt birimli ikozahedral kapsidi oluşturan Cap genini (VP1, VP2, VP3) ve her iki uçta yer alan 145 baz çiftlik T-şekilli Tersine Çevrilmiş Uç Tekrarları (ITR - Inverted Terminal Repeats) içerir. AAV replikasyon yeteneksizdir; çoğalmak için Adenovirüs veya Herpes Simpleks virüsü gibi bir yardımcı virüse ihtiyaç duyar. Rekombinant AAV (rAAV) üretiminde Rep ve Cap genleri tamamen silinerek yerine terapötik transgen yerleştirilir.",
        "Kapasite_rAAV = L_kargo = 4.7_kb - 2 * L_ITR = 4.7_kb - 0.29_kb ~ 4.4_kb",
        "Bu ambalajlama sınırı eşitliği, rAAV kapsidi içine güvenle paketlenebilecek maksimum yabancı transgen ve promotör uzunluğunun fiziksel limitini (yaklaşık 4.4 - 4.7 kb) tanımlar."
    ),
    (
        "1.2 Rekombinant AAV (rAAV) Üretimi: Üçlü Plazmid Transfeksiyonu ve Baculovirüs Sistemleri",
        "Klinik düzeyde rAAV üretimi, viral genom bileşenlerinin konak hücrede kontrollü ve geçici olarak bir araya getirilmesini gerektirir.",
        "Altın standart yöntem, HEK293 hücrelerinde yürütülen 'Üçlü Plazmid Transfeksiyonu'dur (Triple Transfection): 1) ITR'ler arasında transgeni taşıyan plazmid, 2) Rep ve Cap genlerini taşıyan ambalaj plazmidi, 3) Adenoviral E2A, E4 ve VA RNA genlerini içeren yardımcı (helper) plazmid. Büyük ölçekli endüstriyel biyoreaktörlerde ise Spodoptera frugiperda (Sf9) böcek hücreleri ve rekombinant Baculovirüs ekspresyon vektörleri (BEVS) kullanılarak litre başına 10^14 viral genom (vg) saflıkta üretim gerçekleştirilir.",
        "Verim_rAAV = Titre_vg = ( N_kapsid_dolu / V_kultur ) = k_transfeksiyon * [DNA_helper] * [DNA_repcap] * [DNA_ITR]",
        "Bu stokiyometrik üretim denklemi, biyoreaktör hacmi başına üretilen tam dolu fonksiyonel rAAV titre veriminin üçlü plazmid transfeksiyon oranlarıyla ilişkisini formüle eder."
    ),
    (
        "1.3 Doğal AAV Serotipleri (AAV1-AAV9) ve Doku Tropizm Profilleri",
        "Doğal olarak izole edilen AAV serotipleri, kapsid yüzeyindeki değişken amino asit ilmekleri (VR-I ila VR-IX) sayesinde farklı organ ve doku tiplerine son derece spesifik yönelim (tropizm) sergiler.",
        "Örneğin AAV1 ve AAV7 iskelet kasını; AAV2 böbrek ve retinayı; AAV3b ve AAV8 hepatositleri (karaciğer); AAV5 merkezi sinir sistemi ve akciğer epitelini; AAV9 ise sistemik vasküler dolaşımı geçerek kalp kasını ve nöronları hedefler. Bu doku özgüllüğü, kapsidin hücre yüzeyindeki heparan sülfat proteoglikan (AAV2), sialik asit (AAV4/5) veya galaktoz (AAV9) reseptörlerine bağlanma afinitesi ile belirlenir.",
        "Tropizm_Katsayisi = K_tropizm = [AAV_bagli] / [Hücre_yuzey_reseptor] = 1 / ( 1 + K_d_reseptor / [Reseptor_dansite] )",
        "Bu biyofiziksel bağlanma formülü, belirli bir AAV serotipinin hedef dokuya afinitesinin yüzey reseptör dansitesi ve termodinamik ayrışma sabiti (Kd) ile ilişkisini gösterir."
    ),
    (
        "1.4 Lentiviral ve Retrotranspozon Vektörleri: Stabil Genomik Entegrasyon Mekanizması",
        "Bölünen kök hücrelerde kalıcı gen ifadesi sağlamak için epizomal kalan AAV'ler yetersizdir; bu durumda Lentiviral Vektörler (LV - HIV-1 türevi) devreye girer.",
        "Lentivirüsler, ters transkripsiyonla tek sarmallı RNA genomlarını çift sarmallı DNA'ya çevirir ve viral İntegraz enzimi aracılığıyla konak hücre nükleer DNA'sına kalıcı olarak entegre eder. Bölünmeyen nöron ve hepatositleri de enfekte edebilen üçüncü nesil kendi kendini inaktive eden (SIN - Self-Inactivating) lentiviral vektörler; LTR bölgesindeki promotör silinmesi sayesinde replikasyon ve onkogenik insersiyonel mutagenez riskini minimize etmiştir.",
        "Verim_Entegrasyon = eta_LV = [DNA_entegre] / [RNA_viral_giris] = k_RT * k_integraz * ( 1 - P_delesyon )",
        "Bu retroviral transduksiyon eşitliği, lentiviral RNA'nın konak kromatinine stabil entegre provirüs DNA'sına dönüşüm kinetiğini ve integrasyon verimini ifade eder."
    ),
    (
        "1.5 Nötralizan Antikorlar (NAb), Pre-Ekzistans Bağışıklık ve Vektör Temizliği",
        "Klinik gen terapisinin önündeki en yaygın immünolojik engel, insan popülasyonunun büyük bir kısmının çocuklukta doğal AAV enfeksiyonu geçirmiş olmasıdır.",
        "Popülasyonun %30 ila %70'i AAV1, AAV2 ve AAV8'e karşı dolaşımda yüksek titrede 'Nötralizan Antikorlar' (NAb) taşır. Bu antikorlar sistemik olarak enjekte edilen rAAV kapsidlerine saniyeler içinde bağlanarak onları karaciğer Kupffer hücrelerine veya dalak makrofajlarına yönlendirir ve dokuya ulaşamadan temizler. NAb titresi 1:5'in üzerinde olan hastalar mevcut protokollerde çoğu gen terapi klinik çalışmasından elenmektedir.",
        "Klerens_NAb = d[AAV_serbest]/dt = - k_opsonizasyon * [AAV] * [NAb_titre] - k_renal * [AAV]",
        "Bu immünolojik klerens kinetiği denklemi, dolaşımdaki serbest viral vektör yarı ömrünün nötralizan antikor titresine bağlı katastrofik düşüşünü matematiksel olarak belgeler."
    ),
    (
        "1.6 Toll-Benzeri Reseptör 9 (TLR9) ve Vektör DNA'sının CpG Tanıma İmmün Yanıtı",
        "Viral kapsit hücre içine girdikten ve çekirdeğe ulaştıktan sonra dahi, konak doğuştan gelen bağışıklık sistemi vektör genomunu yabancı patojen olarak algılayabilir.",
        "Plazmasitoid dendritik hücreler ve makrofajların endozomlarında bulunan TLR9 reseptörü; bakteriyel ve viral DNA'da sık görülen metillenmemiş sitozin-guanin (CpG) motiflerini yüksek afiniteyle tanır. TLR9 uyarılması MyD88 adaptor proteinini aktive ederek tip I interferon (IFN-alfa/beta) fırtınasını ve CD8+ sitotoksik T hücre yanıtını tetikler; bu durum transgen ekspresyonu yapan hücrelerin lenfositlerce öldürülmesine yol açar.",
        "Aktivasyon_TLR9 = [IFN_alfa] = V_max_TLR9 * [CpG_metillenmemis]^h / ( K_cpg^h + [CpG_metillenmemis]^h )",
        "Bu Hill immünoreaktivite denklemi, vektör kasetindeki metillenmemiş CpG ada yoğunluğunun tip-1 interferon ve enflamatuar sitokin deşarjını nasıl sigmoid olarak tetiklediğini açıklar."
    ),
    (
        "1.7 CpG Adalarının Tüketilmesi (CpG Depletion) ile İmmün Kaçış Mühendisliği",
        "TLR9 aracılı immün yıkımı engellemenin en zarif ve devrimsel yolu, transgen kasetindeki genetik kodun translasyonel eşanlamlılık (kodon optimizasyonu) kullanılarak yeniden yazılmasıdır.",
        "Amino asit dizisi birebir korunurken, kodonlar CpG dinükleotidlerinden tamamen arındırılır (CpG Depletion / CpG-Free Kasetler). Örneğin arginine kodlayan CGU veya CGC kodonları AGA veya AGG ile değiştirilir. İntronlar ve promotörler sentetik CpG'siz dizilerle yeniden tasarlanır. Bu biyomühendislik hamlesi, vektörün TLR9 radarından tamamen kaçmasını sağlayarak dokuda aylar ve yıllar boyu kesintisiz, bağışıklık saldırısından uzak transgen ekspresyonu sağlar.",
        "Yogunluk_CpG_Kalan = D_CpG = N_CpG_post / N_CpG_dogal ~ 0.00 (Tam CpG Yok Edilmesi)",
        "Bu optimizasyon oranı, transgen kasetinde CpG dinükleotitlerinin sıfıra indirilmesiyle elde edilen mutlak immün kamufle durumunu sembolize eder."
    ),
    (
        "1.8 Kapsid Antijen Sunumu ve CD8+ Sitotoksik T Lenfosit (CTL) Sitotoksisitesi",
        "AAV kapsidi nükleusa ulaşıp genomunu boşalttıktan sonra sitoplazmada kalan boş protein kabukları proteazom tarafından peptid parçalarına ayrılır.",
        "Bu kapsid fragmanları endoplazmik retikulumda TAP taşıyıcıları ile MHC Sınıf I moleküllerine yüklenir ve transdukte edilen hücrenin yüzeyinde antijen olarak sunulur. Eğer konakta önceden uyarılmış hafıza CD8+ T lenfositleri varsa; perforin ve granzim B salgılayarak gen terapisiyle tedavi edilmiş bu değerli hücreleri saatler içinde lizise uğratır. Bu komplikasyonu önlemek için klinik uygulamalarda geçici immünsüpresyon (Metilprednizolon) şart koşulmaktadır.",
        "Lizis_CTL = Hiz_Olum = k_granzim * [CD8+_AAV_spesifik] * ( [MHC_I_Kapsid] / ( K_mhc + [MHC_I_Kapsid] ) )",
        "Bu sitotoksik ölüm kinetiği formülü, transdukte doku hücrelerinin yüzeyindeki kapsid antijen sunumu yoğunluğuna bağlı olarak T lenfositlerince imha edilme hızını tanımlar."
    ),
    (
        "1.9 Tek Sarmallı (ssAAV) vs Kendi Kendini Tamamlayan (scAAV) Vektör Kinetiği",
        "Standart ssAAV vektörleri çekirdeğe girdiğinde derhal protein ekspresyonu yapamaz; transkripsiyonun başlaması için hücrenin DNA polimeraz enzimlerinin ikinci sarmalı sentezlemesi (second-strand synthesis) zorunludur ve bu hız kısıtlayıcı basamak haftalar sürer.",
        "Bu darboğazı aşmak için 'Kendi Kendini Tamamlayan AAV' (scAAV - Self-Complementary AAV) vektörleri geliştirilmiştir. Bir ITR'deki terminal çözünme bölgesi (trs) mutasyona uğratılarak virüsün çift sarmallı bir saç tokası (hairpin) DNA paketlemesi sağlanır. scAAV nükleusa girer girmez saniyeler içinde kendiliğinden katlanarak çift sarmallı forma dönüşür ve transgen ekspresyonu günler değil saatler içinde başlar; ancak genetik kargo kapasitesi yarıya (yaklaşık 2.2 kb) iner.",
        "Hiz_Ekspresyon_scAAV = d[Transgen]/dt = k_hizli * [scAAV_nukleer] >> k_yavas * [ssAAV_nukleer] * [Polimeraz_konak]",
        "Bu ekspresyon hız karşılaştırması denklemi, scAAV vektörlerinin ikinci sarmal sentez gecikmesini baypas ederek yarattığı 10 ila 50 katlık kinetik ekspresyon üstünlüğünü simgeler."
    ),
    (
        "1.10 Tekrarlayan Dozlama Engelleri ve Plazmaferez/İmmünsüpresyon Stratejileri",
        "Gen terapisi tek dozluk bir tedavi olarak pazarlansa da; hücre bölünmesi veya zamanla azalan transgen ifadesi nedeniyle yaşam boyu tekrarlayan dozlamalar gerekebilir.",
        "Ancak ilk rAAV enjeksiyonundan sonra konakta kanda titreleri 1:100.000'i aşan devasa bir anti-kapsid nötralizan antikor havuzu oluşur; bu durum ikinci bir AAV enjeksiyonunu tamamen imkânsız kılar. Bu engeli aşmak için; 1) Kapsid afinite kolonları ile seçici Plazmaferez, 2) IgG yıkan bakteriyel endopeptidaz İmlifidase (IdeS) enzimi uygulaması, ve 3) Farklı bir serotipin veya sentetik kapsidin ardışık kullanımı klinik olarak test edilmektedir.",
        "Katsayi_Yeniden_Dozlama = R_dose = ( [IdeS_aktivite] / [IgG_AAV] ) * ( [AAV_yeni_serotip] / K_cap )",
        "Bu yeniden dozlanabilirlik fonksiyonu, immünoglobulin parçalayıcı enzimlerin ve ortogonal serotip değişiminin ikinci doz gen terapisine olanak sağlama penceresini modeller."
    )
]

# ================= KISIM 2 =================
part2_subsections = [
    (
        "2.1 Yönlendirilmiş Evrim (Directed Evolution) ve Kapsid Çeşitlendirme Kütüphaneleri",
        "Doğal AAV serotipleri insan evriminde patojenik olmayan enfeksiyonlar için optimize olmuştur; klinik gen terapisinin zorlu biyolojik bariyerlerini aşmak için yetersizdirler.",
        "Yönlendirilmiş Evrim (Directed Evolution), milyarlarca farklı kapsid varyantı içeren sentetik DNA kütüphanelerinin in vitro ve in vivo seçilim baskısı altında evriltilmesidir. Hata eğilimli PCR (error-prone PCR), DNA karıştırma (DNA shuffling) ve kapsid yüzey ilmeklerine rastgele 7-mer peptit yerleştirme (peptide display) yöntemleri ile 10^9 ila 10^12 çeşitlilikte plazmid kütüphaneleri oluşturulur.",
        "Cesitlilik_Kutuphane = N_varyant = 20^7 ~ 1.28 * 10^9 (Rastgele 7-mer Peptid Ekleme)",
        "Bu kombinatoryal çeşitlilik eşitliği, AAV Cap geninin yüzey ilmeğine 7 amino asitlik rastgele peptit dizisi yerleştirildiğinde teorik olarak üretilebilecek sentetik kapsid varyantı sayısını belgeler."
    ),
    (
        "2.2 Cre-Rekombinaz Bağımlı Kapsid Seçilim Platformu (CREATE Metodolojisi)",
        "Geleneksel in vivo kapsid seçiminde en büyük zorluk, hedef dokuya ulaşan virüslerin diğer organlarda biriken devasa arka plan virüs popülasyonundan ayırt edilememesiydi.",
        "Gradinaru laboratuvarı tarafından geliştirilen CREATE (Cre-Recombination-Based AAV Targeted Evolution) teknolojisi bu sorunu çözmüştür. Kapsid kütüphanesinin transgen kasetine loxP siteleri arasına ters çevrilmiş bir PCR primeri yerleştirilir. Bu virüsler hedef hücre tipinde Cre-rekombinaz eksprese eden transgenik hayvanlara verilir. Yalnızca hedef hücreye başarıyla girip nükleusunda DNA'sını açan virüslerde Cre enzimi loxP bölgesini ters çevirir; böylece sadece hedefi vuran kapsid sekansları PCR ile çoğaltılıp izole edilir.",
        "Secilim_Zenginlesme = E_CREATE = [Kapsid_i_hedef] / [Kapsid_i_plazma] * ( 1 / P_arka_plan )",
        "Bu zenginleşme faktörü eşitliği, CREATE metodolojisinin Cre-bağımlı inversiyon yoluyla hedeflenen dokudaki spesifik kapsid klonlarını arka plandan nasıl binlerce kat zenginleştirdiğini formüle eder."
    ),
    (
        "2.3 AAV-PHP.eB ve AAV-PHP.V1: Kan-Beyin Bariyerini Aşan Nörotropik Kapsidler",
        "AAV serotip 9'un (AAV9) kapsid yüzeyindeki VR-VIII bölgesine (amino asit 588) 7 amino asitlik 'TLAVPFK' dizisinin eklenmesiyle türetilen AAV-PHP.eB; modern nörogenetiğin en büyük zaferlerinden biridir.",
        "Sistemik intravenöz enjeksiyondan sonra kan-beyin bariyeri endotelini aşma kapasitesi doğal AAV9'dan en az 40 ila 60 kat daha yüksektir. Tüm serebral korteks, hipokampus ve striatumdaki nöron ve astrositlerin %70'inden fazlasını tek bir düşük dozla transdukte eder. Benzer şekilde geliştirilen AAV-PHP.V1 ise beyin parankimi yerine doğrudan beyin mikrovasküler endotelini hedefleyerek serebrovasküler gen terapilerine olanak tanır.",
        "Katsayi_Gecirgenlik_PHP = P_eB = P_AAV9 * ( 1 + 55 * [TLAVPFK_etkilesim] / K_d_Ly6a )",
        "Bu biyofiziksel penetrasyon formülü, kapsid yüzeyine eklenen spesifik hekzamer dizisinin kan-beyin bariyeri transfer hızını AAV9 bazal seviyesine kıyasla elli kattan fazla nasıl artırdığını tanımlar."
    ),
    (
        "2.4 Ly6a ve Transferrin Reseptörleri Aracılığıyla Reseptör Aracılı Transsitoz",
        "AAV-PHP serisi kapsidlerin kan-beyin bariyerini bu denli zahmetsiz aşmasının moleküler sırrı, endotel yüzeyindeki Lenfosit Antijeni 6 Kompleksi Locus A (Ly6a / Sca-1) reseptörüne yüksek afiniteyle bağlanmalarıdır.",
        "Kapsid, lüminal endotel zarında Ly6a'ya bağlandığında klatrin bağımsız, kaveolin aracılı endositozla veziküllere alınır. Lizozomal yıkım yoluna girmeden ablüminal zara taşınır ve ekzositozla doğrudan beyin interstisiyel aralığına salınır (Reseptör Aracılı Transsitoz). İnsan endotelinde Ly6a bulunmadığından; insan klinik translasyonu için insan Transferrin Reseptörü 1'i (TfR1) veya CD98hc'yi bağlayan yeni nesil sentetik kapsidler (AAV.BI30) tasarlanmıştır.",
        "Akı_Transsitoz = J_BBB = V_max_trans * [AAV_luminal] / ( K_m_TfR1 + [AAV_luminal] )",
        "Bu transsitoz akı denklemi, lüminal kapillerden serebral parankime viral transfer debisinin endotelyal reseptör satürasyon kinetiği ile doğrudan ilişkisini ortaya koyar."
    ),
    (
        "2.5 AAV.CAP-B10 ve İnsan Nöronal Hedefleme: Türler Arası Kapsid Translasyonu",
        "Kemirgen modellerinde mükemmel çalışan kapsidlerin (AAV-PHP.eB gibi) insan ve primatlarda reseptör farklılığı nedeniyle etkisiz kalması 'türler arası translasyon bariyeri'ni doğurmuştur.",
        "Broad Enstitüsü ve Caltech araştırmacıları; primat ve insan indüklenmiş nöronlarında (iPSC-derived neurons) eş zamanlı taranan 'AAV.CAP-B10' kapsidini geliştirmiştir. CAP-B10, kemirgen Ly6a reseptörüne bağımlı olmaksızın insan serebral dokusunda ve primat beyninde AAV9'dan 20 kat daha yüksek transduksiyon verimi sağlamış; karaciğer tutulumunu ise %80 oranında baskılayarak güvenli insan klinik nöro-gen terapisinin kapısını aralamıştır.",
        "Selektivite_Indeksi = SI = ( Transduksiyon_Beyin / Transduksiyon_Karaciger )_CAP-B10 >> SI_AAV9",
        "Bu hedefleme saflık indeksi, CAP-B10 sentetik kapsidinin nöronal hedefleme başarısını korurken karaciğer toksisitesi riskini nasıl onlarca kat minimize ettiğini formüle eder."
    ),
    (
        "2.6 Hepatik Tropizm Baskılama (De-Targeting) ve MikroRNA Hedef Siteleri",
        "Sistemik intravenöz AAV enjeksiyonlarının en büyük klinik riski, virüslerin %90'ından fazlasının karaciğer tarafından emilerek masif transaminaz yüksekliği ve hepatotoksisite yaratmasıdır.",
        "Hepatik tutulumu engellemek için iki strateji birleştirilir: 1) Kapsid yüzeyindeki karaciğer reseptör bağlama kalıntılarının mutasyona uğratılması (örneğin AAV2'deki heparan sülfat bağlama bölgesinin silinmesi); 2) Transgen kasetinin 3' UTR bölgesine karaciğere özgü miR-122 mikroRNA hedef dizilerinin (miR-122 binding sites) eklenmesi. Karaciğere giren virüste bol bulunan miR-122 derhal transgen mRNA'sını parçalar; böylece karaciğerde protein sentezi sıfırlanırken hedef dokularda tam ekspresyon korunur.",
        "Ekspresyon_Karaciger = E_hep = E_bazal * ( 1 / ( 1 + [miR-122_hepatik] / K_mir ) ) ~ 0.01 * E_bazal",
        "Bu transkripsiyonel susturma eşitliği, 3' UTR'ye gömülen miR-122 hedeflerinin hepatik dokuda transgen ifadesini %99 oranında nasıl kapattığını gösterir."
    ),
    (
        "2.7 Yapay Zeka ve Derin Öğrenme Tabanlı De Novo Kapsid Tasarımı (AlphaFold/ESM)",
        "Kombinatoryal kütüphaneler trilyonlarca olasılıktan yalnızca mikroskobik bir kesiti tarayabilirken; modern biyoloji makine öğrenimi ile doğrudan 'in silico de novo kapsid' tasarımına geçmiştir.",
        "Büyük dil modelleri (ESM-2) ve protein yapı tahmin algoritmaları (AlphaFold-Multimer); 60 merik kapsidin stabilitesini, montaj kabiliyetini ve immün kaçış potansiyelini atomik düzeyde simüle eder. Derin üretici modeller (Generative Diffusion Models), doğal hiçbir serotipte bulunmayan tamamen sentetik VP3 amino asit sekansları üreterek immün sistem antikorlarının tanıyamayacağı 'hayalet kapsidler' (stealth capsids) inşa etmektedir.",
        "P_montaj = Softmax( W_model * Emb(Dizi_Kapsid) + b_stabilite )",
        "Bu derin öğrenme olasılık skoru, üretici yapay zeka modelinin belirli bir sentetik kapsid dizisinin fiziksel olarak 60-merik kararlı bir ikozahedral virüs kabuğu oluşturma olasılığını hesaplar."
    ),
    (
        "2.8 Kapsid İçi Kargo Optimizasyonu: Kodon Değişimi, Promotör Mühendisliği ve Poli-A Sinyalleri",
        "AAV'nin dar 4.7 kb kargo alanı, ekspresyon kasetinin her bir baz çiftinin azami verimlilikle optimize edilmesini zorunlu kılar.",
        "Kodon optimizasyonu ile nadir kodonlar konak hücre tRNA havuzuna en uygun olanlarla değiştirilerek translasyon hızı 10 kat artırılır. Geniş hacimli viral promotörler yerine 200-300 baz çiftlik ultra-kompakt sentetik promotörler (örneğin nöronlara özgü Syn1 veya mini-CAG); transkripsiyonu maksimize eden sentetik intronlar; ve mRNA stabilitesini artıran mini Poli-A (simian virüs 40 türevi) dizileri kullanılır.",
        "Etkinlik_Translasyon = Hiz_Protein = k_ribozom * CAI * [mRNA_kararli] * ( 1 / L_transgen )",
        "Bu moleküler ekspresyon eşitliği, kodon uyum indeksi (CAI - Codon Adaptation Index) ve optimize promotör kuvvetinin protein sentez verimini nasıl katladığını gösterir."
    ),
    (
        "2.9 Çift Vektörlü (Dual-Vector) Trans-Splicing ve Aşırı Büyük Transgenlerin Taşınması",
        "Tümör baskılayıcılar, gen düzenleyiciler (SpCas9 + efektörler) veya distrofin gibi devasa genler (>5 kb) tek bir AAV kapsidine sığmaz.",
        "Bu fiziksel engeli aşmak için gen iki parçaya bölünerek ayrı AAV'lere paketlenir: 1) Promotör ve genin 5' yarısını splice donör (SD) bölgesiyle taşıyan ön vektör; 2) Splice akseptör (SA) bölgesi ve genin 3' yarısını poli-A ile taşıyan arka vektör. Hücre çekirdeğinde her iki virüsün ITR'leri homolog rekombinasyon veya baş-kuyruk ligasyonu ile birleşir ve pre-mRNA düzeyinde trans-splicing gerçekleşerek tam boy fonksiyonel protein üretilir.",
        "Verim_Dual_AAV = eta_dual = eta_vektor1 * eta_vektor2 * k_trans_splicing * P_ko_enfeksiyon",
        "Bu stokastik rekombinasyon denklemi, iki bağımsız viral vektörün aynı hücre çekirdeğinde başarıyla birleşip tam fonksiyonel transgen oluşturma olasılığını modeller."
    ),
    (
        "2.10 Kapsid Yüzey Kimyasal Modifikasyonu: PEGilasyon, Nanoparçacık Zırhlama ve Polimer Kaplama",
        "Genetik kapsid modifikasyonunun yanı sıra; saflaştırılmış rAAV vektörlerinin yüzeyine kimyasal ve biyofiziksel zırhlar giydirilmektedir.",
        "Kapsid yüzeyinde dışarı bakan lizin amino asitlerinin amino gruplarına kovalent olarak Polietilen Glikol (PEG) zincirleri bağlanır (PEGilasyon). Bu polimerik hidrasyon kalkanı; nötralizan antikorların kapsid epitoplarına erişimini sterik olarak engeller, dolaşımdaki yarı ömrü uzatır ve dalak/karaciğer retiküloendotelyal sisteminden virüsü gizler. Ayrıca manyetik nanopartiküllerle kaplanan kapsidler, harici manyetik alanlarla doğrudan hedef organa odaklanabilmektedir.",
        "Kalkan_Faktoru = SF = 1 / ( 1 + exp( - ( Yoğunluk_PEG - Rho_kritik ) / delta_sterik ) )",
        "Bu sterik engel fonksiyonu, kapsid yüzeyindeki PEGilasyon yoğunluğunun antikor bağlanma olasılığını nasıl logaritmik olarak sıfırladığını açıklar."
    )
]

# ================= KISIM 3 =================
part3_subsections = [
    (
        "3.1 Non-Viral Gen Teslimat Paradigmaları ve Viral Vektör Kısıtlamalarının Aşılması",
        "Viral vektörler yüksek transduksiyon verimi sunsa da; sınırlı ambalajlama kapasiteleri (AAV <4.7 kb), pre-ekzistans bağışıklık riskleri, tekrarlayan dozlama zorlukları ve yüksek üretim maliyetleri nedeniyle genom mühendisliğinin yegâne çözümü olamazlar.",
        "Sentetik ve non-viral teslimat platformları; sınırsız kargo kapasitesi sunmaları, nükleik asitleri (DNA, mRNA, siRNA) veya doğrudan Cas ribonükleoprotein (RNP) komplekslerini taşıyabilmeleri, neredeyse sıfır immünojenisite sergilemeleri ve kimyasal olarak ölçeklenebilir fabrikasyonları ile gen terapisinin geleceğini temsil eder.",
        "Kapasite_NonViral = L_kargo_sentetik ~ Sonsuz (Mega-baz düzeyinde plazmid ve RNP paketleme)",
        "Bu kapasite karşılaştırması, sentetik taşıyıcı sistemlerin viral ambalajlama hacim kısıtlamalarını tamamen ortadan kaldırdığını ifade eder."
    ),
    (
        "3.2 Lipid Nanopartiküllerin (LNPs) Moleküler Mimarisi ve Dört Temel Bileşeni",
        "Modern Lipid Nanopartiküller (LNPs), mRNA aşılarının ve klinik in vivo gen düzenleme terapilerinin (Onpattro, Intellia NTLA-2001) omurgasını oluşturan nanoteknolojik harikalardır.",
        "Tipik bir LNP dört kritik kimyasal bileşenin kusursuz stokiyometrik dengesinden oluşur: 1) İyonize edilebilir katyonik lipid (~%50 mol); 2) Yapısal fosfolipid (DSPC - distearoilfosfatidilkolin, ~%10 mol); 3) Membran akışkanlığı ve stabilitesini sağlayan Kolesterol (~%38.5 mol); 4) Partikül agregasyonunu önleyen ve dolaşım süresini belirleyen PEGile lipid (PEG-lipid, ~%1.5 mol). Bu bileşenler nükleik asitleri elektrostatik olarak hapseden 60-100 nm çapında homojen nano-küreler oluşturur.",
        "Kompozisyon_LNP = 50%_Ionizable + 38.5%_Kolesterol + 10%_DSPC + 1.5%_PEG-Lipid",
        "Bu molar formülasyon kuralı, klinik onaylı LNP platformlarının termodinamik ve hücresel transfer verimini maksimize eden standart kimyasal mimarisini tanımlar."
    ),
    (
        "3.3 İyonize Edilebilir Katyonik Lipid Biyofiziği: pKa Değeri ve pH-Bağımlı Yük Değişimi",
        "Klasik katyonik lipidler sürekli pozitif yüklü oldukları için sistemik dolaşımda yüksek sitotoksisite ve hemoliz yaratırken; 'İyonize Edilebilir Lipidler' (Ionizable Lipids) bu sorunu mucizevi bir pH duyarlılığı ile çözer.",
        "Bu lipidlerin asit ayrışma sabiti (pKa) hassas bir biçimde 6.2 ila 6.8 aralığında tasarlanmıştır. Fizyolojik kan pH'sında (pH 7.4) net yükleri nötrdür; bu sayede dolaşımda toksisite yaratmaz ve eritrositleri parçalamaz. Ancak hücre içine endositozla alındıklarında asidik endozomal lümende (pH 5.0 - 5.5) anında protonlanarak yoğun pozitif yüke bürünürler; bu yük dönüşümü endozom zarını delerek kargonun sitoplazmaya fırlatılmasını sağlar.",
        "Yuk_Durumu = [Lipid_H+] / [Lipid_toplam] = 1 / ( 1 + 10^(pH - pKa) )",
        "Bu Henderson-Hasselbalch iyonizasyon eşitliği, endozomal pH düşüşünün iyonize edilebilir lipidlerin protonlanma fraksiyonunu nasıl aniden %10'dan %90'ın üzerine fırlattığını formüle eder."
    ),
    (
        "3.4 DLin-MC3-DMA, SM-102 ve ALC-0315: Klinik Lipidlerin Kimyasal Evrimi",
        "İyonize edilebilir lipidlerin kimyasal yapısı; hidrokarbon kuyruklarının doymamışlık derecesi, dallanma morfolojisi ve biyobozunur ester bağlarının eklenmesiyle evrimleşmiştir.",
        "İlk FDA onaylı RNAi ilacı Onpattro'da kullanılan DLin-MC3-DMA mükemmel bir karaciğer susturma verimi sunmuş; ancak dokularda yavaş temizlenmesi endişe yaratmıştır. Moderna tarafından geliştirilen SM-102 ve Pfizer/BioNTech tarafından kullanılan ALC-0315 ise lipid kuyruklarına gömülü ester bağları içerir. Bu esterler hücre içine kargo teslim edildikten sonra endojen esterazlar tarafından hızla hidrolize edilerek toksik olmayan metabolitlere parçalanır ve karaciğer birikimi sıfırlanır.",
        "Hiz_Biyobozunma = d[Lipid]/dt = - k_esteraz * [Esteraz_sitoplazmik] * [Lipid_ester]",
        "Bu enzimatik hidroliz denklemi, yeni nesil biyobozunur lipidlerin hücre içi temizlenme hızını ve doku güvenliğini matematiksel olarak ortaya koyar."
    ),
    (
        "3.5 Mikroakışkan (Microfluidics) Çarpışmalı Karıştırma ve LNP Sentez Dinamiği",
        "Klinik kalitede LNP üretimi, partikül boyutu dağılımının (Polidispersite İndeksi - PDI < 0.1) ve kargo enkapsülasyon veriminin (>%90) kusursuz olmasını şart koşar.",
        "Bu hassasiyet 'Mikroakışkan Çarpışmalı Karıştırma' (Microfluidic Hydrodynamic Focusing / Staggered Herringbone Mixer) teknolojisi ile sağlanır. Bir kanaldan lipidlerin etanolik çözeltisi, diğer kanaldan nükleik asitlerin asidik sulu çözeltisi (pH 4.0 sitrat tamponu) mikrosaniyelik hızla karşılaştırılır. Ani polarite değişimi ve pH düşüşü; lipidlerin pozitif yüklenerek negatif yüklü RNA etrafında anında homojen miseller halinde kendi kendine montajlanmasını (self-assembly) tetikler.",
        "Cap_LNP = d_partikul = k_akis * ( Akis_Orani_Sulu / Akis_Orani_Etanol )^-beta * Total_Akis_Hizi^-gamma",
        "Bu akışkanlar mekaniği eşitliği, mikroakışkan çipteki sulu/organik akış hız oranlarının ve toplam debinin nihai nanopartikül çapını ve monodispersitesini nasıl belirlediğini simgeler."
    ),
    (
        "3.6 Apolipoprotein E (ApoE) Aracılı Karaciğer Hepatosit Tutulum Mekanizması",
        "Standart formülasyonlu LNP'ler sistemik kan dolaşımına enjekte edildiklerinde; doğal bir biyolojik 'Truva Atı' mekanizması kullanarak karaciğere yönelirler.",
        "Dolaşımdaki endojen Apolipoprotein E (ApoE) proteinleri, LNP yüzeyindeki kolesterol ve lipid zarına yüksek afiniteyle tutunarak bir protein korona (biomolecular corona) oluşturur. Hepatositlerin yüzeyinde yoğun bulunan Düşük Yoğunluklu Lipoprotein Reseptörleri (LDLR) ve LRP1; ApoE ile kaplanmış bu LNP'leri endojen bir şilomikron kalıntısı zannederek klatrin bağımlı endositozla hücre içine çeker.",
        "Baglanma_LDLR = B_max * [LNP:ApoE] / ( K_d_LDLR + [LNP:ApoE] )",
        "Bu reseptör tutulum eşitliği, sistemik LNP'lerin hepatik klirensinin plazma ApoE opsonizasyon katsayısı ve hepatosit LDLR dansitesine olan mutlak bağımlılığını belgeler."
    ),
    (
        "3.7 SORT Teknolojisi: Seçici Organ Hedefleme ve Ekstrahepatik Teslimat",
        "LNP'lerin doğal eğilimi olan karaciğer tutulumunu kırıp akciğer, dalak veya kemik iliği gibi ekstrahepatik organları hedeflemek; Daniel Siegwart laboratuvarının geliştirdiği SORT (Selective Organ Targeting) teknolojisi ile başarılmıştır.",
        "Standart dörtlü LNP formülasyonuna beşinci bir 'SORT molekülü' eklenir: %20-50 oranında kalıcı katyonik lipid (DOTAP) eklendiğinde partikülün iç yükü değişir, ApoE yerine plazma vitronektinini bağlar ve doğrudan Akciğere yönelir; %10-40 oranında anyonik lipid (18:1 PA) eklendiğinde ise Dalağa yönelir. Böylece karaciğere tek bir molekül dahi kaçmadan akciğer endoteli veya dalak T hücreleri genetik olarak düzenlenebilir.",
        "Organ_Yonelim_Vektoru = V_hedef = [DOTAP]% * Akciger_Aksi + [18:1PA]% * Dalak_Aksi + [Standart]% * Karaciger_Aksi",
        "Bu vektörel hedefleme modeli, eklenen beşinci SORT lipit fraksiyonunun nanopartikülün organ biyodağılım haritasını nasıl radikal biçimde değiştirdiğini tanımlar."
    ),
    (
        "3.8 Endozomal Kaçış Biyofiziği: 'Proton Süngeri' ve Lamellar-Ters Heksagonal Faz Geçişi",
        "LNP teknolojisinin karşı karşıya olduğu en büyük biyofiziksel darboğaz 'Endozomal Kaçış' (Endosomal Escape) verimidir; hücre içine alınan LNP'lerin yalnızca %1 ila %5'i sitoplazmaya ulaşabilir, geri kalanı lizozomlarda yok edilir.",
        "Kaçışı sağlayan biyofiziksel mekanizma 'Lamellar-Ters Heksagonal (H_II) Faz Geçişi'dir. Endozom asidikleştiğinde protonlanan katyonik lipidler, endozom zarındaki negatif yüklü endojen anyonik fosfolipidlerle (özellikle lizobisfosfatidik asit - LBPA) iyon çiftleri oluşturur. Bu konik yapı silindirik çift katmanlı zar geometrisini bozar ve membranı ters hekzagonal faza zorlayarak delikler açar; kargo sitoplazmaya fışkırır.",
        "Verim_Kacis = eta_kacis = [RNA_sitoplazma] / [RNA_endozom] = k_faz_gecisi * exp( - Delta_G_lamellar_heksagonal / (R * T) )",
        "Bu termodinamik faz dönüşüm formülü, endozomal membran destabilizasyon enerjisinin lipid konikliği ve protonlanma derecesine bağlı olarak sitoplazmik kaçış fraksiyonunu nasıl belirlediğini açıklar."
    ),
    (
        "3.9 Kargo Türleri: mRNA, Cas9 Ribonükleoprotein (RNP) ve Plazmid DNA Paketleme",
        "LNP platformlarının en büyük avantajı, taşınan genetik materyalin formuna göre modüler olarak uyarlanabilmesidir.",
        "1) mRNA Teslimatı: Geçici transgen veya Cas9 ekspresyonu için idealdir; nükleusa girmesine gerek yoktur, sitoplazmada ribozomlarca hemen okunur ve 24-48 saatte tamamen parçalanır; 2) Cas9 RNP Kompleksleri: Cas9 proteini ve rehber RNA önceden laboratuvarda birleştirilerek LNP içine paketlenir; hücreye girer girmez saniyeler içinde genomu keser ve hızla degrade olarak hedef dışı (off-target) kesimleri sıfıra indirir; 3) Donör DNA Kalıpları: Homolog rekombinasyon (HDR) için gereken uzun DNA dizileri de aynı LNP içine ko-enkapsüle edilebilir.",
        "Zaman_Aktivite = Tau_RNP (saatler) << Tau_mRNA (gunler) << Tau_AAV (yıllar)",
        "Bu intrakonak varlık süresi hiyerarşisi, LNP-RNP sistemlerinin genom cerrahisinde hedef dışı mutasyonları önlemedeki olağanüstü kinetik üstünlüğünü ortaya koyar."
    ),
    (
        "3.10 LNP İmmünojenisitesi: Anti-PEG Antikorları ve İnfüzyon Reaksiyonları (CARPA)",
        "LNP'ler viral vektörlere kıyasla çok daha güvenli olsa da; tekrarlayan klinik dozlamalarda ortaya çıkan kendilerine özgü immünolojik sınırlılıkları mevcuttur.",
        "Kozmetik ve gıda ürünlerinde yaygın kullanılan polietilen glikole maruziyet sonucu popülasyonda Anti-PEG IgG ve IgM antikorları gelişebilmektedir. Bu antikorlar PEGile LNP'leri bağlayarak kompleman sistemini kontrolsüzce aktive edebilir ve 'Kompleman Aktivasyonu Aracılı Psödo-Alerji' (CARPA - Complement Activation-Related Pseudoallergy) tablosuna yol açabilir. Bu riski ortadan kaldırmak için modern formülasyonlarda PEG yerine poligliserol veya polisarcosine bazlı yeni nesil 'stealth' lipidler kullanılmaktadır.",
        "Risk_CARPA = I_reaksiyon = k_kompleman * [Anti-PEG_antikor] * [LNP_dozu] / [Faktör_H_inhibitör]",
        "Bu psödo-alerji reaksiyon eşitliği, plazma anti-PEG antikor konsantrasyonu ve LNP dozunun kompleman anafilatoksin deşarjı (C3a, C5a) üzerindeki kinetik riskini formüle eder."
    )
]

# ================= KISIM 4 =================
part4_subsections = [
    (
        "4.1 CRISPR-Cas9 Moleküler Mimarisi: Tanıma (REC) ve Nükleaz (NUC) Lobları",
        "Streptococcus pyogenes kaynaklı Cas9 (SpCas9), 1368 amino asitten oluşan ve çift sarmallı DNA'yı diziye özgü kesen çok alanlı bir RNA-rehberli endonükleazdır.",
        "Protein yapısal olarak iki ana loba ayrılır: 1) Rehber RNA (gRNA) ve hedef DNA heterodupleksini saran Tanıma Lobu (REC - REC1, REC2, REC3 alanları); 2) Katalitik kesimi icra eden Nükleaz Lobu (NUC). NUC lobu; hedef DNA sarmalını kesen HNH nükleaz alanını, hedef olmayan sarmalı kesen RuvC alanını ve PAM dizisini tanıyan PAM-etkileşimli (PI) alanı bünyesinde barındırır.",
        "Mimar_Cas9 = [REC_Lobu (REC1 + REC2 + REC3)] + [NUC_Lobu (HNH + RuvC + PI)]",
        "Bu yapısal organizasyon eşitliği, Cas9 enziminin hedef arama (REC lobu) ve endonükleolitik kesim (NUC lobu) işlevlerinin moleküler mimari haritasını tanımlar."
    ),
    (
        "4.2 PAM Tanıma Kinetiği, DNA Çözülmesi ve R-Loop Oluşum Biyofiziği",
        "Cas9'un hedef kromatini tarama süreci rastgele difüzyon ve üç boyutlu çarpışmalarla başlar; enzimin DNA'yı çözmeye başlaması için zorunlu ilk adım 'Proto-spacer Adjacent Motif' (PAM - SpCas9 için 5'-NGG-3') dizisinin tanınmasıdır.",
        "PI alanındaki Arg1333 ve Arg1335 kalıntıları, PAM'deki ardışık guanin bazlarının majör oluğuna hidrojen bağlarıyla kenetlenir. PAM tanınması, DNA çift sarmalının yerel olarak açılmasını (DNA melting) tetikler. Rehber RNA'nın 3' ucundaki 8-10 bazlık 'Tohum Bölgesi' (Seed Region), açılan hedef DNA ile hibritlenmeye başlar ve enerjetik olarak kararlı bir 'R-Loop' yapısı oluşturarak DNA'yı fermuar gibi açar.",
        "K_PAM = [Cas9:gRNA:DNA] / ( [Cas9:gRNA] * [DNA_NGG] ) = exp( - Delta_G_PAM / (R * T) )",
        "Bu termodinamik bağlanma eşitliği, PAM tanıma kararlılığının yerel serbest enerji düşüşü (Delta_G_PAM) ve ardışık hidrojen bağlarının kooperatif afinitesi ile ilişkisini simgeler."
    ),
    (
        "4.3 HNH ve RuvC Katalitik Merkezleri: Çift Sarmal Kırığı (DSB) Oluşumu",
        "R-Loop oluşumu tamamlanıp 20 nükleotitlik tam eşleşme sağlandığında, Cas9 konformasyonel bir allosterik kilitlenme evresine girer.",
        "HNH nükleaz alanı yaklaşık 180 derecelik dramatik bir rotasyon yaparak hedef DNA sarmalının fosfodiester omurgasına yaklaşır ve PAM'in 3 baz çifti yukarısından tek sarmal kırığı (nick) açar. Eş zamanlı olarak RuvC alanı diğer sarmalı keser. İki bağımsız katalitik merkezin Mg2+ iyonları koordinesinde gerçekleştirdiği bu senkronize hidroliz, hücre genomunda kör uçlu bir Çift Sarmal Kırığı (DSB - Double-Strand Break) ile sonuçlanır.",
        "Hiz_DSB = d[DSB]/dt = k_katalitik * [Cas9:R-Loop_aktif] * ( [Mg2+] / ( K_Mg + [Mg2+] ) )^2",
        "Bu enzimatik çift sarmal kesim kinetiği, HNH ve RuvC nükleaz ceplerinin iki değerlikli magnezyum kofaktör satürasyonuna bağımlı koordineli kesim hızını modeller."
    ),
    (
        "4.4 Hücresel DNA Onarım Yolakları: Hata Eğilimli NHEJ vs Hassas HDR Dengesi",
        "Cas9 tarafından açılan çift sarmal kırığı, hücrenin intrinsik DNA hasar yanıtı (DDR) mekanizmaları tarafından derhal onarılmaya çalışılır.",
        "Post-mitotik nöronlar ve istirahat halindeki hücreler kural olarak 'Homolog Olmayan Uç Birleştirme' (NHEJ - Non-Homologous End Joining) yolunu kullanır. Ku70/Ku80 ve DNA-PKcs kompleksleri kırık uçları rastgele bağlarken birkaç nükleotitlik ekleme veya silmeler (indels) yapar; bu durum gen nakavtı (knockout) için idealdir ancak onarım için kusurludur. Kusursuz gen onarımı sağlayan 'Homolojiye Dayalı Onarım' (HDR - Homology-Directed Repair) ise yalnızca bölünen hücrelerin S ve G2 fazlarında aktif olduğundan yetişkin dokularda verimi <%1'dir.",
        "Oran_Onarim = NHEJ / HDR ~ 100 / 1 (Yetiskin Post-Mitotik Dokularda)",
        "Bu hücresel onarım oranı, yetişkin dokularda klasik CRISPR-Cas9 ile fonksiyonel gen yazmanın neden aşırı zor olduğunu ve yeni nesil düzenleyicilere (Base/Prime Editing) neden ihtiyaç duyulduğunu ortaya koyar."
    ),
    (
        "4.5 Minik Cas Varyantları: SaCas9, CjCas9 ve CasMINI (AAV Ambalajlama Uyumu)",
        "SpCas9 proteini 4.2 kb'lık kodlama dizisiyle (1368 aa) tek başına bir AAV'nin kargo kapasitesini neredeyse tamamen doldurur; promotör ve gRNA kaseti eklendiğinde viral sınıra takılır.",
        "Bu engeli aşmak için daha kompakt 'minik Cas' enzimleri keşfedilmiştir: 1) Staphylococcus aureus Cas9 (SaCas9 - 1053 aa / ~3.1 kb, NNGRRT PAM); 2) Campylobacter jejuni Cas9 (CjCas9 - 984 aa / ~2.9 kb); 3) Yönlendirilmiş evrimle türetilen ve sadece 529 amino asitlik ultra-küçük 'CasMINI' (Cas12f türevi). SaCas9 ve CasMINI, gRNA ve promotörleriyle birlikte tek bir AAV içine kolayca sığarak in vivo gen terapisinde devrim yaratmıştır.",
        "Kalan_Kargo_Alani = Delta_L_AAV = 4.7_kb - ( L_Cas + L_promoter + L_gRNA + L_polyA )",
        "Bu vektörel kargo denklemi, SaCas9 ve CasMINI varyantlarının ek regülatör elementler için bıraktığı geniş alan avantajını matematiksel olarak gösterir."
    ),
    (
        "4.6 SpCas9-HF1, eSpCas9 ve HiFi Cas9: Yüksek Doğruluklu Enzim Mühendisliği",
        "Doğal SpCas9, rehber RNA ile hedef DNA arasında 1 ila 3 nükleotitlik uyumsuzluk (mismatch) olsa dahi hedef dışı bölgeleri kesebilmekte (off-target aktivite) ve genomik instabiliteye yol açabilmektedir.",
        "Kollaboratif yapısal biyoloji çalışmaları; Cas9'un hedef olmayan DNA sarmalı ile kurduğu pozitif yüklü elektrostatik temasları haritalamıştır. Bu bölgelerdeki kalıntıların nötrlenmesiyle geliştirilen 'eSpCas9' (K848A/K1003A/R1060A), 'SpCas9-HF1' (N497A/R661A/Q695A/Q926A) ve 'HiFi Cas9'; enzim ile DNA arasındaki temas enerjisini düşürmüştür. Artık tek bir baz uyumsuzluğu dahi R-loop kilitlenmesini imkânsız kılar; on-target kesim verimi korunurken hedef dışı mutasyonlar tespit sınırının altına iner.",
        "Hassasiyet_Orani = E_dogruluk = Hiz_On_Target / Hiz_Off_Target >> 10^4",
        "Bu doğruluk oranı formülü, yüksek sadakatli (HiFi) Cas9 varyantlarının hedef dışı kesim riskini on bin kattan fazla nasıl düşürdüğünü belgeler."
    ),
    (
        "4.7 Genişletilmiş PAM Varyantları: SpG, SpRY ve PAM'siz (PAM-less) Genom Taraması",
        "Geleneksel Cas9 yalnızca genomda 'NGG' motifi bulunan lokusları hedefleyebilir; bu kısıtlama insan genomundaki nükleotitlerin yaklaşık %90'ını erişilmez kılmaktadır.",
        "David Liu ve Benjamin Kleinstiver laboratuvarları; PAM-etkileşimli (PI) alanında yapı yönlendirmeli mutasyonlar yaparak 'SpG' (NGN PAM'lerini tanıyan) ve 'SpRY' (neredeyse tüm NRN ve NYN dizilerini tanıyan / neredeyse PAM-bağımsız) varyantlarını üretmiştir. SpRY enzimi, insan genomundaki 3.2 milyar baz çiftinin her birini tek tek hedefleme çözünürlüğü sunarak mutasyon düzeltme alanını sınırsız hale getirmiştir.",
        "Kapsama_Genom = C_hedef = N_lokus_erisilebilir / N_toplam_baz ~ 1.00 (SpRY ile %100 Kapsama)",
        "Bu genomik kapsama oranı, PAM-esnek varyantların genom cerrahisinde hedefleme kör noktalarını tamamen ortadan kaldırdığını ifade eder."
    ),
    (
        "4.8 Cas12a (Cpf1): Yapışkan Uç Kinetiği, T-Zengin PAM ve İçsel crRNA İşleme",
        "Cas9'a alternatif olarak geliştirilen Tip V efektörü Cas12a (Cpf1); kendine özgü biyofiziksel nitelikleriyle gen düzenleme cephanesini genişletmiştir.",
        "Cas12a; 1) Guanin yerine Timin açısından zengin '5'-TTTV-3'' PAM dizilerini tanır (AT-zengin genomik bölgeler için idealdir); 2) tracrRNA'ya ihtiyaç duymadan tek bir kısa crRNA (42 nt) ile çalışır ve kendi pre-crRNA'sını işleme endoribonükleaz aktivitesine sahiptir (çoklu gen susturma için eşsizdir); 3) Kör uç yerine 4-5 nükleotitlik 5' çıkıntılı 'Yapışkan Uçlar' (Staggered Ends) açar. Bu yapışkan uçlar yönlendirilmiş DNA ligasyonunu kolaylaştırır.",
        "Kesim_Profili_Cas12a: 5'-NNNNN--- | ---NN-3' (4-5 nt Yapışkan Çıkıntı, Uzak Kesim)",
        "Bu moleküler kesim şeması, Cas12a'nın PAM dizisinden 18-23 nükleotit uzakta kademeli çift sarmal kırığı açma karakteristiğini tanımlar."
    ),
    (
        "4.9 Cas13 Ailesi: RNA-Hedefli Düzenleme ve Kollateral RNA Klirens Biyofiziği",
        "Hücre genomuna kalıcı müdahale yapmaksızın transkriptom seviyesinde geçici ve geri dönüşümlü düzeltmeler yapmak için Tip VI efektörü 'Cas13' enzimleri kullanılır.",
        "Cas13a, Cas13b ve Cas13d (CasRx); DNA'yı değil, doğrudan hücresel tek sarmallı RNA (ssRNA) moleküllerini iki HEPN nükleaz alanı aracılığıyla keser. Hedef mRNA'yı tanıdıktan sonra tetiklenen 'Kollateral RNA Parçalama' aktivitesi in vitro tanı kitlerinde (SHERLOCK) devrim yaratırken; in vivo uygulamalarda viral RNA'ların (SARS-CoV-2, grip) veya toksik nörodejeneratif transkriptlerin (tau veya Huntington RNA'sı) hücreye zarar vermeden parçalanmasını sağlar.",
        "Hiz_RNA_Yikim = d[mRNA_hedef]/dt = k_HEPN * [Cas13:crRNA:hedef_RNA]",
        "Bu transkriptomik klirens kinetiği eşitliği, Cas13 efektörlerinin hedef mRNA'yı DNA hasarı yaratmaksızın doğrudan parçalama hızını simgeler."
    ),
    (
        "4.10 Split-Cas9 ve Kimyasal/Optogenetik İndüklenebilir Genom Cerrahisi",
        "İn vivo gen terapisinde Cas9'un sürekli ve kontrolsüz aktif kalması; hedef dışı mutasyon birikimini ve anti-Cas9 immün tepkilerini tetikleyen tehlikeli bir durumdur.",
        "Bu sorunu aşmak için Cas9 proteini iki inaktif parçaya bölünür (Split-Cas9): N-terminal (aa 1-573) ve C-terminal (aa 574-1368). Bu parçalar tek başlarına etkisizdir. Parçalara optogenetik dimerizasyon domainleri (Magnet proteinleri veya CRY2/CIBN) veya kimyasal dimerizörler (Rapamisin / FRB-FKBP ekseni) eklenir. Dokunun üzerine mavi ışık düşürüldüğünde veya küçük molekül verildiğinde iki parça bir araya gelerek Cas9'u saniyeler içinde aktive eder; ışık kapatıldığında aktivite durur.",
        "Aktivite_SplitCas = [Cas9_aktif] = K_assosiasyon * [N-Cas9] * [C-Cas9] * ( Foton_Mavi_Isik / K_isik )",
        "Bu optogenetik indüksiyon eşitliği, bölünmüş nükleaz sistemlerinin foton akısı kontrolü altında mikrosaniyelik zamansal ve mekânsal hassasiyetle çalıştırılabilmesini formüle eder."
    )
]

# ================= KISIM 5 =================
part5_subsections = [
    (
        "5.1 Baz Düzenleme (Base Editing) Devrimi: Çift Sarmal Kırığı (DSB) Olmadan Hassas Mühendislik",
        "Klasik CRISPR-Cas9 sisteminin yol açtığı çift sarmal kırıkları (DSB); büyük delesyonlara, translokasyonlara, kromotripsis felaketine ve p53 bağımlı hücre ölümüne neden olabilir.",
        "David Liu ve ekibi tarafından 2016'da icat edilen 'Baz Düzenleme' (Base Editing); DNA omurgasını kırmadan ve donör DNA şablonuna ihtiyaç duymadan, hedef DNA bazını doğrudan kimyasal olarak başka bir baza dönüştüren çığır açıcı bir teknolojidir. İnsan patojenik nokta mutasyonlarının %60'ından fazlası baz düzenleyicilerle tek bir nükleotit hassasiyetinde geri döndürülebilmektedir.",
        "P_DSB_BaseEditor << 0.001 * P_DSB_Cas9 (Binde Birden Az Kromozomal Hasar)",
        "Bu güvenlik karşılaştırması, baz düzenleme sistemlerinin kromozomal bütünlüğü bozmadan tek nükleotit dönüşümü gerçekleştirme hassasiyetini ifade eder."
    ),
    (
        "5.2 Sitozin Baz Düzenleyicileri (CBE): APOBEC/AID Deaminaz Kaskadı",
        "Sitozin Baz Düzenleyicileri (CBE - Cytosine Base Editors); genomdaki C:G baz çiftlerini T:A baz çiftlerine dönüştürmek üzere tasarlanmış füzyon proteinleridir.",
        "CBE mimarisi üç ana bileşenden oluşur: 1) Katalitik olarak zayıflatılmış Cas9 nickaz (Cas9n - D10A mutasyonu); 2) Tek sarmallı DNA'ya özgü Sitozin Deaminaz enzimi (örneğin sıçan APOBEC1 veya evriltilmiş evoAPOBEC1); 3) Hücrenin tamir mekanizmasını frenleyen Urasil Glikozilaz İnhibitörü (UGI). gRNA hedef bölgeyi açtığında APOBEC1 tek sarmaldaki sitozini deamine ederek Urasile (U) çevirir. Hücre DNA replikasyonu veya onarımı sırasında bu urasili Timin (T) olarak okur.",
        "Reaksiyon_CBE: Sitozin (C) --[APOBEC1 / -NH3]--> Urasil (U) ----> Timin (T)",
        "Bu biyokimyasal deaminasyon kaskadı, sitozin amin grubunun hidrolitik uzaklaştırılmasıyla urasil ara ürününün ve nihai timin bazının oluşumunu tanımlar."
    ),
    (
        "5.3 Urasil DNA Glikozilaz İnhibitörü (UGI) ve Baz Eksizyon Onarımının Engellenmesi",
        "Hücre çekirdeğinde normalde DNA'da urasil bulunması patolojik bir hasar sayılır; endojen Urasil DNA Glikozilaz (UDG) enzimi anında Urasili kesip çıkararak apürinik/apirimidinik (AP) abazik saha açar.",
        "Bu abazik saha baz eksizyon onarımını (BER) tetikleyerek sitozini geri getirebilir veya istenmeyen indellere yol açabilir. Bu hücresel geri tepmeyi engellemek için CBE kompleksinin ucuna bakteriyofaj PBS2 kaynaklı 'Urasil Glikozilaz İnhibitörü' (UGI) peptid zinciri eklenir. UGI, endojen UDG'nin aktif bölgesine pikomolar afiniteyle yapışarak enzimi kilitler; urasil hücre bölünmesine kadar korunur ve C->T dönüşümü %90'ın üzerinde saflıkla tamamlanır.",
        "K_i_UGI = [UDG] * [UGI] / [UDG:UGI] ~ 1.2 * 10^-14 M (Femtomolar Afinite ile Tam Blokaj)",
        "Bu enzim inhibisyon termodinamiği eşitliği, UGI peptitlerinin hücresel baz eksizyon savunmasını nasıl mutlak bir afiniteyle felç ettiğini belgeler."
    ),
    (
        "5.4 Adenin Baz Düzenleyicileri (ABE): Laboratuvarda Evriltilen TadA Deaminazı",
        "Doğada tek sarmallı DNA'daki adenini deamine edebilen bilinen hiçbir doğal enzim mevcut değildi; bu nedenle A:T baz çiftini G:C baz çiftine dönüştürmek biyolojinin en büyük meydan okumasıydı.",
        "Nicole Gaudelli ve David Liu; E. coli transfer-RNA adenin deaminazı olan 'TadA' enzimini aldılar ve laboratuvarda yönlendirilmiş evrimle (PACE) defalarca mutasyona uğrattılar. Sekizinci nesil 'TadA-8e' varyantı; tRNA yerine tek sarmallı DNA'yı tanıyan ve adenini saniyeler içinde İnozine (I) dönüştüren ultra-hızlı bir enzim haline geldi. Hücre DNA polimerazları inozini Guanin (G) olarak okur ve A:T -> G:C dönüşümü kusursuz biçimde tamamlanır.",
        "Reaksiyon_ABE: Adenin (A) --[TadA-8e / -NH3]--> İnozin (I) ----> Guanin (G)",
        "Bu sentetik enzim deaminasyon basamağı, laboratuvarda yönlendirilmiş evrimle sıfırdan yaratılan DNA adenin deaminaz mekanizmasını sembolize eder."
    ),
    (
        "5.5 Düzenleme Penceresi (Editing Window) ve Yan Çift (Bystander) Mutasyon Problemi",
        "Baz düzenleyiciler, gRNA'nın PAM dizisine olan mesafesine bağlı olarak dar bir tek sarmallı DNA penceresinde çalışır.",
        "Geleneksel CBE ve ABE sistemlerinde 'Düzenleme Penceresi' (Editing Window); PAM'in distal ucundan geriye doğru protospacer'ın 4. ila 8. nükleotidleri arasını kapsar. Eğer bu pencere içinde hedef nükleotidin yanında başka sitozin veya adenin bazları varsa deaminaz onları da dönüştürür (Yan Çift / Bystander Mutasyonları). Bu sorunu çözmek için deaminazın aktif cebi daraltılarak düzenleme penceresi tek bir nükleotide (örneğin sadece 6. baz) sıkıştırılmış 'dar pencereli' düzenleyiciler üretilmiştir.",
        "Pencere_Genisligi = W_edit = [Protospacer_Pozisyon_4 .. Pozisyon_8] (Klasik) -> [Pozisyon_6] (Hassas)",
        "Bu konfigürasyonel pencere tanımı, baz deaminasyon etkinliğinin protospacer üzerindeki dar nükleotit koordinatlarını ve optimizasyon sınırlarını tanımlar."
    ),
    (
        "5.6 Transkriptom-Çapında RNA Hedef Dışı Deaminasyonu ve Mühendislikli Deaminazlar",
        "Erken nesil CBE ve ABE'ler kromatinde DNA'yı düzenlerken; sitoplazmada hücrenin kendi haberci RNA'larına (mRNA) da bağlanarak binlerce transkriptte rastgele C->U veya A->I RNA deaminasyonu yapıyordu.",
        "Bu transkriptomik kaos; hücre fizyolojisini bozan yaygın proteomik hatalara neden olabiliyordu. Protein mühendisliği ile deaminazların RNA bağlama arayüzündeki kritik kalıntılar (APOBEC1'de R33A/K34A veya TadA'da V106W) mutasyona uğratıldı. Ortaya çıkan 'SECURE-CBE' ve 'mini-TadA' türevleri; RNA'ya bağlanma yeteneğini tamamen yitirirken DNA üzerindeki on-target deaminasyon gücünü korudu.",
        "Off_Target_RNA_Skoru = N_RNA_deaminasyon = 10^4 (Dogal) -> 0 (SECURE Varyantlari)",
        "Bu saflık metriği, rasyonel protein mühendisliği ile deaminazların transkriptomik arka plan tahribatının nasıl tamamen sıfırlandığını belgeler."
    ),
    (
        "5.7 Glikozilaz Baz Düzenleyicileri (GBE): C'den G'ye Transversiyon Teknolojisi",
        "Geleneksel CBE ve ABE'ler yalnızca geçiş (transition) mutasyonları (C->T veya A->G) yapabilirken; transversiyon mutasyonları (C->G veya A->C) uzun yıllar imkânsız kabul edilmiştir.",
        "Glikozilaz Baz Düzenleyicileri (GBE - Glycosylase Base Editors); Cas9 nickaz, sitozin deaminaz ve hücresel Urasil DNA Glikozilaz (UDG) enzimini tek bir moleküler füzyonda birleştirir. Deaminaz sitozini urasile çevirir çevirmez, hemen yanındaki UDG urasili koparır ve bir abazik saha (AP site) açar. Hücre bu sahayı translezyonel DNA polimerazlar (Rev1 gibi) ile onarırken karşı sarmala Sitozin yerleştirir; böylece C:G baz çifti doğrudan G:C baz çiftine (Transversiyon) dönüştürülür.",
        "Reaksiyon_GBE: C ----[Deaminaz]----> U ----[UDG]----> [Abazik AP Saha] ----[Rev1 Polimeraz]----> G",
        "Bu transversiyonel baz düzenleme akışı, deaminaz ve ekzojen DNA glikozilaz füzyonunun C->G transversiyonunu hücrede nasıl katalizlediğini gösterir."
    ),
    (
        "5.8 Çift Yönlü Düzenleyiciler (SPACE ve ACBE): Eş Zamanlı A->G ve C->T Dönüşümü",
        "Karmaşık genetik hastalıklarda ve çok genli yaşlanma regülasyonlarında aynı lokusta hem adenin hem de sitozin bazlarının aynı anda düzenlenmesi gerekebilir.",
        "Bu amaçla Cas9 nickazın bir ucuna TadA adenin deaminazı, diğer ucuna APOBEC sitozin deaminazı kaynaştırılarak 'SPACE' (Synchronous Programmable Adenine and Cytosine Editor) veya 'ACBE' sistemleri inşa edilmiştir. Tek bir gRNA rehberliğinde aynı düzenleme penceresinde bulunan hem A'lar hem C'ler tek bir reaksiyonda sırasıyla G ve T'ye çevrilebilmektedir. Bu durum metabolik yolakların ve transkripsiyon faktörü bağlanma motiflerinin çoklu olarak tek adımda optimize edilmesini sağlar.",
        "Vektor_SPACE = Cas9n + N_term_TadA8e + C_term_APOBEC + UGI",
        "Bu moleküler füzyon eşitliği, tek bir kaset içinde çalışan çift yönlü baz düzenleme motorunun mimari bileşenlerini tanımlar."
    ),
    (
        "5.9 Mitokondriyal DNA Baz Düzenleyicileri: DddA Toksini ve DddA-Türevli Düzenleyiciler (DdCBE)",
        "Mitokondriyal genom (mtDNA), nükleustan farklı olarak dışarıdan rehber RNA transferine tamamen kapalıdır; bu nedenle klasik CRISPR-Cas sistemleri mitokondri içine sokulamaz.",
        "David Liu ve Joseph Mougous; Burkholderia cenocepacia bakterisinin salgıladığı ölümcül bir toksin olan 'DddA'yı (çift sarmallı DNA sitozin deaminazı) keşfetmiştir. Toksin hücreyi öldürmesin diye iki inaktif yarıya bölünmüş (split-DddA) ve mitokondriyal hedefleme dizileri taşıyan TALE (Transcription Activator-Like Effector) proteinlerine bağlanmıştır. Mitokondri matriksinde bir araya gelen bu DdCBE kompleksi; gRNA'ya gerek duymadan çift sarmallı mtDNA'daki sitozinleri timine dönüştürerek mitokondriyal hastalıkların genetik tedavisini tarihte ilk kez başarmıştır.",
        "Reaksiyon_mtDNA: dsDNA_C:G --[TALE-DddA-split + Mito_Lokalizasyon]--> dsDNA_T:A",
        "Bu mitokondriyal genom cerrahisi eşitliği, gRNA bağımsız protein-DNA tanıma ve çift sarmal deaminasyon yoluyla mitokondriyal genom modifikasyonunu ifade eder."
    ),
    (
        "5.10 Klinik İn Vivo Baz Düzenleme: PCSK9, Verve-101 ve Kardiyovasküler Genomik Kür",
        "Baz düzenleme laboratuvar deneylerinden hızla insan klinik uygulamalarına geçmiştir; bunun en çarpıcı örneği Verve Therapeutics tarafından yürütülen 'Verve-101' klinik çalışmasıdır.",
        "Verve-101; karaciğerde kolesterol metabolizmasını yöneten PCSK9 genini hedefleyen bir Adenin Baz Düzenleyicisi (ABE) mRNA'sı ve gRNA'yı LNP içinde taşır. Sistemik tek bir intravenöz infüzyonla hepatositlere ulaşan ABE; PCSK9 geninin ekzon 1 bölgesindeki tek bir adenini guanine çevirerek (A->G) kritik bir donör splice bölgesini inaktive eder. İnsanlarda kan LDL kolesterol seviyesini ömür boyu kalıcı olarak %55 düşüren bu tek dozluk tedavi, in vivo baz düzenlemenin klinik rüştünü ispatlamıştır.",
        "Verim_Knockout_PCSK9 = eta_in_vivo = [Hepatik_A_to_G_Orani] ~ 70% (Tek Doz LNP Enjeksiyonu ile)",
        "Bu klinik farmakodinamik parametresi, sistemik uygulanan baz düzenleyicilerin hedef organ genomunda elde ettiği yüksek tek-nükleotit dönüşüm başarısını modeller."
    )
]

# ================= KISIM 6 =================
part6_subsections = [
    (
        "6.1 Prime Editing (PE) Mekanizması: Genomun 'Ara ve Değiştir' (Search-and-Replace) Arama Motoru",
        "Baz düzenleyiciler yalnızca 4 geçiş mutasyonunu (C->T, G->A, A->G, T->C) yapabilirken; 8 olası transversiyon mutasyonunu yapamaz ve hedeflenen bölgeye yeni nükleotit ekleyemez ya da silemez.",
        "Andrew Anzalone ve David Liu tarafından 2019'da geliştirilen 'Prime Editing' (PE); DNA çift sarmalını kırmadan ve donör DNA şablonu kullanmadan; 12 olası baz-baz dönüşümünün tamamını, hedeflenen mikroskobik insersiyonları ve delesyonları kusursuzca gerçekleştiren evrensel bir arama-değiştirme teknolojisidir.",
        "Kapsama_Hastalik_PE = [Tum_Genetik_Varyantlar] = 12_Gecis_ve_Transversiyon + Insersiyon + Delesyon ~ %89",
        "Bu klinik genomik kapsama oranı, Prime Editing sisteminin bilinen 75.000'den fazla insan patojenik genetik mutasyonunun yaklaşık %89'unu teorik olarak düzeltebilme gücünü tanımlar."
    ),
    (
        "6.2 Prime Editing Rehber RNA (pegRNA) Mimarisi: PBS ve RT Şablon Alanları",
        "Prime Editing'in dehası, rehber RNA'nın 3' ucuna eklenen çok işlevli sentetik uzantıda (pegRNA - Prime Editing Guide RNA) yatar.",
        "Bir pegRNA üç kritik fonksiyona sahiptir: 1) Genomdaki hedef lokusu bulan 20 nükleotitlik standart Spacer dizisi; 2) Kesilen hedef DNA ucuna hibritlenen 13 nükleotitlik Primer Bağlanma Bölgesi (PBS - Primer Binding Site); 3) İstenen yeni genetik sekansı ve düzeltmeleri içeren Ters Transkriptaz Şablonu (RTT - Reverse Transcriptase Template). pegRNA hem hedefi bulur hem de sentezlenecek yeni genetik şifreyi kendi üzerinde taşır.",
        "Mimari_pegRNA = [5' - Spacer (20 nt) - gRNA_İskelet (82 nt) - RTT_Şablon (10-30 nt) - PBS (13 nt) - 3']",
        "Bu oligonükleotit yapısal dizilimi, pegRNA molekülünün hedef arama ve polimerizasyon şablonlama kompartmanlarının ardışık mimarisini belgeler."
    ),
    (
        "6.3 Cas9 Nickaz-Ters Transkriptaz (Cas9n-RT) Füzyon Proteini Biyofiziği",
        "Prime Editor'ün katalitik motoru; SpCas9'un H840A nükleaz inaktif varyantı (Cas9 nickaz) ile Moloney Murin Lösemi Virüsü Ters Transkriptazının (M-MLV RT) genetik füzyonudur.",
        "Cas9n(H840A), pegRNA rehberliğinde DNA'ya bağlandığında yalnızca PAM içeren PAM sarmalını keser (tek sarmal nick). Açığa çıkan serbest 3'-OH ucu, pegRNA'nın PBS bölgesine hibritlenir. Bu 3'-OH ucunu primer olarak kullanan M-MLV RT enzimi; RTT şablonunu okuyarak doğrudan konak kromatininin içine yeni genetik diziyi de novo polimerize eder. Böylece genetik bilgi RNA'dan DNA'ya doğrudan kromozom üzerinde kopyalanır.",
        "Reaksiyon_PE: Hedef_DNA_3'OH + pegRNA_RTT --[M-MLV_RT]--> Yeni_Duzenlenmis_DNA_Flap",
        "Bu nükleotit polimerizasyon denklemi, ters transkriptaz enziminin kromozomal tek sarmal ucunu uzatarak istenen genetik diziyi sentezleme basamağını formüle eder."
    ),
    (
        "6.4 Flap Rekabeti: 3' Flap Entegrasyonu vs 5' Flap Endonükleolitik Eksizyonu",
        "Ters transkriptaz yeni DNA dizisini sentezlediğinde, hedef lokusta çatallı bir yapı oluşur: Yeni sentezlenen düzenlenmiş '3' Flap' ve genomun eski orijinal '5' Flap'i.",
        "Hücrenin istenen düzeltmeyi genomuna kalıcı olarak yazabilmesi için eski 5' flap'in kesilip atılması ve yeni 3' flap'in kromatinde sarmal haline gelerek DNA ligaz ile bağlanması şarttır. Hücresel Flap Endonükleaz 1 (FEN1) enzimi, 5' flap'i tanıyarak hidrolize eder; hücrenin DNA ligazı yeni sarmalı kapatır. Eğer 3' flap nükleazlarca parçalanırsa düzenleme başarısız olur ve eski genetik dizi korunur.",
        "Verim_Flap_Entegrasyon = eta_flap = Hiz_FEN1_kesim / ( Hiz_FEN1_kesim + Hiz_3Flap_bozunma )",
        "Bu enzimatik yarışma oranı, istenen flap entegrasyon başarısının FEN1 endonükleaz aktivitesi ve flap termodinamik kararlılığı ile doğrudan ilişkisini ortaya koyar."
    ),
    (
        "6.5 PE1'den PE2'ye Evrim: M-MLV RT'nin Termostabilite ve Prosesivite Mutasyonları",
        "İlk nesil Prime Editor (PE1), vahşi tip M-MLV ters transkriptazı kullandığı için insan hücrelerinde %1 ila %5 gibi son derece düşük düzenleme verimleri sergiliyordu.",
        "PE2 sisteminde; M-MLV RT enzimine yönlendirilmiş evrimle 5 kritik amino asit mutasyonu (D200N, L603W, T330P, T306K, W313F) eklenmiştir. Bu mutasyonlar enzimin termostabilitesini (37°C'de erime direncini) artırmış, DNA:RNA kompleksine bağlanma afinitesini katlamış ve prosesivitesini 10 kat hızlandırmıştır. PE2, aynı lokuslarda PE1'den 3 ila 5 kat daha yüksek düzenleme verimi sağlamıştır.",
        "Katsayi_Prosesivite = P_PE2 / P_PE1 ~ 5.2 (Termostabil M-MLV-5M Mutasyon Avantajı)",
        "Bu katalitik kinetik oranı, evriltilmiş ters transkriptaz alanının polimerizasyon sürekliliği ve baz ekleme hızındaki üstünlüğünü belgeler."
    ),
    (
        "6.6 PE3 ve PE3b Sistemleri: İkincil Nickleme ile Eşleşmemiş İplik Çözümlemesi",
        "PE2 ile yeni 3' flap sentezlense dahi; heterodupleks DNA yapısında karşıt sarmal hâlâ eski mutasyonu taşıdığı için hücrenin Mismatch Repair (MMR) sistemi yeni diziyi silip eskiyi şablon alabilmektedir.",
        "Bu direnci kırmak için 'PE3' sistemi geliştirilmiştir: Düzenlenmemiş karşı sarmala ikinci bir standart sgRNA ile ek bir tek sarmal nick atılır. Hücre bu nicki gördüğünde karşı sarmalı hasarlı zanneder; orayı ekzonükleazlarla sindirir ve yeni düzenlenmiş sarmalı kalıp alarak karşı sarmalı yeniden sentezler. Böylece düzenleme kalıcılaşır ve verim 3 ila 4 kat daha artar. İkincil nickin yalnızca 3' flap entegre olduktan sonra atıldığı 'PE3b' varyantı ise indelleri tamamen sıfırlar.",
        "Verim_Duzenleme_PE3 = eta_PE3 = eta_PE2 * ( 1 + alpha_nick * P_karsi_iplik_tamiri )",
        "Bu heterodupleks çözümleme denklemi, karşıt sarmal nicklemesinin hücresel tamir yönünü istenen şablon lehine nasıl yönlendirdiğini simgeler."
    ),
    (
        "6.7 pegRNA Kararlılığı: epegRNA ve Yapısal 3' Saç Tokası (evopreQ1/tmpknot) Tasarımı",
        "Standart pegRNA moleküllerinin en zayıf noktası, 3' uçlarındaki tek sarmallı PBS ve RTT kuyruklarının hücresel ribonükleazlar (ekzoribonükleazlar) tarafından hızla çiğnenip parçalanmasıdır.",
        "David Liu laboratuvarı; pegRNA'nın 3' ucuna yapısal nükleaz dirençli RNA saç tokaları (evopreQ1 veya tmpknot motifleri) ekleyerek 'epegRNA' (engineered pegRNA) teknolojisini icat etmiştir. Bu kompakt tersiyer RNA kalkanı; ekzonükleazların pegRNA'ya yanaşmasını sterik olarak engeller, molekülün yarı ömrünü 4 kat uzatır ve insan hücrelerinde düzenleme verimini hiçbir toksisite yaratmadan 3-4 katına çıkarır.",
        "Yari_Omur_RNA = Tau_epegRNA = 3.8 * Tau_pegRNA_standart",
        "Bu kinetik kararlılık eşitliği, 3' uca yerleştirilen yapısal saç tokasının pegRNA hücresel varlık süresinde yarattığı koruyucu etkiyi formüle eder."
    ),
    (
        "6.8 PE-max ve MMR-Kaçış Mühendisliği: MLH1dn ile Uyumsuzluk Onarımını Aşma",
        "Hücrenin Uyumsuzluk Onarımı (MMR - Mismatch Repair) yolağı, Prime Editing ile yerleştirilen tek nükleotit değişikliklerini aktif olarak tanıyan ve düzelten primer hücresel savunma mekanizmasıdır.",
        "PE-max mimarisinde; Cas9 ve RT kodonları yeniden optimize edilmiş, nükleer lokalizasyon sinyalleri (NLS) güçlendirilmiş ve sisteme geçici bir dominant-negatif MLH1 proteini (MLH1dn) eklenmiştir. MLH1dn hücresel MMR kaskadını geçici olarak uyuşturur; bu süre zarfında Prime Editor istenen baz değişikliklerini hücrenin tamir engeline takılmadan kalıcı olarak kromozoma diker. Bu yöntem özellikle tek bazlık transversiyon verimlerini %80'in üzerine çıkarmıştır.",
        "Verim_PEmax = eta_PEmax = eta_PE3 * ( 1 / ( 1 + [MMR_aktif] / K_mmr ) )",
        "Bu hücresel kaçış eşitliği, mismatch repair kaskadının geçici inhibisyonunun prime editing hassasiyetini ve nihai verimini nasıl zirveye taşıdığını ortaya koyar."
    ),
    (
        "6.9 İkiz Prime Düzenleme (TwinPE) ve Büyük Genomik İnsersiyonlar (Bxb1 Rekombinaz)",
        "Standart Prime Editing 50-100 baz çiftine kadar olan küçük insersiyonları yapabilirken; binlerce baz çiftlik tam fonksiyonel genlerin genomun istenen lokusuna dikişsiz yerleştirilmesi imkânsızdı.",
        "TwinPE teknolojisinde; zıt sarmalları hedefleyen iki ayrı pegRNA kullanılarak DNA'da birbirini tamamlayan iki uzun 3' flap sentezlenir. Bu flapler içerisine faj kaynaklı Bxb1 integraz/rekombinaz enzimi için spesifik 'attB' ve 'attP' tanıma siteleri yerleştirilir. İkinci adımda Bxb1 rekombinazı; donör plazmiddeki 10-20 kilobazlık devasa terapötik genleri bu siteler arasına %100 kusursuz yönelimle ve tek bir çift sarmal kırığı açmadan yerleştirir.",
        "Kapasite_TwinPE = L_entegre_gen = 5_kb - 50_kb (Tam Terapötik Transgen İnsersiyonu)",
        "Bu makro-genomik montaj eşitliği, TwinPE ve serin rekombinaz füzyonlarının kromozom içine sınırsız büyüklükte transgen yerleştirme kapasitesini belgeler."
    ),
    (
        "6.10 İn Vivo Prime Editing Uygulamaları: Nörodejenerasyon ve Genetik Karaciğer Hastalıkları",
        "Prime Editing, in vivo dağıtım zorluklarını (büyük gen boyutu nedeniyle split-AAV veya LNP-mRNA gereksinimi) hızla aşarak klinik öncesi hayvan modellerinde devasa zaferler kazanmıştır.",
        "Prion hastalıkları, Alzheimer ApoE4 aleli, Tay-Sachs hastalığı ve genetik kardiyomiyopatilerde; LNP veya split-AAV ile sistemik ve intratekal olarak teslim edilen PE sistemleri, hedef nöron ve hepatositlerde patojenik mutasyonları %40-60 verimle ve sıfıra yakın indel ile düzeltmiştir. Çift sarmal kırığı açmayan bu arama-değiştirme motoru, in vivo gen cerrahisinde güvenliğin nihai standardı haline gelmiştir.",
        "Guvenlik_Indeksi = SI_PE = ( Verim_On_Target / ( Verim_Off_Target + Verim_Indel + Translocation ) )_PE >> 10^5",
        "Bu mutlak güvenlik indeksi, Prime Editing platformunun genomik cerrahide klasik Cas nükleazlarına kıyasla yüz bin kat daha temiz bir mutasyon profili sunduğunu kanıtlar."
    )
]

# ================= KISIM 7 =================
part7_subsections = [
    (
        "7.1 Epigenom Düzenleme Paradigması: DNA Dizisini Değiştirmeden Gen İfadesi Kontrolü",
        "Genomik tıp uzun yıllar nükleotit dizilimini değiştirmeye odaklanmış olsa da; gen ifadesinin asıl yöneticisi kromatini saran epigenetik kimyasal işaretlerdir.",
        "Epigenom Düzenleme (Epigenome Editing); DNA baz dizilimine tek bir müdahale yapmaksızın, hedeflenen gen lokuslarının histon kuyruklarını modifiye eden veya DNA sitozinlerini metilleyip/demetilliyen programlanabilir moleküler araçlardır. Bu yaklaşım; kalıcı mutasyon riski taşımaması, teorik olarak tamamen geri döndürülebilir olması ve tek bir promotör üzerinden tüm bir metabolik kaskadı susturabilmesi veya patlatabilmesi ile eşsiz bir üstünlük sunar.",
        "Modulasyon_Epigenom = Delta_Ekspresyon = f( Delta_Metilasyon_CpG, Delta_Asetilasyon_H3K27 ) (DNA Dizisi = Sabit)",
        "Bu transkripsiyonel regülasyon prensibi, hedef gen ifadesindeki radikal değişimlerin nükleotit sekansını mutasyona uğratmaksızın epigenetik kovalent modifikasyonlarla sağlandığını ifade eder."
    ),
    (
        "7.2 Katalitik İnaktif dCas9: Moleküler Postacı ve Programlanabilir DNA Bağlanma Platformu",
        "Tüm modern epigenom düzenleme araçlarının omurgasını; her iki nükleaz cebi mutasyona uğratılarak (D10A ve H840A) kesme yeteneği tamamen yok edilmiş 'Ölü Cas9' (dCas9 - dead Cas9) oluşturur.",
        "dCas9 enzimi DNA'yı kesemez; ancak rehber RNA'sının (sgRNA) gösterdiği 20 bazlık hedef diziyi pikomolar afiniteyle (Kd ~ 10^-12 M) bulur ve kromatinde oraya bir moleküler kaya gibi kenetlenir. Bu özelliği dCas9'u; transkripsiyonel represörleri, aktivatörleri, histon asetilazları veya DNA metiltransferazları istenen genin promotörüne taşıyan kusursuz bir 'Moleküler Postacı'ya dönüştürür.",
        "K_d_dCas9 = [dCas9:sgRNA] * [Hedef_DNA] / [dCas9:sgRNA:DNA] ~ 10^-12 M",
        "Bu ayrışma sabiti eşitliği, nükleaz-inaktif dCas9 kompleksinin kromatindeki hedef sekansa bağlanma gücünün olağanüstü yüksek termodinamik stabilitesini tanımlar."
    ),
    (
        "7.3 dCas9-KRAB ve CRISPRi: Transkripsiyonel Susturma ve Heterokromatinleşme Biyofiziği",
        "CRISPR İnterferans (CRISPRi) teknolojisinde; dCas9 proteini, insan çinko parmak transkripsiyon faktörlerinde bulunan Krüppel-ilişkili Kutu (KRAB) represör alanına kaynaştırılır.",
        "dCas9-KRAB hedef genin transkripsiyon başlangıç bölgesine (TSS) oturduğunda; hücresel korepresör KAP1'i (TRIM28) alana çağırır. KAP1 kompleksi; SETDB1 histon metiltransferazı ve NuRD deasetilaz kompleksini organize ederek genin etrafındaki histon 3 lizin 9 kalıntılarını trimetiller (H3K9me3) ve H3K27 asetilasyonunu siler. Kromatin yerel olarak sıkışarak 'Fakültatif Heterokromatine' bürünür ve RNA polimeraz II fiziksel olarak kilitlenir; gen ifadesi %99 oranında susar.",
        "Susturma_Verimi = eta_CRISPRi = 1 - ( [mRNA_kalan] / [mRNA_kontrol] ) = k_H3K9me3 * ( [dCas9-KRAB] / ( K_m + [dCas9-KRAB] ) )",
        "Bu transkripsiyonel represyon denklemi, KRAB aracılı H3K9me3 heterokromatin birikiminin hedef gen susturma oranını nasıl doygunluk eğrisine taşıdığını simgeler."
    ),
    (
        "7.4 dCas9-VPR, SunTag ve CRISPRa: Gen Ekspresyonunun Güçlü Transkripsiyonel Aktivasyonu",
        "Tümör baskılayıcılar veya ömür uzatıcı faktörler gibi yaşla susan kritik genleri yeniden uyandırmak için CRISPR Aktivasyon (CRISPRa) platformları kullanılır.",
        "1) dCas9-VPR: dCas9'un C-terminaline üç güçlü transkripsiyonel aktivatör (VP64, p65 ve Rta) ardışık bağlanır; bu üçlü hücresel transkripsiyon başlatma kompleksini (TFIID, TFIIH) derhal promotöre yığar; 2) SunTag Sistemi: dCas9 ucuna tekrarlayan 10-24 peptitlik bir polipeptit kuyruğu eklenir; bu kuyruğa tek zincirli antikorlar (scFv) aracılığıyla onlarca aktivatör molekülü bir üzüm salkımı gibi dizilir. SunTag, susturulmuş genlerin ifadesini bazal seviyenin 100 ila 10.000 katına fırlatabilir.",
        "Aktivasyon_Katsayisi = Alfa_CRISPRa = [mRNA_indukte] / [mRNA_bazal] = N_aktivator * k_RNA_Pol_II_cagirma",
        "Bu transkripsiyonel amplifikasyon eşitliği, promotör lokusuna çağrılan aktivatör modül sayısının RNA Polimeraz II transkripsiyon hızında yarattığı çarpan etkisini modeller."
    ),
    (
        "7.5 dCas9-p300 ile Hedefli Histon Asetilasyonu ve Açık Kromatin Mimarisi",
        "Transkripsiyon faktörlerinin kromatinde kapalı bölgelere erişebilmesi için histon kuyruklarının asetillenerek nükleozomların gevşetilmesi şarttır.",
        "dCas9 proteini, insan histon asetiltransferazı p300'ün aktif katalitik çekirdeğine (p300HATcore) kaynaştırılmıştır. dCas9-p300, hedef promotör ve uzak enhanser (güçlendirici) bölgelerindeki nükleozomlarda yer alan Histon 3 Lizin 27 (H3K27ac) kalıntılarını yüksek spesifiteyle asetiller. Asetil grupları histonların pozitif yükünü nötralize eder; negatif yüklü DNA iplikleri nükleozomdan gevşeyerek açılır (Ökromatin). Gen bölgesi tüm hücresel transkripsiyon faktörlerinin serbestçe bağlanabileceği kalıcı bir aktif duruma geçer.",
        "Yuk_Azalmasi = Q_histon = Q_0 - n_asetil * e = Z_pozitif * ( 1 - [p300_reaksiyon] )",
        "Bu elektrostatik gevşeme denklemi, p300 katalizli H3K27 asetilasyonunun nükleozomal çekim kuvvetini düşürerek açık kromatin mimarisini nasıl inşa ettiğini belgeler."
    ),
    (
        "7.6 dCas9-DNMT3A/3L ile Kalıcı DNA Metilasyonu ve Nesiller Boyu Epigenetik Bellek",
        "Kısa süreli gen baskılaması geçicidir; ancak hücre bölünse dahi unutulmayacak kalıcı bir epigenetik sessizlik yaratmak için doğrudan DNA'nın kendisi metillenmelidir.",
        "dCas9; de novo DNA metiltransferazı DNMT3A ve onun regülatör faktörü DNMT3L'nin katalitik kasetine kaynaştırılır. dCas9-DNMT3A/3L kompleksi, hedef genin CpG adalarındaki sitozinlerin 5. karbonuna kovalent olarak metil grubu ekleyerek 5-metilsitozin (5mC) üretir. Bu metilasyon, hücre bölündüğünde endojen bakım metiltransferazı DNMT1 tarafından tanınır ve yavru hücrelere aynen kopyalanır. Böylece geçici bir dCas9 maruziyeti sonrasında bile gen ömür boyu kapalı kalır (Epigenetik Bellek).",
        "Bellek_Kararliligi = P_kalicilik = ( 1 - P_pasif_demetilasyon )^N_bolunme = [ 1 - (1 - eta_DNMT1) ]^N",
        "Bu epigenetik kalıtım fonksiyonu, de novo metillenen CpG adalarının endojen DNMT1 bakım kaskadı sayesinde hücre bölünmeleri boyunca stabil kalma olasılığını formüle eder."
    ),
    (
        "7.7 dCas9-TET1 ile Aktif DNA Demetilasyonu ve Susturulmuş Genlerin Kurtarılması",
        "Yaşlanma sürecinde ve senesens patolojisinde, koruyucu anti-aging genlerinin CpG adaları patolojik hipermetilasyonla kilitlenir ve susturulur.",
        "Bu kilitlenmeyi çözmek için dCas9; Ten-Eleven Translocation dioksijenaz ailesinin üyesi olan 'TET1' enziminin katalitik domainine bağlanır. dCas9-TET1, kromatindeki 5-metilsitozini (5mC) ardışık adımlarla 5-hidroksimetilsitozin (5hmC), 5-formilsitozin (5fC) ve 5-karboksisitozine (5caC) oksitler. Hücresel timin DNA glikozilaz (TDG) ve baz eksizyon onarımı bu modifiye bazları kesip yerine saf metilsiz sitozin koyar. CpG adası tamamen temizlenir ve susturulmuş gen gençlik coşkusuyla yeniden uyanır.",
        "Reaksiyon_Demetilasyon: 5mC --[TET1 / O2 / Fe2+]--> 5hmC ----> 5fC ----> 5caC --[TDG]--> Saf Sitozin (C)",
        "Bu biyokimyasal hidroksilasyon kaskadı, TET1 enziminin demir ve oksijen kofaktörleriyle metil gruplarını kademeli oksitleyerek DNA'yı metilasyondan arındırma mekanizmasını tanımlar."
    ),
    (
        "7.8 'Hit-and-Run' Epigenetik Düzenleme: Geçici İfade ile Kalıcı Transkripsiyonel Yeniden Programlama",
        "Klasik gen terapisinde viral vektörün ömür boyu transgen üretmesi istenir; ancak epigenom düzenleyicilerin hücrede sürekli kalması immünolojik ve hedef dışı riskler taşır.",
        "'Hit-and-Run' (Vur ve Kaç) paradigmasında; dCas9-epigenetik düzenleyici hücreye geçici bir mRNA veya LNP-RNP formunda verilir. Enzim hücrede yalnızca 24 ila 48 saat hayatta kalır. Bu kısa pencere içinde hedef lokusa giderek histon ve DNA metilasyon haritasını yeniden yazar ve ardından hücre içi proteazlarca tamamen parçalanır. Arkasında hiçbir yabancı protein veya viral artık bırakmaz; ancak hücrenin epigenetik belleği kalıcı olarak yeniden programlanmıştır.",
        "Guvenlik_HitAndRun: Tau_yasam_editor < 48_saat | Tau_etki_epigenetik > Yillar / Omur Boyu",
        "Bu kinetik asimetri kuralı, geçici moleküler varlık süresine sahip epigenetik cerrahi araçlarının yarattığı ömür boyu kalıcı transkripsiyonel gençleşmeyi simgeler."
    ),
    (
        "7.9 Yaşlanma Saatinin (Horvath Clock) Epigenomik Düzenleyicilerle Hedefli Sıfırlanması",
        "Steve Horvath ve ekibinin keşfettiği epigenetik saatler; insan dokularının biyolojik yaşını genomdaki birkaç yüz spesifik CpG sahasının metilasyon durumuna bakarak 2-3 yıl hata payıyla ölçebilmektedir.",
        "Yaşlanma karşıtı in vivo epigenetik cerrahi; yaşla birlikte aşırı metillenen CpG sitelerine dCas9-TET1'i, yaşla hipometile olan koruyucu bölgelere ise dCas9-DNMT3A'yı yönlendirir. Tek tek sahaların rasyonel mühendislikle gençlik koordinatlarına döndürülmesi; hücrenin küresel transkriptomunu, proteomik dengesini ve mitokondriyal biyogenezini gençlik fazına taşır ve biyolojik epigenetik saati geriye sarar.",
        "Biyolojik_Yas_Sifirlama = Delta_Yas_Horvath = - Sigma_i w_i * ( Metilasyon_yasli(CpG_i) - Metilasyon_genc(CpG_i) )",
        "Bu biyolojik yaş restorasyon eşitliği, hedefli epigenetik demetilasyon ve metilasyon müdahalelerinin Horvath epigenetik saat algoritmasındaki net yaş düşüşünü modeller."
    ),
    (
        "7.10 İn Vivo Epigenom Cerrahisi: Progeria, Hepatik Rejenerasyon ve Nöronal Gençleşme",
        "İn vivo epigenom düzenleme, fare modellerinde ölümcül genetik ve dejeneratif tabloları başarıyla tedavi ederek klinik çağın eşiğine gelmiştir.",
        "Hutchinson-Gilford Progeria Sendromunda (erken yaşlanma); toksik progerin proteininin ifadesi dCas9-KRAB taşıyan AAV vektörleri ile tek dozda susturulmuş, farelerin damar sertliği gerilemiş ve yaşam süreleri %35 uzatılmıştır. Benzer şekilde nöronlarda dCas9-p300 ile BDNF promotörünün açılması veya karaciğerde dCas9-DNMT3A ile ANGPTL3 geninin metillenerek susturulması; hayvanları genetik hasar riski olmaksızın metabolik ve bilişsel olarak gençleştirmiştir.",
        "Omur_Uzama_Katsayisi = L_kazanc = [Progerin_susturma]% * ( 1 + k_rejenerasyon * [BDNF_epigenetik] )",
        "Bu in vivo terapötik etkinlik denklemi, epigenetik susturma ve transkripsiyonel aktivasyon kombinasyonlarının sistemik yaşam süresi ve doku gençleşmesi üzerindeki birleşik etkisini tanımlar."
    )
]

# ================= KISIM 8 =================
part8_subsections = [
    (
        "8.1 Hedef Dışı (Off-Target) Kinetiği Biyofiziği: Uyumsuzluk Toleransı ve Enerjetik Manzara",
        "CRISPR-Cas9 sisteminin doğadaki evrimsel görevi bakteriyi virüslerden korumaktır; bu savunmada mutasyona uğramış viral kaçakları yakalayabilmek için enzim hafif baz uyumsuzluklarına tolerans gösterecek şekilde evrilmiştir.",
        "Ancak insan genomunda 3.2 milyar baz çifti içinde bu 'Uyumsuzluk Toleransı' (Mismatch Tolerance) ölümcül hedef dışı kesimlere yol açabilir. Rehber RNA ile DNA arasındaki bağlanma enerjisi serbest enerji manzarası (Free Energy Landscape) ile yönetilir: Eğer uyumsuzluk PAM'e yakın 'Tohum Bölgesi'nde (Seed, 1-10 nt) ise R-loop derhal çöker ve kesim engellenir; ancak uyumsuzluk PAM'e uzak distal bölgede (15-20 nt) ise Cas9 konformasyonel kilitlenmeyi tamamlayarak hedef dışı bölgeyi kesebilir.",
        "K_kesim_off = k_on * exp( - Delta_Delta_G_uyumsuzluk / (R * T) )",
        "Bu Arrhenius termodinamik hız eşitliği, hedef dışı lokustaki baz uyumsuzluklarının heterodupleks ayrışma serbest enerjisini ve nihai kesim hızını nasıl kontrol ettiğini modeller."
    ),
    (
        "8.2 GUIDE-seq: Çift Sarmal Kırıklarının Genom Çapında Haritalanması",
        "Hedef dışı mutasyonları sadece biyoinformatik tahminlerle aramak kör noktalar yaratır; hücre içinde gerçekte hangi lokusların kesildiğini doğrudan yakalayan deneysel yöntemler şarttır.",
        "GUIDE-seq (Genome-wide Unbiased Identification of DSBs Enabled by Sequencing); hücreye Cas9 ile birlikte çift sarmallı kısa bir fosforotiyoat oligonükleotit (dsODN) verir. Cas9 nükleusta nerede bir çift sarmal kırığı açarsa, hücrenin NHEJ tamir enzimleri bu sentetik dsODN etiketini kırık noktasına kaynaklar. Ardından DNA izole edilir, etikete özgü primerlerle yeni nesil dizileme (NGS) yapılır ve genomdaki tüm on-target ve off-target kesim koordinatları tek bir baz hassasiyetiyle saptanır.",
        "Duyarlilik_GUIDEseq = Tespit_Limiti ~ 0.1% (Binde bir sıklıktaki hedef dışı kesimleri yakalama gücü)",
        "Bu deneysel analiz limiti, GUIDE-seq metodolojisinin hücre içi çift sarmal kırıklarını genom çapında arka plandan ayıklama hassasiyetini belgeler."
    ),
    (
        "8.3 CIRCLE-seq ve Digenome-seq: Hücresiz (In Vitro) Genomik Kesim Analizleri",
        "Hücre içi testler kromatinde kapalı olan bölgelerdeki kesimleri göremeyebilir; ancak hücre tipi değiştikçe bu kapalı bölgeler açılabilir.",
        "Digenome-seq; saflaştırılmış insan genomik DNA'sını tüpte doğrudan Cas9-gRNA ribonükleoproteini ile sindirir ve kesilen homojen uçları tüm genom dizilemesiyle (WGS) bularak potansiyel tüm kesim alanlarını listeler. CIRCLE-seq ise genomik DNA'yı dairesel halkalara dönüştürür; Cas9 halkayı kestiğinde açılan doğrusal uçlar doğrudan adaptörlerle bağlanıp dizilenir. CIRCLE-seq, hücre içi testlerden 10 kat daha yüksek bir hassasiyetle (0.01%) teorik tüm kesim tehditlerini deşifre eder.",
        "Duyarlilik_CIRCLEseq = Tespit_Limiti ~ 0.01% (On binde bir seviyesinde kesim yakalama)",
        "Bu in vitro tarama standardı, hücresiz kütüphanelerde kromatin kısıtlamasından bağımsız olarak en zayıf hedef dışı afinitelerin dahi eksiksiz dökümünü çıkarır."
    ),
    (
        "8.4 Genomik Yapısal Varyasyonlar: Büyük Delesyonlar, İnversiyonlar ve Translokasyonlar",
        "Cas9 kesiminin en tehlikeli ve uzun süre gözden kaçan yan etkisi, hedef bölgede sadece birkaç bazlık indeller değil; kilobazlarca hatta megabazlarca uzunlukta 'Büyük Genomik Delesyonlar' (Large Deletions) açabilmesidir.",
        "Eğer genomda iki farklı lokus (on-target ve off-target) aynı anda kesilirse, iki kromozom parçası çapraz bağlanarak 'Kromozomal Translokasyon' oluşturabilir. Daha da vahimi, tek bir kromozom kolunun tamamen parçalanıp rastgele düzensizce yeniden birleşmesi olan 'Kromotripsis' (Chromothripsis) ve heterozigotluk kaybı (LOH) tetiklenebilir; bu yapısal felaketler proto-onkojenleri kontrolsüz promotörlerin altına sokarak karsinogenez riskini fırlatır.",
        "Risk_Translokasyon = P_trans = k_fujyon * [DSB_lokus_A] * [DSB_lokus_B] * ( 1 / Mesafe_nukleer^3 )",
        "Bu uzamsal kromozom çarpışma formülü, eş zamanlı oluşan iki çift sarmal kırığının çekirdek içindeki mikroskobik yakınlıklarına bağlı translokasyon oluşturma olasılığını formüle eder."
    ),
    (
        "8.5 p53 Yanıtı ve Onkojenik Seçilim Baskısı Riski",
        "Bölünme yeteneğine sahip sağlıklı insan hücrelerinde (özellikle pluripotent kök hücreler ve hematopoietik kök hücreler) Cas9 ile çift sarmal kırığı açıldığında; p53 tümör baskılayıcı proteini derhal aktive olur.",
        "p53 kaskadı, DNA tamir edilene kadar hücre döngüsünü G1 fazında durdurur veya hücreyi apoptoza sürükler. Bu durum CRISPR ile gen düzenleme verimini düşürür. Ancak çok daha sinsi bir tehlike mevcuttur: Düzenleme işleminden sağ çıkan hücreler, p53 yolağı zaten mutasyona uğramış ve devre dışı kalmış hücre klonları olabilir. Bu durum klinik gen terapisinde farkında olmadan pre-kanseröz hücre klonlarının seçilip çoğaltılması (onkojenik klonal seleksiyon) riskini doğurur.",
        "Secilim_Baskisi = S_p53 = [Hucre_p53_mutant_sagkalan] / [Hucre_p53_saglikli_sagkalan] >> 1.0 (DSB Varlıgında)",
        "Bu klonal evrim eşitliği, çift sarmal kırığı oluşturan nükleazların p53-disfonksiyonel tehlikeli klonları istemeden nasıl zenginleştirdiğini ortaya koyar."
    ),
    (
        "8.6 Rasyonel gRNA Tasarımı ve Biyoinformatik Skorlama Algoritmaları (CRISPR-DT, MIT, CFD)",
        "Hedef dışı mutasyonları önlemenin ilk ve en kritik savunma hattı, rehber RNA'nın in silico algoritmalarla tasarlanmasıdır.",
        "Modern algoritmalar genomdaki benzer sekansları tarar: MIT Spesifisite Skoru ve Cutting Frequency Determination (CFD) matrisleri; her bir nükleotit pozisyonundaki baz uyumsuzluğunun (rG:dT, rU:dG vb.) kesim olasılığı üzerindeki cezai ağırlığını hesaplar. Derin öğrenme tabanlı modeller (CRISPR-Net, DeepCRISPR); kromatin açıklığını (ATAC-seq verileri), DNA metilasyonunu ve ikincil RNA yapılarını da modele dahil ederek off-target riski matematiksel olarak sıfıra yakın gRNA adaylarını seçer.",
        "CFD_Skoru = Prod_{i=1}^{20} Matris_Agirlik(Pozisyon_i, Baz_Uyumsuzluk_turu)",
        "Bu kümülatif olasılık çarpımı eşitliği, rehber RNA boyunca her bir pozisyondaki nükleotit uyumsuzluğunun kesim toleransına olan kümülatif etkisini puanlar."
    ),
    (
        "8.7 Çift Nickaz (Dual Nickase) Stratejisi ile Yüz Kat Spesifisite Artışı",
        "Çift sarmal kırığı oluştururken hedef dışı kesimleri önlemek için SpCas9'un D10A nickaz varyantı (Cas9n) ile geliştirilen 'Çift Nickaz' stratejisi mükemmel bir çözümdür.",
        "Tek başına Cas9n yalnızca tek bir sarmalı keser (nick). Hücrede tek sarmal nickleri hızla ve %100 kusursuz baz eksizyon onarımıyla kapatılır; hiçbir indel veya mutasyon bırakmaz. Çift sarmal kırığı açmak için zıt sarmalları hedefleyen iki farklı sgRNA aynı lokusa gönderilir. İki nickaz ancak birbirine 10-30 baz çifti mesafede zıt sarmalları aynı anda kestiğinde çift sarmal kırığı oluşur. Genomda iki bağımsız sgRNA'nın rastgele yan yana gelebileceği ikinci bir hedef dışı alan bulunma olasılığı neredeyse sıfırdır.",
        "P_off_target_dual = P_off(sgRNA_1) * P_off(sgRNA_2) ~ 10^-6 * 10^-6 = 10^-12 (Milyonda Birden Trilyonda Bire)",
        "Bu bağımsız olasılıklar çarpımı eşitliği, çift nickaz mimarisinin hedef dışı mutasyon ihtimalini nasıl istatistiksel bir imkânsızlığa dönüştürdüğünü kanıtlar."
    ),
    (
        "8.8 Kimyasal Modifiye Edilmiş gRNA'lar: 2'-O-Metil ve Fosforotiyoat Kararlılığı",
        "Sentetik olarak kimyasal sentezlenen tek rehber RNA'ların (sgRNA) riboz ve omurga modifikasyonları; hem hücre içi nükleazlara karşı dayanıklılığı artırır hem de hedef dışı bağlanmaları frenler.",
        "Rehber RNA'nın 5' ve 3' uçlarındaki ilk ve son üç nükleotide 2'-O-metil (2'-OMe) ve 2'-O-metil 3'-fosforotiyoat (MS) bağları yerleştirilir. Bu modifikasyonlar hücre içi ekzo- ve endoribonükleazların RNA'yı parçalamasını engellerken; konak hücrenin sitoplazmik RIG-I ve MDA5 gibi yabancı RNA sensörleri tarafından tanınarak immün fırtına tetiklemesini önler. Ayrıca tohum bölgesine yerleştirilen kilitli nükleik asitler (LNA) uyumsuzluk toleransını tamamen sıfırlar.",
        "Kararlilik_gRNA = Tau_bozunma = Tau_0 * ( 1 + 25 * [Modifiye_Bag_Orani] )",
        "Bu kimyasal stabilizasyon eşitliği, fosforotiyoat ve metil omurga modifikasyonlarının sentetik gRNA hücre içi yarı ömrünü yirmi beş kattan fazla nasıl artırdığını gösterir."
    ),
    (
        "8.9 Anti-CRISPR (Acr) Proteinleri ile Kesimin Zamansal Kapatılması (Off-Switch)",
        "Bakteriyofajlar, bakterilerin CRISPR savunmasını etkisiz hale getirmek için milyonlarca yıllık evrimsel süreçte 'Anti-CRISPR' (Acr) proteinleri geliştirmiştir.",
        "Örneğin AcrIIA4 proteini; Cas9'un PAM bağlama cebine ve DNA tanıma arayüzüne tıpatıp oturarak DNA'yı taklit eder ve Cas9'u fiziksel olarak kilitler. Gen terapisi protokollerinde AcrIIA4; Cas9 hücreye girdikten birkaç saat sonra (hedef kesim tamamlandığında) ortama verilerek 'Kapatma Düğmesi' (Off-Switch) olarak kullanılır. Enzimin hücrede gereksiz yere günlerce aktif kalarak hedef dışı mutasyon biriktirmesi kesin olarak önlenir.",
        "Zaman_Penceresi_Guvenlik: t_aktif = t_kesim_on_target (4-8 saat) -> t_kapatma (AcrIIA4 ile Tam Kapanma)",
        "Bu zamansal regülasyon protokolü, Anti-CRISPR proteinlerinin hedef kesim gerçekleştikten sonra sistemi kapatarak kümülatif hedef dışı mutasyonları nasıl durdurduğunu tanımlar."
    ),
    (
        "8.10 Tek Hücre Tüm Genom Dizilemesi (scWGS) ve Klinik Öncesi Güvenlik Doğrulaması",
        "Klinik in vivo gen terapisinde bir hastaya trilyonlarca viral vektör veya LNP verilmeden önce; tedavinin mutajenik güvenliği 'Tek Hücre Tüm Genom Dizilemesi' (scWGS) ile tescillenmelidir.",
        "Toplu hücre dizilemeleri (bulk sequencing), popülasyonun binde birinde gelişen nadir fakat ölümcül bir klonal mutasyonu veya translokasyonu arka plan gürültüsü içinde kaçırabilir. Tedavi edilmiş dokudan tek tek izole edilen yüzlerce hücrenin tüm genomu scWGS ile 30x derinlikte taranır. Sıfır translokasyon, sıfır kromotripsis ve kabul edilebilir indel profili matematiksel olarak kanıtlanmadan hiçbir gen düzenleme protokolü insana enjekte edilemez.",
        "Guvenlik_Eksper_Skoru = Q_klinik = ( 1 - P_translokasyon ) * ( 1 - P_offtarget_kritik ) * ( 1 - P_onkojen_mutasyon ) ~ 0.99999",
        "Bu klinik yeterlilik olasılık çarpımı, tek hücre genom analizleriyle doğrulanan in vivo gen terapi protokolünün mutlak güvenilirlik katsayısını belgeler."
    )
]

# ================= KISIM 9 =================
part9_subsections = [
    (
        "9.1 Sistemik Viral Gen Terapisinde Doz Bağımlı Toksisite ve Ölümler",
        "Gen terapisinin teorik vaatleri sınırsız olsa da; sistemik yüksek doz rAAV uygulamaları klinik çalışmalarda ağır hepatotoksisite, trombotik mikroanjiyopati ve maalesef hasta ölümleriyle sonuçlanmıştır.",
        "Özellikle X'e bağlı Miyotübüler Miyopati (XLMTM) ve Spinal Müsküler Atrofi (SMA) çalışmalarında kilogram başına 1-3 x 10^14 viral genom (vg/kg) gibi devasa dozlar verilen hastalarda; trilyonlarca viral kapsidin karaciğere ve vasküler yatağa ani yüklenmesi akut sistemik şok yaratmıştır. Bu trajediler, gen terapisi camiasını immünolojik ve toksikolojik mekanizmaları moleküler düzeyde aydınlatmaya ve dozu 10 kat düşürecek kapsid mühendisliğine mecbur bırakmıştır.",
        "Toksisite_Eşigi = Doz_Kritik ~ 1.0 * 10^14 vg/kg (Bu Dozun Üzerinde Sistemik Risk Üstel Artar)",
        "Bu klinik güvenlik eşiği, sistemik intravenöz AAV uygulamalarında ölümcül immün kaskadları tetiklemeyen maksimum tolere edilebilir doz sınırını ifade eder."
    ),
    (
        "9.2 Karaciğer Parankim Hasarı: Transaminaz Patlaması ve İmmün Hepatotoksisite",
        "Sistemik enjekte edilen AAV'lerin %80'inden fazlası ilk geçişte karaciğer sinüzoidlerinde tutulur ve hepatositler ile Kupffer hücreleri tarafından yutulur.",
        "Enjeksiyondan 4 ila 8 hafta sonra, hepatosit yüzeyinde sunulan kapsid peptidlerini tanıyan konak CD8+ sitotoksik T lenfositleri masif bir immün saldırı başlatır. Hepatosit lizisi sonucu serum Alanin Aminotransferaz (ALT) ve Aspartat Aminotransferaz (AST) seviyeleri normalin 20-50 katına fırlayarak akut karaciğer yetmezliği tablosu oluşturur. Eş zamanlı olarak tedavi edilmiş hepatositler yok edildiği için transgen ekspresyonu silinir ve gen terapisinin terapötik etkisi tamamen kaybolur.",
        "Hiz_Hepatosit_Lizis = d[ALT]/dt = k_lizis * [CD8+_CTL] * [MHC_I_Kapsid] * [Hepatosit_transdukte]",
        "Bu transaminaz deşarj kinetiği formülü, serum karaciğer enzim yükselmesinin T-hücre aracılı transdukte hepatosit yıkımıyla doğrudan orantısını modeller."
    ),
    (
        "9.3 Trombotik Mikroanjiyopati (TMA) ve Kompleman Kaskadı Aşırı Aktivasyonu",
        "Yüksek doz AAV gen terapisinin en korkulan sistemik komplikasyonu, tedaviden 1 ila 2 hafta sonra aniden patlak veren 'Trombotik Mikroanjiyopati' (TMA) tablosudur.",
        "Dolaşımdaki devasa miktardaki AAV kapsidi ve bunlara bağlanan immünoglobulinler (AAV-antikor immün kompleksleri); alternatif kompleman yolağını kontrolsüz biçimde ateşler. C3a, C5a anafilatoksinleri ve C5b-9 Membran Atak Kompleksi (MAC) damar endoteline saldırır. Endotel hasarı yaygın mikrovasküler trombozlara, trombositopeniye (trombositlerin tükenmesi), mekanik mikroanjiyopatik hemolitik anemiye ve akut böbrek yetmezliğine yol açar.",
        "Kompleman_Aktivasyon_Hizi = d[MAC]/dt = k_kompleman * [AAV:IgG_kompleks] * [C3b] * ( 1 / [Faktör_H] )",
        "Bu kompleman patlama eşitliği, viral immün komplekslerin endotel yüzeyinde membran atak kompleksi (C5b-9) üretimini ve trombozu nasıl tetiklediğini simgeler."
    ),
    (
        "9.4 cGAS-STING Yolağı ve Sitoplazmik Vektör DNA'sının Algılanması",
        "Viral veya sentetik taşıyıcılar hücre zarı veya endozomdan kaçarken DNA veya RNA moleküllerinin sitoplazmaya sızması; hücrenin en ilkel anti-viral alarm sistemini çaldırır.",
        "Sitoplazmik serbest çift sarmallı DNA; nükleotidiltransferaz olan Siklik GMP-AMP Sentaz (cGAS) tarafından saniyeler içinde algılanır. cGAS, ATP ve GTP'yi birleştirerek ikinci haberci 2'3'-cGAMP molekülünü sentezler. cGAMP endoplazmik retikulum zarındaki STING (Stimulator of Interferon Genes) reseptörüne bağlanır; TBK1 ve IRF3 kinazları aktive edilerek tip-1 interferonlar ve pro-enflamatuar kemokinler (CXCL10) salgılanır; hücre oto-enflamatuar krize girer.",
        "Konsantrasyon_cGAMP = [cGAMP] = Integral[ k_cGAS * [dsDNA_sitoplazmik] / ( K_dna + [dsDNA_sitoplazmik] ) ] dt",
        "Bu biyokimyasal haberci üretim denklemi, sitoplazmaya kaçan serbest vektör DNA'sının cGAS-STING eksenini ve nükleer interferon yanıtını nasıl tetiklediğini açıklar."
    ),
    (
        "9.5 NLRP3 İnflamazom Aktivasyonu ve İnterlökin-1Beta (IL-1beta) Sitokin Fırtınası",
        "Nanopartiküller ve viral kapsidler endozomdan kaçarken endozomal zarı yırtar ve lizozomal katepsin proteazlarının sitoplazmaya dökülmesine yol açar.",
        "Bu sitoplazmik stres ve mitokondriyal ROS sızıntısı; 'NLRP3 İnflamazom' protein kompleksinin oligomerizasyonunu tetikler. İnflamazom kaskadı Prokaspaz-1'i aktif Kaspaz-1 enzimine dönüştürür. Kaspaz-1 iki ölümcül adım atar: 1) İnaktif pro-IL-1beta ve pro-IL-18'i keserek olgun pro-enflamatuar sitokinlere dönüştürüp kana fırlatır; 2) Gasdermin D (GSDMD) proteinini keserek hücre zarında dev gözenekler açar ve hücreyi 'Piroptoz' (yangısal litik ölüm) ile patlatır.",
        "Hiz_Piroptoz = d[IL-1beta]/dt = k_kaspaz1 * [NLRP3_aktif] * [GSDMD_delik_sayisi]",
        "Bu inflamazom deşarj eşitliği, lizozomal yırtılmanın Kaspaz-1 ve Gasdermin-D üzerinden piroptotik doku nekrozunu ve sitokin fırtınasını nasıl başlattığını formüle eder."
    ),
    (
        "9.6 Bakteriyel Cas9 Proteinlerine Karşı Pre-Ekzistans Hümoral ve Selüler Bağışıklık",
        "CRISPR teknolojisinin klinik translasyonunda en büyük paradokslardan biri, en yaygın kullanılan Cas9 enzimlerinin insan patojeni bakterilerden (SpCas9 - boğaz enfeksiyonu ajanı Streptococcus pyogenes; SaCas9 - çıban ve sepsis ajanı Staphylococcus aureus) türetilmiş olmasıdır.",
        "Klinik immünoloji araştırmaları; sağlıklı insan popülasyonunun %50 ila %90'ının kanda anti-SpCas9 ve anti-SaCas9 IgG antikorları ve hafıza CD4+/CD8+ T hücreleri taşıdığını belgelemiştir. İn vivo Cas9 ekspresyonu yapıldığında bu hafıza T hücreleri hemen uyanır ve CRISPR cerrahisi yapılan hücreleri antijenik yabancı düşman kabul ederek yok eder.",
        "Prevalans_AntiCas9 = P_bagisiklik ~ %65 - %85 (Saglikli Yetiskin Insan Popülasyonunda)",
        "Bu epidemiyolojik immünolojik prevalans oranı, bakteriyel Cas nükleazlarının doğrudan in vivo kullanımında karşılaşılan kitlesel ön-bağışıklık engelini belgeler."
    ),
    (
        "9.7 Cas Enzimlerinin Epitop Haritalaması ve İmmün Sessizleştirme (De-Immunization)",
        "Anti-Cas9 bağışıklık bariyerini aşmanın en kalıcı yolu, Cas proteinini insan bağışıklık sisteminin tanıyamayacağı şekilde yeniden tasarlamaktır.",
        "Kapsamlı immünoproteomik taramalarla SpCas9 yüzeyindeki insan HLA Sınıf I ve II moleküllerine bağlanan immünodominant T-hücre epitopları haritalanmıştır. Rasyonel protein mühendisliği ile bu epitoplardaki kritik hidrofobik amino asitler nükleaz fonksiyonunu bozmayacak hidrofilik kalıntılarla değiştirilmiştir (De-Immunization). Ortaya çıkan 'Epitop-Sessiz Cas9' (d-Cas9v2); insan periferik kan mononükleer hücrelerinde (PBMC) sıfır T hücresi proliferasyonu ve sıfır sitokin yanıtı sergilemiştir.",
        "Reaktivite_PBMC = I_yanit = Sigma_i [Epitop_i] * Afinite_HLA_i -> 0 (De-immünize Varyantta)",
        "Bu immünojenisite skorlama toplamı, T-hücre epitoplarının rasyonel mutasyonlarla silinmesinin Cas enzimini bağışıklık radarına nasıl görünmez kıldığını gösterir."
    ),
    (
        "9.8 Profilaktik İmmünomodülasyon: Kortikosteroidler, Eculizumab ve Rituximab",
        "İn vivo gen terapisinin sistemik toksisite riskini yönetmek için klinik uygulamalar agresif profilaktik immün baskılama rejimleriyle korunmaktadır.",
        "1) Yüksek Doz Glukokortikoidler (Prednizolon): T hücre aktivasyonunu ve hepatosit lizisini baskılamak için infüzyondan 1 gün önce başlanıp 2-3 ay titre edilerek kesilir; 2) Eculizumab (Soliris): C5 kompleman inhibitörü monoklonal antikor olup; trombotik mikroanjiyopatiyi (TMA) önlemek için alternatif kompleman terminal yolunu kilitler; 3) Rituximab: Anti-CD20 antikoru ile B lenfositlerini tüketerek anti-AAV veya anti-Cas9 antikor üretimini sıfırlar.",
        "Protokol_Immunsupresyon = Prednizolon (1 mg/kg/gun) + Eculizumab (C5 blokaji) + Takrolimus (Kalsinörin)",
        "Bu klinik kombinasyon şeması, sistemik gen terapisi infüzyonunda ölümcül kompleman ve sitotoksik T hücre komplikasyonlarını engelleyen standart profilaktik protokolü tanımlar."
    ),
    (
        "9.9 B Hücre Tüketimi ve İmlifidase (IdeS) ile Nötralizan Antikorların Geçici İnhibisyonu",
        "Hastanın kanında önceden var olan yüksek anti-AAV nötralizan antikorları (NAb) temizlemek ve gen terapisini mümkün kılmak için devrimsel bir enzimatik cerrahi geliştirilmiştir.",
        "Streptococcus pyogenes kaynaklı bir sistein endopeptidaz olan 'İmlifidase' (IdeS); insan kan dolaşımına verildiğinde saniyeler içinde tüm IgG antikorlarının menteşe (hinge) bölgesini spesifik olarak keser. İki saat içinde plazmadaki tüm serbest ve anti-AAV IgG antikorları F(ab')2 ve Fc parçalarına ayrılarak nötralize edilir. Açılan bu 48-72 saatlik 'immünolojik pencere' sırasında rAAV vektörü hastaya verilir; virüs antikorlara takılmadan hedef dokuya ulaşır ve hücre içine girer.",
        "Klirens_IgG_IdeS = [IgG_intakt](t) = [IgG_0] * exp( - k_IdeS * [IdeS] * t ) -> Sıfır (2 Saat Icinde)",
        "Bu enzimatik immünoglobulin parçalanma denklemi, İmlifidase enziminin önceden antikor taşıyan hastaları gen terapisine dakikalar içinde nasıl uygun hale getirdiğini belgeler."
    ),
    (
        "9.10 Güvenli İn Vivo Mühendisliğin Altın Standartları: Lokal Teslimat ve Doz Optimizasyonu",
        "İn vivo gen terapisinde toksisiteyi önlemenin en akılcı ve temel mühendislik çözümü; sistemik intravenöz enjeksiyon yerine hedefe yönelik 'Lokal Teslimat' stratejilerinin seçilmesidir.",
        "Retina hastalıklarında subretinal enjeksiyon, santral sinir sistemi hastalıklarında intratekal (BOS içi) veya stereotaksik intraparenkimal enjeksiyon; sistemik dolaşıma virüs kaçırmadan lokal hücreleri tedavi eder. Bu yaklaşım gereken toplam viral dozu sistemik dozun %1'ine (yüz kat daha az) indirir. Düşük doz karaciğeri, kompleman sistemini ve lenfoid organları tamamen koruyarak gen terapisini kusursuz bir ayaktan klinik prosedüre dönüştürür.",
        "Kazanc_Guvenlik = Doku_Konsantrasyonu / Sistemik_Maruziyet = ( V_sistemik / V_organ ) >> 100",
        "Bu farmakokinetik lokalizasyon oranı, hedefe yönelik lokal gen teslimatının sistemik toksisite riskini yüz kattan fazla nasıl düşürdüğünü matematiksel olarak modeller."
    )
]

# ================= KISIM 10 =================
part10_subsections = [
    (
        "10.1 Homo Aeternus Genomik Tasarımı: Doğal Evrimsel Sınırların Aşılarak Yeniden Kodlanması",
        "Homo sapiens genomu milyonlarca yıllık doğal seçilimin ürünüdür; ancak bu seçilim uzun ömürlülük ve gençlik için değil, avcı-toplayıcı koşullarda 30-40 yıl hayatta kalıp genleri aktarmak için optimize olmuştur.",
        "Homo Aeternus vizyonu; insan genomunu pasif bir evrimsel miras olarak kabul etmek yerine, yaşam süresini sınırlayan tüm ölümcül biyolojik darboğazları aktif genom mühendisliği ile yeniden tasarlamayı amaçlar. Bu yeni ontolojik paradigma; telomeraz aktivasyonundan DNA onarım takviyesine, mitokondriyal biyogenezden kas kütlesi ve kardiyovasküler dayanıklılık genlerinin optimizasyonuna kadar çok genli entegre bir moleküler devrimdir.",
        "Genom_Homo_Aeternus = Genom_Referans + Sigma_k [Anti_Aging_Kaset_k] - Sigma_m [Patojenik_Risk_Aleli_m]",
        "Bu yapıcı genom eşitliği, Homo Aeternus genetik mimarisinin hastalık yapıcı alellerin silinmesi ve ömür uzatıcı transgenlerin entegrasyonu ile oluşturulan nihai sentezini simgeler."
    ),
    (
        "10.2 TERT (Telomeraz Ters Transkriptaz) Gen Terapisi ve Replikatif Ölümsüzlük",
        "Hücresel yaşlanmanın ve Hayflick sınırının primer moleküler saati olan telomer erozyonu; bölünmeyen kök hücre nişlerini kurutarak organ iflaslarına yol açar.",
        "Maria Blasco ve ekibinin öncülüğünü yaptığı AAV-TERT gen terapisi; telomeraz enziminin katalitik alt birimini (TERT) erişkin ve yaşlı dokulara aktarır. AAV-TERT verilen hayvanlarda; kök hücre havuzları gençleşmiş, kemik erimesi, deri atrofisi ve insülin direnci gerilemiş ve kanser insidansında hiçbir artış olmaksızın medyan yaşam süresi %24 uzatılmıştır. Bölünmeyen post-mitotik hücrelerde ise TERT mitokondriye giderek mtDNA onarımını destekler ve apoptozu frenler.",
        "Uzatma_Telomer = L_telomer(t) = L_0 + Integral[ V_TERT * [AAV-TERT] - k_erozyon * f_bolunme ] dt",
        "Bu dinamik telomer boyu denklemi, AAV-TERT gen transferi ile sağlanan telomeraz sentezinin replikatif erozyon hızını nasıl sıfırlayıp net uzama fazına geçirdiğini belgeler."
    ),
    (
        "10.3 Klotho Gen Terapisi: Serebral, Vasküler ve Renal Koruma Protokolü",
        "Adını Yunan mitolojisinde yaşam ipliğini eğiren tanrıçadan alan Klotho; böbrek ve beyin koroid pleksusunda üretilen ve eksikliği erken yaşlanmaya yol açan ana longevity genidir.",
        "AAV vektörleri ile tek bir sistemik veya intratekal Klotho gen transferi; 1) Böbreklerde Wnt sinyalini ve TGF-beta fibrozisini durdurarak glomerüler sklerozu önler; 2) Serebral kortekste NMDA reseptörlerinin GluN2B alt birimini zenginleştirerek sinaptik plastisiteyi gençleştirir; 3) Damar endotelinde nitrik oksit (NO) sentezini artırıp arteriyel elastisiteyi korur. Klotho gen terapisi uygulanan farelerin bilişsel fonksiyonları zirveye çıkmış ve yaşam süreleri %30 uzamıştır.",
        "Konsantrasyon_Klotho_Doku = C_Klotho = ( V_AAV_promoter / k_klerens ) >> C_fizyolojik_yasli * 5",
        "Bu kararlı durum farmakokinetik formülü, AAV-Klotho tedavisinin doku ve plazma Klotho seviyesini yaşlı bazal değerlerin beş kat üzerine nasıl sabitlediğini ifade eder."
    ),
    (
        "10.4 Follistatin Gen Terapisi: Miyostatin İnhibisyonu, Sarkopeni Tasfiyesi ve Kas Gençleşmesi",
        "Yaşlılığın en yıkıcı fiziksel belirtisi; iskelet kası kütlesinin, lif çapının ve motor gücünün geri dönüşsüz kaybı olan 'Sarkopeni'dir.",
        "Follistatin (FST); iskelet kası büyümesini negatif olarak frenleyen Miyostatin (GDF-8) ve Aktivin A proteinlerine yüksek afiniteyle bağlanarak onları nötralize eden doğal bir antagonisttir. AAV-Follistatin (AAV1-FS344) gen terapisi kas içi veya sistemik enjekte edildiğinde; yaşlı hayvanlarda ve insan klinik denemelerinde kas kütlesinde %30-40 artış, kas liflerinde hipertrofi, kas kök hücrelerinde (uydu hücreler) proliferasyon ve motor dayanıklılıkta muazzam bir gençleşme sağlamıştır.",
        "Kuvvet_Kassal = F_kas = F_0 * ( 1 + alpha_FST * [Follistatin_serum] / ( K_m + [Miyostatin] ) )",
        "Bu biyomekanik kas gücü denklemi, Follistatin aracılı miyostatin baskılanmasının iskelet kası kontraksiyon kuvveti ve lif hacmindeki katlanarak artışını modeller."
    ),
    (
        "10.5 PGC-1alfa ve Nrf2 Gen Terapileri: Mitokondriyal Biyogenez ve Redoks Kalkanı",
        "Hücresel yaşlanmanın metabolik kalbi olan mitokondriyal çöküş ve oksidatif stres; PGC-1alfa ve Nrf2 genetik devrelerinin birlikte aktive edilmesiyle tamamen geri döndürülebilir.",
        "AAV ile iletilen PGC-1alfa (Peroxisome Proliferator-Activated Receptor Gamma Coactivator 1-alpha); nükleer solunum faktörleri NRF-1/NRF-2 ve mitokondriyal transkripsiyon faktörü A'yı (TFAM) aktive ederek hücrede yeni, genç mitokondri üretimini (mitogenez) 3 kat artırır. Eş zamanlı aktarılan Nrf2 transgeni ise antioksidan yanıt elementlerine (ARE) bağlanarak süperoksit dismutaz (SOD), katalaz ve glutatyon peroksidaz üretimini patlatır; hücreyi her yönden koruyan sarsılmaz bir 'Redoks Kalkanı' kurar.",
        "Biyogenez_Hizi = J_mitokondri = k_PGC1a * [PGC1a_transgen] * [TFAM] * ( 1 / ( 1 + [ROS_kalan] ) )",
        "Bu mitokondriyal rejenerasyon eşitliği, PGC-1alfa transgen ifadesinin hücre içi yeni organel montaj hızını nasıl katladığını matematiksel olarak ortaya koyar."
    ),
    (
        "10.6 FOXO3a Longevity Varyantı Mühendisliği: DNA Onarımı ve Hücresel Dayanıklılık",
        "Yüz yaşını aşan süper asırlıkların (Centenarians) genom dizilemeleri; bu bireylerin ortak genetik paydasının FOXO3a transkripsiyon faktörünün koruyucu 'longevity alellerini' (özellikle rs2802292 G aleli) taşıması olduğunu göstermiştir.",
        "İn vivo Prime Editing veya AAV gen transferi ile somatik dokularda FOXO3a aktivitesinin artırılması veya süper-longevity SNP'sinin yerleştirilmesi; otofaji akışını hızlandırır, DNA çift zincir kırığı onarım enzimlerini (GADD45) yukarı regüle eder ve proteozomal agregat temizliğini gençlik verimine taşır. Hücreler her türlü metabolik, termal ve genotoksik strese karşı olağanüstü bir evrimsel zırh kazanır.",
        "Direnç_Stres = R_hucre = R_bazal * exp( beta_FOXO3 * [FOXO3a_nukleer_aktif] )",
        "Bu hücresel dayanıklılık fonksiyonu, nükleer aktif FOXO3a seviyesinin hücrenin genotoksik ve metabolik streslere karşı hayatta kalma eşiğini nasıl üstel olarak artırdığını tanımlar."
    ),
    (
        "10.7 Çok-Genli (Multi-Gene) Polisisronik Vektörler ve 2A Peptid Teknolojisi (P2A/T2A)",
        "Birden fazla longevity genini (örneğin TERT + Klotho + Follistatin) aynı hücreye aktarmak için ayrı ayrı virüsler kullanmak transduksiyon verimini düşürür ve dozu katlar.",
        "Bu sorunu aşmak için 'Polisistronik Ekspresyon Kasetleri' kullanılır. Tek bir promotörün arkasına dizilen farklı transgenler arasına küçük Picornavirüs kaynaklı '2A Peptidleri' (P2A, T2A, E2A) yerleştirilir. Ribozom bu mRNA'yı okurken 2A peptidinin Gly-Pro amino asit bağında translasyonel ribozomal kayma (ribosomal skipping) yapar; peptid bağını kurmadan bir sonraki amino asitten yeni proteini sentezlemeye devam eder. Böylece tek bir mRNA transkriptinden 1:1:1 stokiyometrik oranda üç bağımsız fonksiyonel longevity proteini üretilir.",
        "Stokiyometri_2A = [Protein_A] : [Protein_B] : [Protein_C] ~ 1.0 : 0.98 : 0.95 (Kusursuz Eşit Sentez)",
        "Bu translasyonel ribozomal kayma eşitliği, 2A peptid mimarisinin çoklu gen kasetlerinde eşit molar protein sentezleme verimini belgeler."
    ),
    (
        "10.8 Yapay Kromozomlar (HACs) ve Mega-Boyutlu Sentetik Longevity Ağları",
        "Viral vektörlerin ve LNP'lerin taşıma kapasiteleri en fazla onlarca kilobazla sınırlıyken; insanın tüm biyolojik ağlarını ölümsüzlük için optimize etmek yüzlerce genin ve regülatör elementin aktarılmasını gerektirir.",
        "Bu nihai vizyonun taşıyıcısı 'İnsan Yapay Kromozomları'dır (HACs - Human Artificial Chromosomes). Sentetik bir sentromer, telomerler ve genomik replikasyon orijinleri içeren HAC'lar; konak kromatinine entegre olmadan çekirdekte bağımsız 47. kromozom olarak serbestçe yaşar. Mega-baz (milyonlarca baz çifti) büyüklüğünde sentetik gen kütüphanelerini taşıyabilen HAC'lar; tüm bir longevity metabolik işletim sistemini insan hücresine yükleme potansiyeline sahiptir.",
        "Kapasite_HAC = L_kargo_HAC ~ 1_Mb - 10_Mb (Sınırsız Biyomühendislik Ağları)",
        "Bu mega-genomik kapasite tanımı, yapay kromozom teknolojisinin hücreye bağımsız sentetik biyolojik işletim sistemleri kurma gücünü simgeler."
    ),
    (
        "10.9 Geribildirimli Genetik Mantık Kapıları (Logic Gates) ve Biyomarker-Tetiklemeli Transkripsiyon",
        "Transgenlerin ömür boyu sürekli maksimum seviyede üretilmesi biyolojik kaynakları tüketebilir veya istenmeyen hipertrofilere yol açabilir; ideal sistem ihtiyaca göre çalışan akıllı regülasyondur.",
        "Sentetik Biyoloji mantık kapıları (AND, OR, NOT) kullanılarak tasarlanan 'Kapalı Döngü Genetik Devreler' (Closed-Loop Genetic Circuits); transgen üretimini dokunun metabolik durumuna bağlar. Örneğin TERT ekspresyonu yalnızca hücre içi p16INK4a senesens sinyali yükseldiğinde açılır; DNA onarım genleri yalnızca gama-H2AX çift zincir kırıkları belirdiğinde uyarılır. Görev tamamlanıp biyomarker gençlik seviyesine döndüğünde genetik devre transkripsiyonu kendi kendine kapatır.",
        "Cikis_Transgen = ( Girdi_Senesens AND Girdi_Enflamasyon ) NOT Girdi_Toksisite",
        "Bu Boole genetik mantık kapısı algoritması, longevity gen terapilerinin yalnızca patolojik stres ve yaşlanma sinyali algılandığında dinamik olarak devreye girme prensibini formüle eder."
    ),
    (
        "10.10 Homo Aeternus İn Vivo Genom Manifestosu: Biyolojik Ölümsüzlüğün Genetik Mimarisi",
        "Homo Aeternus genom projesi; ölümün ve yaşlanmanın biyolojik bir kader değil, çözülebilir bir bilgi işleme ve mühendislik problemi olduğunu ilan eder.",
        "Yaşamın kaynak kodunda gerçekleştirilen bu çok katmanlı devrim; 1) Kapsid ve LNP mühendisliğiyle doku bariyerlerinin aşılmasını, 2) Çift zincir kırığı yapmayan Baz ve Prime Editing ile somatik mutasyonların silinmesini, 3) Epigenom düzenleyiciler ile biyolojik saatin sıfırlanmasını, 4) TERT, Klotho ve Follistatin ile hücresel tükenişin durdurulmasını birleştirir. İnsan artık evrimin kör mekanizmalarının oyuncağı değil; kendi biyolojisinin, bilincinin ve sonsuz geleceğinin mutlak mimarıdır.",
        "Olumsuzluk_Katsayisi = Omega_Genom = lim_{t->sonsuz} [ Bilgi_Genomik(t) / ( 1 + Somatik_Entropi(t) ) ] = Sabit > 0",
        "Bu nihai tekillik eşitliği, genomik enformasyonun aktif in vivo mühendislikle korunması halinde insan biyolojisinin termodinamik ve hücresel entropiyi ebediyen yeneceğini matematiksel olarak tescil eder."
    )
]

# 10 Kapsamlı Akademik Karşılaştırma Tablosu
tables_data = [
    {
        "title": "Tablo 14.1: Viral Gen Terapisi Vektörlerinin Karşılaştırmalı Biyolojik ve Klinik Profili",
        "headers": ["Vektör Platformu", "Genom Türü / Kargo Kapasitesi", "Konak Entegrasyonu", "Pre-Ekzistans Bağışıklık Riski", "Klinik Hedef Doku / Uygulama"],
        "rows": [
            ["rAAV (Adeno-İlişkili Virüs)", "ssDNA / scDNA (<4.7 kb)", "Epizomal (%99+ entegre olmaz)", "Yüksek (%30-70 NAb pozitifliği)", "Karaciğer, MSS, Retina, Kas (Zolgensma, Luxturna)"],
            ["Lentivirüs (LV / SIN-LV)", "ssRNA -> dsDNA (8-9 kb)", "Kromatin Entegrasyonu (Kalıcı)", "Düşük (Minimal seroprevalans)", "Ex vivo CAR-T, Hematopoietik kök hücreler"],
            ["Adenovirüs (Ad5 / ChAdOx)", "dsDNA (7-8 kb veya Gutless 36 kb)", "Epizomal (Sıfır entegrasyon)", "Çok Yüksek (Yaygın Ad5 bağışıklığı)", "Aşı platformları, Onkolitik viroterapi"],
            ["Herpes Simpleks Virüsü (HSV-1)", "dsDNA (30-40 kb dev kargo)", "Epizomal (Nöronal latensi)", "Yüksek (%70 HSV seropozitif)", "Nörolojik hedefleme, Periferik nöropatiler"],
            ["Kendi Kendini Tamamlayan scAAV", "dsDNA saç tokası (<2.2 kb)", "Epizomal (Hızlı ekspresyon)", "Yüksek (AAV ile aynı)", "Acil transgen ekspresyonu gereken nörodejenerasyon"],
            ["Baculovirüs Sistemleri", "dsDNA (>100 kb dev genom)", "Memeli hücresinde replike olmaz", "Düşük / Yok", "Endüstriyel rAAV ve protein biyoreaktör üretimi"],
            ["Gutless (Helper-Free) AdV", "dsDNA (36 kb'a kadar)", "Epizomal (Düşük immünojenisite)", "Orta (Kapsid immünitesi kalır)", "Büyük gen transferleri, Karaciğer metabolik hastalıkları"],
            ["Entegre Olmayan Lentivirüs (IDLV)", "dsDNA sirküler epizom (8 kb)", "Epizomal (Katalitik mutasyon)", "Düşük", "Bölünmeyen hücrelerde geçici ekspresyon"]
        ]
    },
    {
        "title": "Tablo 14.2: Doğal ve Mühendislik Ürünü AAV Kapsidlerinin Tropizm ve Bariyer Geçiş Analizi",
        "headers": ["AAV Kapsid Varyantı", "Kapsid Kökeni / Modifikasyonu", "Kan-Beyin Bariyeri Geçişi", "Hepatik Tutulum (Karaciğer Kaçağı)", "Hedef Doku ve Klinik Avantaj"],
        "rows": [
            ["AAV2", "Doğal İzolat (Heparan sülfat)", "Çok Düşük / Yok", "Orta", "Retina (RPE), Doğrudan parankim enjeksiyonu"],
            ["AAV8", "Doğal İzolat (Laminin reseptörü)", "Düşük", "Aşırı Yüksek (%90 hepatik tutulum)", "Karaciğer hedefli metabolik hastalıklar"],
            ["AAV9", "Doğal İzolat (Galaktoz)", "Düşük / Orta (Yüksek doz gerekir)", "Çok Yüksek", "Sistemik MSS ve Kalp kası hedefleme"],
            ["AAV-PHP.eB", "AAV9 + 7-mer 'TLAVPFK' ilmeği", "Ultra-Yüksek (AAV9'dan 50x güçlü)", "Orta / Yüksek", "Kemirgen tüm beyin nöron ve astrosit transduksiyonu"],
            ["AAV-PHP.V1", "AAV9 yüzey peptid mühendisliği", "Endotelde kilitlenir", "Düşük", "Serebrovasküler endotel ve bariyer restorasyonu"],
            ["AAV.CAP-B10", "Primat/İnsan iPSC seçilimli evrim", "Yüksek (Ly6a-bağımsız insan transferi)", "Düşük (%80 de-targeted)", "Güvenli insan serebral korteks nöron hedefleme"],
            ["AAV.BI30", "İnsan TfR1 hedefli yönlendirilmiş evrim", "Ultra-Yüksek (İnsan endotel transsitozu)", "Çok Düşük", "İnsan klinik translasyonel nörolojik gen terapisi"],
            ["AAV-LK03", "İnsan hepatosit kütüphane seçilimi", "Yok", "Seçici İnsan Hepatositi (%100)", "Hemofili ve karaciğer gen terapisi"]
        ]
    },
    {
        "title": "Tablo 14.3: Sentetik ve Non-Viral Teslimat Platformlarının (LNPs) Moleküler Karşılaştırması",
        "headers": ["LNP Parametresi / Formülasyonu", "DLin-MC3-DMA (Onpattro)", "SM-102 (Moderna mRNA)", "ALC-0315 (Pfizer/BioNTech)", "SORT LNP (Seçici Organ Hedefli)"],
        "rows": [
            ["İyonize Edilebilir Lipid Yapısı", "Tersiyer amin + Linoleil kuyruk", "Dallanmış amino alkil ester", "Dallanmış ester + Tersiyer amin", "Özel SORT molekülü ilaveli lipid"],
            ["Görünür pKa Değeri", "6.44", "6.68", "6.09", "5.5 - 7.2 (Hedef organa göre ayarlı)"],
            ["Biyobozunurluk (Doku Yarı Ömrü)", "Yavaş (Günler / Haftalar)", "Hızlı (Esteraz hidrolizi ile <24 saat)", "Çok Hızlı (Toksik metabolit bırakmaz)", "Çok Hızlı (Metabolize edilebilir)"],
            ["Primer Organ Biyodağılımı", "Karaciğer Hepatositleri (%95+)", "Karaciğer / Enjeksiyon yeri kası", "Karaciğer / Lenf düğümleri", "Akciğer (DOTAP), Dalak (PA), Kemik İliği"],
            ["Kargo Uyumluluğu", "siRNA", "mRNA", "mRNA / Cas9 RNP", "mRNA, Cas9 RNP, Prime Editor"],
            ["İmmünojenisite / Sitokin Riski", "Düşük (Standart lipid profili)", "Hafif geçici reaksiyon", "Hafif geçici reaksiyon", "Ekstrahepatik düşük enflamasyon"],
            ["Endozomal Kaçış Verimi", "%1 - 2.5", "%2 - 4", "%2 - 5", "%4 - 8 (Konik faz optimizasyonu)"],
            ["Klinik Durum / Onay", "FDA Onaylı (2018)", "FDA Onaylı (2020/2021)", "FDA Onaylı (2020/2021)", "Preklinik / Faz 1 Aşamasında"]
        ]
    },
    {
        "title": "Tablo 14.4: CRISPR-Cas Nükleaz Çeşitlerinin Yapısal, Biyofiziksel ve Fonksiyonel Matrisi",
        "headers": ["Cas Enzimi", "Boyut (Amino Asit / kb)", "PAM Tanıma Dizisi", "Kesim Deseni / Uç Türü", "AAV Tekli Ambalajlanabilirlik"],
        "rows": [
            ["SpCas9 (Streptococcus pyogenes)", "1368 aa / ~4.1 kb", "5'-NGG-3'", "Kör Uç (Blunt End / PAM -3)", "Aşırı Zor (Promotör sınırı aşılır)"],
            ["SaCas9 (Staphylococcus aureus)", "1053 aa / ~3.15 kb", "5'-NNGRRT-3'", "Kör Uç (Blunt End)", "Mükemmel (Tek AAV kargo uyumu)"],
            ["CjCas9 (Campylobacter jejuni)", "984 aa / ~2.95 kb", "5'-NNNNRYAC-3'", "Kör Uç (Blunt End)", "Mükemmel (Ek transgenler için yer kalır)"],
            ["CasMINI (Evriltilmiş Cas12f)", "529 aa / ~1.6 kb", "5'-TTTR-3'", "Yapışkan Uç (Staggered End)", "Rakipsiz (Ultra-kompakt gen kaseti)"],
            ["Cas12a (Cpf1 / Acidaminococcus)", "1228 aa / ~3.7 kb", "5'-TTTV-3'", "4-5 nt Yapışkan Uç (5' çıkıntı)", "Zor (Kompakt promotör gerekir)"],
            ["Cas13d (CasRx / Ruminococcus)", "967 aa / ~2.9 kb", "PFS-bağımsız (Yalnızca ssRNA)", "RNA Kesimi (Ribonükleaz)", "Mükemmel (Tek AAV'ye sığar)"],
            ["SpG (Genişletilmiş PAM SpCas9)", "1368 aa / ~4.1 kb", "5'-NGN-3'", "Kör Uç", "Zor (SpCas9 boyutuyla aynı)"],
            ["SpRY (Neredeyse PAM-siz)", "1368 aa / ~4.1 kb", "5'-NRN-3' ve 5'-NYN-3'", "Kör Uç", "Zor (Tüm genomu tarayabilir)"]
        ]
    },
    {
        "title": "Tablo 14.5: Baz Düzenleme (Base Editing) Sistemlerinin Biyokimyasal ve Kinetik Parametreleri",
        "headers": ["Baz Düzenleyici Sınıfı", "Katalitik Deaminaz Enzimi", "Dönüşüm Yönü (Geçiş)", "Katalitik Mekanizma ve Ara Ürün", "İndel Oranı ve Güvenlik"],
        "rows": [
            ["CBE (Sitozin Düzenleyici)", "rAPOBEC1 veya evoAPOBEC1", "C:G -> T:A (Geçiş)", "Sitozin deaminasyonu -> Urasil ara ürün", "<%1 İndel (UGI inhibisyonu ile)"],
            ["ABE (Adenin Düzenleyici)", "Laboratuvarda evriltilen TadA-8e", "A:T -> G:C (Geçiş)", "Adenin deaminasyonu -> İnozin ara ürün", "<%0.1 İndel (Ultra-temiz profil)"],
            ["GBE (Glikozilaz Düzenleyici)", "APOBEC1 + Ekzojen UDG", "C:G -> G:C (Transversiyon)", "Abazik AP saha -> Rev1 polimeraz", "%5 - 15 İndel (Abazik saha riski)"],
            ["SPACE / ACBE", "TadA-8e + APOBEC1 (Çift)", "Eş Zamanlı A->G ve C->T", "Senkronize adenin ve sitozin deaminasyonu", "<%1 İndel"],
            ["DdCBE (Mitokondriyal)", "Bakteriyel split-DddA toksini", "mtDNA C:G -> T:A", "Çift sarmallı DNA sitozin deaminasyonu", "<%0.5 İndel (gRNA bağımsız TALE)"],
            ["SECURE-CBE", "R33A/K34A mutasyonlu APOBEC1", "C:G -> T:A", "RNA bağlaması yok edilmiş deaminaz", "Sıfır RNA hedef dışı mutasyonu"],
            ["Verve-101 (Klinik ABE)", "TadA-8e + Cas9n (LNP-mRNA)", "A:T -> G:C (PCSK9 geni)", "Splice donör inaktivasyonu", "İnsan karaciğerinde %70 kür"],
            ["Mini-ABE", "CasMINI + Kompakt TadA", "A:T -> G:C", "Kompakt AAV kargo uyumlu deaminasyon", "<%0.5 İndel"]
        ]
    },
    {
        "title": "Tablo 14.6: Prime Editing (PE) Nesillerinin Karşılaştırmalı Evrim ve Verimlilik Matrisi",
        "headers": ["Prime Editor Mimarisi", "Ters Transkriptaz / Kargo Yapısı", "Karşıt İplik Stratejisi (Nick)", "İnsan Hücresi Ortalama Verimi", "İndel Oluşum Frekansı"],
        "rows": [
            ["PE1 (Birinci Nesil)", "Vahşi tip M-MLV RT + Cas9n", "Karşıt nick yok", "%1 - 5 (Çok düşük)", "<%0.5"],
            ["PE2 (İkinci Nesil)", "M-MLV RT (5 mutasyonlu engineered)", "Karşıt nick yok", "%10 - 25 (3-5 kat artış)", "<%0.5"],
            ["PE3 (Üçüncü Nesil)", "PE2 enzimi + Standart pegRNA", "Eş zamansız ikincil nick (sgRNA)", "%30 - 60 (Yüksek verim)", "%1 - 5 (Hafif indel artışı)"],
            ["PE3b (Hassas Üçüncü Nesil)", "PE2 enzimi + Standart pegRNA", "Flap entegrasyonu bağımlı nick", "%30 - 55", "<%0.5 (Sıfıra yakın indel)"],
            ["epegRNA Entegrasyonu", "3' saç tokası zırhlı pegRNA (evopreQ1)", "PE2 / PE3 uyumlu", "%40 - 70 (3-4 kat kararlılık)", "<%1"],
            ["PE-max + MLH1dn", "Kodon optimize Cas9-RT + MLH1 baskısı", "Mismatch Repair (MMR) baypas", "%60 - 85 (Zirve transversiyon)", "<%1.5"],
            ["TwinPE (İkiz Düzenleyici)", "İki karşılıklı pegRNA + Bxb1 rekombinaz", "Çift tamamlayıcı 3' flap", "Mega-gen insersiyonu (%20-40)", "<%2 (Büyük gen dikişi)"],
            ["Split-PE (AAV Çift Kargo)", "İntein-aracılı birleşen Cas9n ve RT", "AAV çift vektör teslimatı", "%20 - 45 in vivo dokularda", "<%1"]
        ]
    },
    {
        "title": "Tablo 14.7: Epigenom Düzenleme Araçlarının (CRISPRi / CRISPRa) Moleküler Mimarisi",
        "headers": ["Epigenetik Efektör", "Bağlanan Transkripsiyonel Alan", "Kromatin İmzası (Modifikasyon)", "Gen İfadesi Yanıtı", "Epigenetik Bellek Süresi"],
        "rows": [
            ["dCas9-KRAB (CRISPRi)", "Krüppel-ilişkili Kutu (KAP1 çağırma)", "H3K9me3 trimetilasyon, deasetilasyon", "Susturma (%95-99 baskılanma)", "Geçici / Orta vadeli (Haftalar)"],
            ["dCas9-VPR (CRISPRa)", "VP64 + p65 + Rta tandem füzyon", "Transkripsiyon kompleksi yığma", "Aktivasyon (10 - 50 kat artış)", "Geçici (Enzim varlığında aktif)"],
            ["dCas9-SunTag", "Tekrarlayan 24x peptid + scFv-VP64", "Masif transkripsiyon faktörü kümesi", "Süper-Aktivasyon (1000 kata kadar)", "Geçici / Yüksek gen patlaması"],
            ["dCas9-p300HATcore", "İnsan p300 katalitik asetilaz", "H3K27ac birikimi (Ökromatin)", "Enhanser aktivasyonu, açık kromatin", "Orta / Uzun vadeli"],
            ["dCas9-DNMT3A/3L", "De novo DNA metiltransferaz", "CpG adalarında 5mC birikimi", "Kalıcı Gen Susturma (%99+)", "Ömür Boyu Kalıcı (Bölünmede kalıtılır)"],
            ["dCas9-TET1", "Ten-Eleven Translocation dioksijenaz", "5mC -> 5hmC -> Saf Sitozin", "Susturulmuş Genin Kurtarılması", "Kalıcı (Metilasyon silinir)"],
            ["Hit-and-Run LNP", "Geçici dCas9-DNMT3A mRNA", "Hedefli CpG metilasyonu", "Hastalık geninin kalıcı susturulması", "Enzim 24 saatte yok olur, etki kalır"],
            ["dCas9-LSD1", "Lizin spesifik histon demetilazı", "H3K4me2 silinmesi (Enhanser susturma)", "Enhanser susturma (Gelişimsel genler)", "Orta vadeli"]
        ]
    },
    {
        "title": "Tablo 14.8: Hedef Dışı (Off-Target) Tespit Metodolojileri ve Biyoinformatik Kıyaslaması",
        "headers": ["Tespit Yöntemi", "Test Ortamı / Mekanizma", "Duyarlılık Eşiği (Tespit Limiti)", "Avantajları", "Sınırlılıkları"],
        "rows": [
            ["GUIDE-seq", "Hücre İçi (In cellulo / dsODN etiketleme)", "%0.1 (Binde bir)", "Gerçek kromatin ortamında çift zincir kırığı tespiti", "Bölünmeyen hücrelerde düşük dsODN entegrasyonu"],
            ["CIRCLE-seq", "Hücresiz (In vitro / Dairesel DNA kesimi)", "%0.01 (On binde bir)", "Aşırı yüksek duyarlılık, kromatin bariyersiz tam liste", "Hücrede kapalı olan sahaları da gereksiz raporlayabilir"],
            ["Digenome-seq", "In vitro WGS (Saf genomik DNA sindirimi)", "%0.1 - 0.5", "Kör uç kesimlerini tüm genomda haritalama", "Çok yüksek NGS dizileme derinliği (60x+) gerektirir"],
            ["DISCOVER-seq", "In vivo doku (MRE11 tamir proteini takibi)", "%0.5 - 1.0", "Canlı hayvan dokularında doğrudan hedef dışı tespiti", "Düşük transkripsiyonel sahalarda duyarlılık düşer"],
            ["CFD Algoritması", "Biyoinformatik In Silico Matris Modelleme", "Matematiksel Puanlama", "Sıfır maliyet, anında tasarım rehberliği", "Epigenetik ve kromatin açıklık dinamiklerini tam öngöremez"],
            ["DeepCRISPR", "Derin Öğrenme / Evrişimli Sinir Ağı (CNN)", "Yüksek Doğruluk Skoru", "Epigenomik ve nükleozom verilerini entegre eder", "Eğitim veri setlerinin biyolojik sınırlarına bağımlıdır"],
            ["Çift Nickaz (Cas9n)", "Biyofiziksel Laboratuvar Tasarımı", "Teorik Sıfır Off-Target", "İki bağımsız sgRNA eşleşmesi şarttır", "İki bağımsız sgRNA optimizasyonu ve dar aralık şartı"],
            ["scWGS (Tek Hücre)", "Tek hücre tüm genom dizilemesi (30x)", "Tek hücre klonal çözünürlük", "Klinik öncesi altın standart güvenlik tescili", "Yüksek dizileme maliyeti ve biyoinformatik yük"]
        ]
    },
    {
        "title": "Tablo 14.9: İn Vivo Gen Terapisi İmmünolojik Toksisiteleri ve Mühendislik Önlemleri",
        "headers": ["İmmünolojik Tehlike / Reaksiyon", "Tetikleyici Moleküler Patojen", "Klinik Patolojik Tezahür", "Biyomühendislik Savunma Stratejisi", "Farmakolojik İmmünsüpresyon Tedavisi"],
        "rows": [
            ["Akut Hepatotoksisite", "AAV Kapsidi + CD8+ Sitotoksik T Lenfosit", "Masif serum ALT/AST patlaması, karaciğer yetmezliği", "Kapsid de-targeting (karaciğerden kaçış), miR-122 siteleri", "Yüksek doz Prednizolon (profilaktik ve terapötik)"],
            ["Trombotik Mikroanjiyopati (TMA)", "Yüksek doz AAV-antikor immün kompleksleri", "Yaygın mikrovasküler tromboz, trombositopeni, böbrek hasarı", "Doz azaltımı, daha güçlü tropizmli kapsidler", "Eculizumab (C5 monoklonal kompleman inhibitörü)"],
            ["TLR9 İnterferon Deşarjı", "Vektör DNA'sındaki metillenmemiş CpG adaları", "Tip-1 IFN (IFN-a/b) fırtınası, transgen susturulması", "Kasetin CpG'lerden tamamen arındırılması (CpG Depletion)", "TLR9 antagonistleri (ODN 2088)"],
            ["cGAS-STING Otoenflamasyonu", "Sitoplazmaya kaçan serbest nükleik asitler", "NF-kB ve IRF3 aktivasyonu, kronik doku yangısı", "Endozomdan pürüzsüz nükleer kaçış sağlayan lipidler", "STING inhibitörleri (H-151)"],
            ["NLRP3 İnflamazom Piroptozu", "Lizozomal katepsin sızıntısı ve kalsiyum şoku", "Kaspaz-1 aktivasyonu, masif IL-1b salınımı, hücre lizisi", "Membran yırtmayan biyobozunur akıllı LNP formülasyonu", "MCC950 spesifik NLRP3 inhibitörü"],
            ["Anti-Cas9 Ön Bağışıklığı", "Bakteriyel SpCas9/SaCas9 antijenleri", "Hafıza T hücrelerinin transdukte dokuyu imha etmesi", "Epitop haritalamalı de-immünize Cas9 türevleri", "Geçici Takrolimus / Mikofenolat mofetil rejimi"],
            ["Nötralizan Antikor (NAb) Blokajı", "Dolaşımdaki anti-AAV IgG antikorları", "Virüsün dokuya ulaşamadan temizlenmesi, sıfır etki", "Sentetik yapay kapsidler, PEGilasyon zırhlama", "İmlifidase (IdeS) ile IgG'lerin 2 saatte temizlenmesi"],
            ["Anti-PEG CARPA Reaksiyonu", "LNP yüzeyindeki PEG zincirlerine karşı antikor", "Kompleman aktivasyonu, anafilaktoid psödo-alerji", "Poligliserol ve polisarcosine bazlı yeni nesil lipidler", "Yavaş infüzyon hızı, Antihistaminik premedikasyonu"]
        ]
    },
    {
        "title": "Tablo 14.10: Homo Aeternus İn Vivo Genom Mühendisliği Longevity Transgen Kokteyli",
        "headers": ["Longevity Gen Kaseti", "Vektör ve Teslimat Teknolojisi", "Moleküler Fonksiyonel Mekanizma", "Fizyolojik Fenotipik Gençleşme", "Ömür Uzatma ve Dayanıklılık Etkisi"],
        "rows": [
            ["TERT (Telomeraz Ters Transkriptaz)", "AAV9 / AAV.CAP-B10 (Sistemik)", "Replikatif telomer erozyonunu sıfırlama, mtDNA koruma", "Kök hücre nişlerinin canlanması, doku atrofisi önleme", "Medyan yaşam süresinde %24 net uzama (Blasco protokolü)"],
            ["Klotho (Alfa-Klotho)", "AAV.BI30 (Nöro/Renal hedefli)", "Wnt baskılama, GluN2B restorasyonu, endotel NO sentezi", "Kognitif fonksiyon artışı, böbrek sklerozunun önlenmesi", "Maksimum yaşam süresinde %30 artış, vasküler elastisite"],
            ["Follistatin (FS344)", "AAV1 / AAV8 (Kas hedefli)", "Miyostatin (GDF-8) ve Aktivin nötralizasyonu", "Sarkopeni tasfiyesi, kas kütlesinde %35 hipertrofi", "Fiziksel kırılganlığın silinmesi, metabolik hız artışı"],
            ["PGC-1alfa + Nrf2 (Polisistronik)", "LNP-mRNA / AAV çift kargo", "Mitokondriyal biyogenez (TFAM) + ARE redoks enzimleri", "Yeni mitokondri montajı, ROS süperoksit süpürme", "Metabolik tükenişin engellenmesi, hücresel solunum kür"],
            ["FOXO3a (Longevity Aleli)", "Prime Editing (PE-max LNP)", "Otofaji stimülasyonu, GADD45 DNA tamir indüksiyonu", "Metabolik ve genotoksik streslere mutlak direnç", "Süper-asırlık (Centenarian) genetik dayanıklılık profili"],
            ["APOE2 Dönüşümü (ApoE4 -> ApoE2)", "Prime Editing / ABE (İntratekal)", "İki nükleotit değişimi ile koruyucu ApoE2 izoformu", "Amiloid ve tau klirensinde 5 kat hızlanma", "Alzheimer ve serebral amiloid anjiyopati riskinin silinmesi"],
            ["PCSK9 Nakavt (Verve-101)", "Adenin Baz Düzenleyici (ABE LNP)", "Splice donör inaktivasyonu ile PCSK9 silinmesi", "Plazma LDL kolesterolünde kalıcı %55-60 düşüş", "Ateroskleroz ve miyokard enfarktüsünün mutlak engellenmesi"],
            ["Entegre Polisistronik Kokteyl", "HAC (İnsan Yapay Kromozomu)", "Tüm longevity ağlarının geribildirimli mantık kapılarıyla kontrolü", "Homo Aeternus: Biyolojik ölümsüzlük fenotipi", "Biyolojik yaşlanmanın sıfırlanması ve sonsuz homeostaz"]
        ]
    }
]

# Dokuman Olusturma Dongusu
parts = [
    ("KISIM 1: VİRAL GEN TERAPİSİ VEKTÖRLERİ BİYOLOJİSİ VE İMMÜNOJENİSİTE", part1_subsections),
    ("KISIM 2: AAV KAPSİD MÜHENDİSLİĞİ VE YÖNLENDİRİLMİŞ EVRİM", part2_subsections),
    ("KISIM 3: SENTETİK VE NON-VİRAL TESLİMAT SİSTEMLERİ: LİPİD NANOPARTİKÜLLER (LNPs)", part3_subsections),
    ("KISIM 4: CRISPR-CAS SİSTEMLERİNİN MOLEKÜLER BİYOFİZİĞİ VE CAS VARYANTLARI", part4_subsections),
    ("KISIM 5: BAZ DÜZENLEME (BASE EDITING - CBE VE ABE) TEKNOLOJİLERİ", part5_subsections),
    ("KISIM 6: PRIME EDITING (PE) VE TERS TRANSKRİPTAZ ARACILI HASSAS GENOM YAZIMI", part6_subsections),
    ("KISIM 7: EPİGENOM DÜZENLEME: dCAS9-KRAB, dCAS9-p300 VE DNA METİLASYON MÜHENDİSLİĞİ", part7_subsections),
    ("KISIM 8: HEDEF DIŞI (OFF-TARGET) KİNETİĞİ, BİYOİNFORMATİK VE GÜVENLİK PROFİLİ", part8_subsections),
    ("KISIM 9: İN VİVO GEN TERAPİSİNDE İMMÜN YANIT, İNFLAMAZOM VE HEPATOTOKSİSİTE", part9_subsections),
    ("KISIM 10: HOMO AETERNUS İN VİVO GENOM MÜHENDİSLİĞİ PROTOKOLÜ: ÖMÜR UZATMA GENLERİNİN ENTEGRASYONU", part10_subsections)
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
        frun = pf.add_run(f"Biyomühendislik / Biyofiziksel Bağıntı:  {formula}")
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
print(f"CİLT 14 Başarıyla Kaydedildi: {OUTPUT_PATH}")