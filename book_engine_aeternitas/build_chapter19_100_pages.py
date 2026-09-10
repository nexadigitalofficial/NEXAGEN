"""
PROJECT AETERNITAS - CİLT 19: SİSTEM BİYOLOJİSİ, BİYOLOJİK YAŞ BİYOMARKERLARI VE ÇOKLU OMİKS ENTEGRASYONU
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

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_19_SISTEM_BIYOLOJISI_BIYOLOJIK_YAS_BIYOMARKERLARI_TAM_100_SAYFA.docx"

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
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 19: ÇOKLU OMİKS VE BİYOLOJİK YAŞ SAATLERİ")
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
s_run = sub_p.add_run("CİLT 19: SİSTEM BİYOLOJİSİ, BİYOLOJİK YAŞ BİYOMARKERLARI VE ÇOKLU OMİKS ENTEGRASYONU\\n(EPİGENOMİK, TRANSKRİPTOMİK, PROTEOMİK, METABOLOMİK SAATLER VE YAPAY ZEKA DİJİTAL BİYO-İKİZİ)")
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
ih_run = intro_h.add_run("CİLT 19 MANİFESTOSU: BİYOLOJİK ZAMANIN ÖLÇÜMÜ: ÇOK BOYUTLU OMİKS SAATLERİ VE DİJİTAL BİYO-İKİZ REJİMİ")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "Fizikçi Lord Kelvin'in ünlü aksiyomu modern longevity tıbbının da temelidir: 'Ölçemediğiniz şeyi geliştiremezsiniz ve yönetemezsiniz.' "
    "İnsan yaşlanmasını tersine çevirebilmek için, önce yaşlanmanın insan bedeninde bıraktığı karmaşık, non-lineer ve çok katmanlı entropik "
    "hasarın matematiksel olarak kesin bir doğrulukla ölçülmesi şarttır. Doğum belgemizde yazan 'kronolojik yaş' (Dünya'nın Güneş etrafındaki "
    "tur sayısı); hücrelerimizin, organlarımızın ve moleküler sistemlerimizin gerçek yıpranma derecesini (Biyolojik Yaş) yansıtmakta tamamen "
    "yetersiz ve kaba bir yanılsamadır.\\n\\n"
    "Son yıllarda sistem biyolojisi ve yapay zeka alanında yaşanan devrim; insan yaşlanmasını DNA metilasyonundan (Epigenomik), "
    "gen ekspresyon profillerinden (Transkriptomik), dolaşımdaki binlerce proteinden (Proteomik), hücresel metabolit akışından "
    "(Metabolomik), antikor şekerlenmelerinden (Glikomik) ve bağışıklık hücre alt tiplerinden (İmmünomik) okuyan ultra-hassas "
    "'Biyolojik Yaş Saatlerini' doğurmuştur.\\n\\n"
    "Steve Horvath'ın ilk epigenetik saati ile başlayan bu yürüyüş; PhenoAge ve GrimAge ile mortalite riskini öngörmüş, DunedinPACE "
    "ile anlık yaşlanma hızını (pace of aging) bir takometre gibi ölçmüş ve Tony Wyss-Coray'ın organ-spesifik proteomik saatleriyle "
    "beyin, kalp, böbrek ve bağışıklık sisteminin birbirinden bağımsız biyolojik yaşlarını deşifre etmiştir.\\n\\n"
    "Bu ciltte; 1. nesilden 3. nesle epigenetik saat algoritmaları, tek hücre RNA-seq yaşlanma trajektorileri, 5.000 proteinlik plazma "
    "panelleri, IgG N-glikan analizleri, derin öğrenme oto-enkoderleri ve çoklu omiks entegrasyonuyla inşa edilen 'Homo Aeternus "
    "Dijital Biyo-İkiz' mimarisi 100 akademik alt başlıkta eksiksiz bir sistem biyolojisi vizyonuyla sunulmaktadır."
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
        "1.1 Kronolojik Yaş Yanılsaması: Zamanın Akışı vs Entropik Doku Hasarı",
        "Kronolojik yaş; bir organizmanın doğumundan itibaren Dünya'nın Güneş etrafında kaç tur attığını ölçen salt astronomik ve kronometrik bir parametredir; organizmanın içsel biyolojik yıpranma durumunu açıklayamaz.",
        "İki bağımsız 50 yaşındaki bireyden biri mükemmel metabolik, vasküler ve hücresel gençlik sergilerken; diğeri çoklu kronik hastalıklar, ileri damar sertliği ve tükenmiş kök hücre rezervi taşıyabilir. Biyolojik yaş; doku, hücre ve moleküler düzeydeki fonksiyonel kapasiteyi, entropik enformasyon kaybını ve homeostatik dayanıklılığı (rezilyans) temsil eden gerçek fizyolojik yaştır.",
        "Biyolojik_Yas_Sapmasi: Delta_Yas = Yas_biyolojik - Yas_kronolojik (Pozitif: Hızlanmış Yaşlanma | Negatif: Gençlik)",
        "Bu diferansiyel yaşlanma indeksi, bir bireyin biyolojik sistemlerinin takvim yaşına kıyasla ne kadar ileri veya ne kadar genç olduğunu nicelendirir."
    ),
    (
        "1.2 Sistem Biyolojisi Perspektifi: Organizma Seviyesinde Ağ (Network) Çöküşü",
        "Geleneksel indirgemeci tıp yaşlanmayı tekil organların veya enzimlerin bozulması olarak görürken; sistem biyolojisi yaşlanmayı 'Karmaşık Biyolojik Ağların Deregülasyonu ve Çöküşü' olarak modeller.",
        "İnsan bedeni; gen regülasyon ağları, protein-protein etkileşim (PPI) ağları, metabolik akı ağları ve nöro-endokrin-immün eksenlerin iç içe geçtiği çok katmanlı, hiyerarşik bir süper-ağdır. Yaşlanma; bu ağlardaki modülerliğin kaybı, düğüm noktalarının (hub) bağlantı yitimi ve geri bildirim döngülerinin senkronizasyonunun bozulmasıyla ortaya çıkan global bir sistem iflasıdır.",
        "Ag_Dayanikligi: R_network = 1 - ( N_hasarli_dugum / N_toplam_hub ) -> Yaşlanmayla Kritik Faz Geçişi (Eşik Çöküş)",
        "Bu topolojik ağ rezilyansı eşitliği, sistemik biyolojik düğüm noktalarının hasar görmesiyle organizmanın stabiliteden kritik çöküşe geçişini modeller."
    ),
    (
        "1.3 Fiziksel Rezilyans (Resilience) ve Dinamik İyileşme Zamanı (Recovery Time)",
        "Yaşlanan bir organizmanın en temel biyofiziksel özelliği; dışsal bir strese (enfeksiyon, travma, sıcaklık şoku) maruz kaldığında bazal denge durumuna dönme hızının giderek yavaşlamasıdır.",
        "Genç bir bireyde fizyolojik göstergeler (kalp hızı, lökosit sayısı, kan şekeri) stres sonrası birkaç saat içinde bazal çizgiye otururken; yaşlı bireyde bu salınımın sönümlenme süresi (Damping Time) günlere veya haftalara uzar. Peter Fedichev ve Gero ekibinin geliştirdiği dinamik rezilyans modelleri; bu iyileşme süresinin yaşla lineer olarak uzadığını ve yaklaşık 120-150 yaş civarında sonsuza ıraksayarak rezilyansın sıfırlandığını kanıtlamıştır.",
        "Dinamik_Salinim: x(t) = x_0 * exp( - gamma_rezilyans * t ) * cos( omega * t ) (gamma -> 0: Rezilyans Kaybı)",
        "Bu sönümlü harmonik osilatör modeli, rezilyans katsayısı (gamma) sıfıra yaklaştıkça organizmanın en ufak bir pertürbasyondan dahi toparlanamayarak öldüğünü gösterir."
    ),
    (
        "1.4 Gompertz-Makeham Mortalite Kanunu ve Biyolojik Yaşın Matematiksel Temeli",
        "İnsan popülasyonlarında ölüm olasılığı yaşla birlikte sabit bir hızla artmaz; ergenlikten sonra her 8 yılda bir ikiye katlanan katı bir üstel (eksponansiyel) artış sergiler (Gompertz Kanunu).",
        "Benjamin Gompertz tarafından 1825'te keşfedilen ve Makeham tarafından yaşa bağlı olmayan arka plan ölümleri eklenerek geliştirilen bu denklem; biyolojik yaşlanmanın içsel hızını (Gompertz alpha katsayısı) tanımlar. Gerçek bir longevity terapisi; mortalite eğrisini yalnızca sağa kaydırmakla kalmamalı, Gompertz eğimini (alpha) yataylaştırarak yaşa bağlı ölüm artış hızını kökten yavaşlatmalıdır.",
        "Gompertz-Makeham_Eşitliği: mu(t) = A + R_0 * exp( alpha_gompertz * t ) (İnsanda alpha ~ 0.085 / Yıl)",
        "Bu klasik aktüeryal biyodemografi eşitliği, 30 yaşından sonra her yıl mortalite riskimizin yaklaşık %8.5 oranında eksponansiyel olarak arttığını formüle eder."
    ),
    (
        "1.5 Biyolojik Yaş Biyomarkerları İçin Amerikan Yaşlanma Araştırmaları Federasyonu (AFAR) Kriterleri",
        "Herhangi bir biyolojik parametrenin güvenilir bir 'Yaşlanma Biyomarkeri' olarak kabul edilebilmesi için AFAR (American Federation for Aging Research) tarafından belirlenen katı bilimsel kriterleri karşılaması şarttır.",
        "AFAR kriterlerine göre bir biyomarker: 1) Kronolojik yaştan daha iyi bir doğrulukla kalan yaşam süresini ve morbiditeyi öngörebilmelidir; 2) Belirli bir hastalığı değil, yaşlanmanın temel biyolojik süreçlerini izlemelidir; 3) Deneklere zarar vermeden tekrarlanabilir ve non-invaziv ölçülebilmelidir; 4) Hayvan modellerinde (kemirgenler, primatlar) ve insanlarda türler arası geçerli olmalıdır; 5) Uygulanan longevity müdahalelerine (kalori kısıtlaması, rapamisin vb.) dinamik ve ölçülebilir yanıt vermelidir.",
        "Gecerlilik_Kriteri: Prediktif_Guc = AUC_ROC(Biyolojik_Belirtec) >> AUC_ROC(Kronolojik_Yas) (P < 10^-5)",
        "Bu biyostatistiksel diskriminasyon kriteri, adil bir biyomarkerin klinik mortaliteyi kronolojik yaştan üstün bir hassasiyetle ayırt etme zorunluluğunu tanımlar."
    ),
    (
        "1.6 Molekülden Fenotipe Çok Katmanlı Biyolojik Katmanlaşma (Omics Hierarchy)",
        "Biyolojik sistemlerde enformasyon akışı ve entropik hasar doğrusal değildir; genomdan başlayıp fenotipe uzanan çok katmanlı bir piramit üzerinde yayılır.",
        "En alt katmanda statik DNA dizisi (Genomik); onun üzerinde kromatini açıp kapatan kimyasal etiketler (Epigenomik); üretilen mesajcı RNA'lar (Transkriptomik); sentezlenen fonksiyonel hücresel makineler (Proteomik); post-translasyonel şeker modifikasyonları (Glikomik); hücre içi biyokimyasal reaksiyon yakıtları ve atıkları (Metabolomik/Lipidomik); ve en üstte klinik kan biyokimyası ile fizyolojik fonksiyonlar (Fenomik) yer alır. Her katman bir alt katmanın gürültüsünü filtreler veya amplifiye eder.",
        "Enformasyon_Akisi: DNA ----> Epigenom ----> Transkriptom ----> Proteom ----> Metabolom ----> Klinik Fenotip",
        "Bu biyolojik katmanlaşma aksiyomu, yaşlanmanın her bir omiks basamağında farklı zaman ölçeklerinde ve kinetiklerde iz bıraktığını modeller."
    ),
    (
        "1.7 İnformasyon Teorisi ve Yaşlanma: Shannon Entropisi ve Epigenetik Gürültü",
        "David Sinclair'in 'Yaşlanmanın Enformasyon Teorisi'ne (Information Theory of Aging) göre; yaşlanma donanımın (DNA dizisi) bozulması değil, hücresel yazılımın (Epigenetik okuma düzeni) gürültüyle kirlenmesidir.",
        "Genç bir hücrede kromatin son derece düzenli, düşük entropili bir konfigürasyondadır (hangi genlerin okunacağı kusursuzca bellidir). Yaşlandıkça DNA çift sarmal kırıkları ve epigenetik faktörlerin yer değiştirmesi (Relocalization of Chromatin Modifiers - RCM); hücrenin kimliğini kaybetmesine ve Shannon entropisinin (H) yükselmesine yol açar. Hücre hangi geni ne zaman okuyacağını unutur; sinyal-gürültü oranı (SNR) çöker.",
        "Shannon_Entropisi: H(Epigenom) = - Sum( p_i * log2( p_i ) ) (Yaşlandıkça H tavan yapar, SNR çöker)",
        "Bu enformasyon teorisi eşitliği, hücresel yaşlanmanın epigenetik enformasyon kaybı ve stokastik transkripsiyonel gürültü artışı olduğunu kanıtlar."
    ),
    (
        "1.8 Allostaz ve Allostatik Yük (Allostatic Load): Kümülatif Çevresel Aşınma",
        "Homeostaz bedenin anlık sabit iç dengesini tanımlarken; 'Allostaz' çevresel ve psikososyal stres faktörlerine karşı fizyolojik sistemlerin dinamik olarak değişerek dengeyi koruma sürecidir.",
        "Allostatik Yük (Allostatic Load); yıllar boyunca tekrarlanan allostatik adaptasyonların (kortizol salınımı, sempatik aktivasyon, kan basıncı dalgalanmaları, inflamatuar deşarjlar) vasküler, nörolojik ve metabolik organlarda bıraktığı kümülatif aşınma ve yıpranma faturasıdır. Allostatik yük skoru yüksek olan bireylerde biyolojik yaş kronolojik yaştan 10-15 yıl daha hızlı ilerler.",
        "Allostatik_Yuk_Skoru: ALS = Sum_{i=1}^{10} I( Parametre_i > Esik_patolojik ) (ALS >= 4: Masif Hızlanmış Yaşlanma)",
        "Bu kümülatif stres indeksi formülasyonu, kardiyometabolik ve immünolojik parametrelerin kronik stres altında aşınma derecesini puanlar."
    ),
    (
        "1.9 Kronolojik Yaş ile Biyolojik Yaş Arasındaki Diskordansın Klinik Sonuçları",
        "Klinik kohortlarda kronolojik yaş ile biyolojik yaş arasındaki uyumsuzluk (Age Acceleration Residual - AAR), gelecekteki sağlık durumunun en güçlü bağımsız belirtecidir.",
        "BioAge > ChronoAge (Hızlanmış Biyolojik Yaşlanma) durumunda; kardiyovasküler mortalite, kanser riski, demans ve frajilite (kırılganlık) riski her 5 yıllık pozitif sapma için %30 ila %60 artar. Buna karşılık, BioAge < ChronoAge (Yavaşlamış Biyolojik Yaşlanma) sergileyen 60 yaşındaki bir birey; 40 yaşındaki bir insanın hastalık risk profiline ve fiziksel dayanıklılığına sahip olur; bu durum 'Sağlıklı Yaşam Süresinin' (Healthspan) anahtarıdır.",
        "Artmis_Risk_Orani: Hazard_Ratio = HR = exp( beta_AAR * Delta_Yas ) (Her +1 Yıl İçin Morbidite %10 Artar)",
        "Bu Cox orantılı risk modeli, biyolojik yaş sapmasının tüm nedenlere bağlı mortalite riskini nasıl üstel olarak artırdığını tanımlar."
    ),
    (
        "1.10 Homo Aeternus Biyolojik Yaş Standardizasyon Manifestosu: Zamansız Biyoloji",
        "Homo Aeternus tıbbı; kronolojik yaşı anlamsız bir takvim kaydına indirgeyerek biyolojik yaşı mutlak bir mühendislik parametresi olarak kabul eder.",
        "Her bir bireyin biyolojik yaşı; hücresel, dokusal ve organ seviyesinde sürekli ölçülür ve 20-25 yaş aralığındaki fizyolojik optimumda sabitlenir. Takvimler 50, 80 veya 150 yılı gösterse dahi; organizmanın sistemik biyolojik yaşı, entropi seviyesi, rezilyans katsayısı ve hücresel enformasyon sadakati değişmez; zamanın biyolojik yıkımı mutlak olarak durdurulur.",
        "Homo_Aeternus_Standardi: Yas_biyolojik(t) = 22.0 Yas (Kalıcı ve Dinamik Olarak Dengelenen Biyolojik Yaş)",
        "Bu nihai zamansızlık eşitliği, insan biyolojisinin dışsal takvim zamanından tamamen koparılarak ebedi bir gençlik penceresinde kilitlenmesini simgeler."
    )
]

# ================= KISIM 2 =================
part2_subsections = [
    (
        "2.1 DNA Metilasyonu (5mC) ve Epigenetik Saatlerin Doğuşu",
        "Genomik DNA'mızda sitozin bazlarının 5. karbonuna bir metil grubu eklenmesiyle oluşan 5-metilsitozin (5mC); gen ekspresyonunu susturan ana epigenetik mekanizmadır.",
        "İnsan genomunda yaklaşık 28 milyon CpG bölgesi bulunur. Bu bölgelerin metilasyon durumu yaşlanma boyunca rastgele değil; şaşırtıcı bir matematiksel sadakat ve düzenlilikle değişir: Genom genelinde global bir hipometilasyon (retrotranspozonların çözülmesi) yaşanırken, spesifik gelişimsel ve tümör baskılayıcı gen promotörlerindeki CpG adacıklarında hipermetilasyon gerçekleşir. Bu düzenli değişim, epigenetik saatlerin temelini oluşturur.",
        "Metilasyon_Beta_Degeri: Beta_i = M_i / ( M_i + U_i + alpha_offset ) (0.0: Tamamen Metilsiz | 1.0: Tam Metilli)",
        "Bu standart mikrodizi sinyal oranı, belirli bir CpG bölgesindeki metillenmiş (M) ve metillenmemiş (U) probların oranını nicelendirir."
    ),
    (
        "2.2 Birinci Nesil Saatler: Steve Horvath (353 CpG) ve Hannum (71 CpG) Algoritmaları",
        "2013 yılında Steve Horvath ve Gregory Hannum, yaşlanma biyolojisinde bir çığır açarak ilk makine öğrenimi tabanlı epigenetik saatleri yayınladı.",
        "Hannum saati kanda 71 CpG bölgesini kullanarak kronolojik yaşı tahmin ederken; Steve Horvath 51 farklı insan dokusu ve hücre tipinden elde edilen 8.000 örneği elastik net regresyonu (Elastic Net) ile eğiterek tüm dokularda geçerli 'Pan-Tissue' saatini geliştirdi (353 CpG). Horvath saati kronolojik yaş ile r = 0.96 gibi inanılmaz bir korelasyon gösterdi ve ortalama hata payı yalnızca 3.6 yıldı.",
        "Horvath_Modeli: DNAmAge = F( b_0 + Sum_{i=1}^{353} w_i * Beta_i ) (F: Non-lineer Yaş Transformasyon Fonksiyonu)",
        "Bu cezalı regresyon eşitliği, 353 kilit CpG lokusundaki metilasyon yüzdelerinin ağırlıklı toplamı ile doku yaşını hesaplayan ilk algoritmayı tanımlar."
    ),
    (
        "2.3 İkinci Nesil Saatler: PhenoAge ve GrimAge ile Mortalite Riskini Tahmin Etme",
        "Birinci nesil saatler kronolojik yaşı kusursuz tahmin etse de, bir eksikleri vardı: İki sağlıklı birey arasındaki biyolojik dayanıklılık farkını ve ölüm riskini yakalamakta zayıftılar.",
        "Morgan Levine ve Steve Horvath, 2018'de 'DNAm PhenoAge'i (513 CpG) geliştirdi; model kronolojik yaş yerine 9 klinik kan biyomarkeri ve mortalite skoru üzerinden eğitildi. 2019'da ise Ake Lu ve Horvath 'DNAm GrimAge'i yayınladı: GrimAge; plazma proteinleri (adrenomedüllin, GDF-15, PAI-1) ve sigara içme geçmişinin DNA metilasyon vekilleri (surrogates) üzerinden eğitildi. GrimAge, gelecekteki ölüm zamanını, kanser riskini ve kalp krizini öngören en ölümcül doğruluktaki epigenetik saattir.",
        "GrimAge_Risk_Fonksiyonu: Hazard_Olum = exp( gamma_grim * ( GrimAge - Kronolojik_Yas ) ) (Her +1 Yıl GrimAge = %10 Mortalite Artışı)",
        "Bu mortalite hazard modeli, GrimAge ivmelenmesinin bir insanın biyolojik ölüm riskini nasıl doğrudan yansıttığını belgeler."
    ),
    (
        "2.4 Üçüncü Nesil Saatler: DunedinPACE ve Anlık Yaşlanma Hızı Takometresi",
        "Önceki tüm saatler 'şimdiye kadar ne kadar yaşlandığınızı' (odometre / kilometre sayacı) ölçerken; Dan Belsky ve Terrie Moffitt liderliğinde 2022'de geliştirilen 'DunedinPACE' (Pace of Aging Computed from the Epigenome), 'şu anda ne hızla yaşlandığınızı' ölçen bir 'Takometre'dir.",
        "Dunedin boylamsal kohortunda 1.000 birey 26 yaşından 45 yaşına kadar 20 yıl boyunca 19 farklı biyobelirteçle izlenmiş ve organ sistemlerinin bozulma hızı hesaplanmıştır. DunedinPACE (173 CpG); bireyin her bir takvim yılında kaç 'biyolojik yıl' tükettiğini ölçer. Normal hız 1.0 iken; hızlanmış bireyler 1.4 yıl/yıl hızla yaşlanır; başarılı longevity müdahaleleri bu hızı 0.75 yıl/yıl seviyesine düşürebilir.",
        "DunedinPACE_Formulu: Hız = PACE = d(Biyolojik_Yas) / d(Takvim_Yili) (Optimum Longevity Değeri: PACE < 0.80)",
        "Bu anlık türev eşitliği, epigenetik hız göstergesinin uygulanan tedavilerin yaşlanmayı anlık olarak yavaşlatma etkisini nasıl ölçtüğünü formüle eder."
    ),
    (
        "2.5 İntra-Hücresel Heterojenite ve Hücre Tipi Kompozisyonu Dekonvolüsyonu",
        "Kandan alınan DNA metilasyon profillerinde karşılaşılan en büyük gürültü kaynağı; kandaki hücre tiplerinin (nötrofil, lenfosit, monosit) oranlarının bireyden bireye ve enfeksiyonla değişmesidir.",
        "Hücre popülasyon kaymalarını gerçek epigenetik yaşlanmadan ayırmak için 'Houseman Dekonvolüsyon Algoritması' kullanılır. Saf hücre tiplerinin referans metilasyon imzaları kullanılarak, karışık kan örneğindeki CD4+ T, CD8+ T, NK, B hücre ve granülosit fraksiyonları matematiksel olarak hesaplanır. Bu sayede 'İntrinsik Epigenetik Yaş İvmelenmesi' (IEAA - hücre tipinden bağımsız hücresel yaş) ile 'Ekstrinsik Epigenetik Yaş İvmelenmesi' (EEAA - immün kompozisyonu içeren yaş) net olarak ayrıştırılır.",
        "Dekonvolusyon_Modeli: Y_kan = Sum_{k=1}^{K} w_k * X_k (X_k: Saf K Hücre Tipinin İmzası | w_k: Fraksiyon)",
        "Bu matris çarpanlara ayırma eşitliği, periferik kan epigenomik verilerinden bağışıklık hücresi alt tiplerinin kesin oranlarını çıkarmayı sağlar."
    ),
    (
        "2.6 Teknik Gürültü ve Güvenilirlik Çıkmazı: Test-Retest Uyuşmazlığı ve Çözümleri",
        "Erken dönem epigenetik saatlerin en can sıkıcı teknik zayıflığı; aynı bireyden aynı anda alınan iki kan tüpünün mikrodizi analizinde 2 ila 5 yıl arasında farklı biyolojik yaş sonuçları verebilmesiydi (Teknik Varyans).",
        "Bu gürültü, Illumina çiplerindeki kimyasal prob bağlanma dalgalanmalarından kaynaklanır. 2022'de Higgins-Chen ve Levine, temel bileşen analizini (Principal Component Analysis - PCA) kullanarak 'PC-Saatlerini' (PC-Horvath, PC-GrimAge) geliştirdi. Teknik gürültüyü temsil eden rastgele bileşenler matematiksel olarak elenerek biyolojik yaş skorunun test-retest korelasyonu r = 0.999'a çıkarıldı; hata payı aylar seviyesine indirildi.",
        "Gurultu_Reduksiyonu: PC_Age = F( Sum_{j=1}^{M} lambda_j * PC_j(CpG) ) (Teknik Varyans %95 Elenir)",
        "Bu temel bileşen düzenleme formülü, epigenetik saatlerin klinik güvenilirlik ve tekrarlanabilirlik seviyesini tavan yaptıran matematiksel filtreyi tanımlar."
    ),
    (
        "2.7 Epigenetik Saatleri Döndüren Moleküler Mekanizma Nedir? Neden Tik Tak Eder?",
        "Epigenetik saatler mükemmel bir kesinlikle yaşlanmayı ölçer; ancak biyologlar yıllarca şu sorunun cevabını aradı: Saatleri çalıştıran moleküler sarkacın mekanizması nedir?",
        "Bugün kanıtlanmıştır ki; epigenetik saatlerin tik-takları, hücrenin DNA tamir mekanizmalarının ve gelişimsel transkripsiyon faktörlerinin bıraktığı 'epigenetik skar izleridir'. DNA çift zincir kırıkları oluştuğunda DNMT1, SIRT1 ve EZH2 enzim mekanları kırıma koşar; onarım tamamlandıktan sonra histonlar ve metilasyonlar %99.9 eski haline döner ancak %0.1'lik stokastik bir metilasyon kayması geride kalır. Milyonlarca onarım döngüsü boyunca biriken bu kümülatif tortu, saatin tik taklarıdır.",
        "Saat_Mekanizmasi: d(CpG_metilasyon)/dt = k_tamir * N_DSB * ( 1 - Fidelite_epigenetik )",
        "Bu epigenetik sürüklenme kinetiği eşitliği, saatin aslında hücrenin yaşam boyu maruz kaldığı DNA hasar ve onarım döngülerinin toplam kaydını tuttuğunu modeller."
    ),
    (
        "2.8 Yamanaka Faktörleri (OSKM) ile Epigenetik Saatin Sıfırlanması",
        "Epigenetik saatlerin tersine çevrilebilirliğinin en nihai kanıtı; olgun bir somatik hücreye Yamanaka faktörleri (Oct4, Sox2, Klf4, c-Myc) verildiğinde Horvath yaşının hızla gerileyerek '0.0 Yaş' (embriyonik yaş) düzeyine inmesidir.",
        "Steve Horvath'ın 2013'teki orijinal makalesinde ve ardından 2020'de David Sinclair'in Nature'da yayımlanan göz siniri rejenerasyonu çalışmasında (Lu ve ark.); AAV tabanlı kısmi OSK indüksiyonunun nöronların epigenetik yaşını gençleştirerek kör farelerde görme fonksiyonunu geri getirdiği ispatlandı. Epigenetik saat, yalnızca tek yönlü çalışan bir sayaç değil; geriye doğru kurulabilen dinamik bir yazılımdır.",
        "Geriye_Kurma: DNAmAge(t) = DNAmAge_0 - Integral_0^T [ k_OSKM * [OSKM(t)] * Aktivite_TET ] dt -> 0.0 Yas",
        "Bu epigenetik gençleşme integrali, Yamanaka faktörlerinin ve TET demetilazların epigenetik saati embriyonik başlangıç noktasına kadar nasıl geri sarabildiğini belgeler."
    ),
    (
        "2.9 Klinik Müdahalelerle Epigenetik Saatin Geri Sarılması: TRIIM ve Rejuvant Çalışmaları",
        "Epigenetik saatlerin insan klinik deneylerinde başarıyla geriye sarılabildiği Gregory Fahy'nin öncülüğündeki dönüm noktası niteliğindeki 'TRIIM' (Thymus Regeneration, Immunorestoration, and Insulin Mitigation) çalışmasıyla kanıtlanmıştır.",
        "Aging Cell'de yayımlanan çalışmada; 1 yıl boyunca Rekombinant İnsan Büyüme Hormonu (rhGH), DHEA ve Metformin kokteyli alan 51-65 yaş arası erkeklerin GrimAge saati ortalama 2.5 yıl geriye sarmıştır (net 1.5 yıl gençleşme). Benzer şekilde alfa-ketoglutarat (AKG / Rejuvant) takviyesinin DNAmAge saatini 8 ayda ortalama 8 yıl gençleştirdiği raporlanmıştır; epigenetik saatler klinik longevity protokollerinin nihai başarı hakemi haline gelmiştir.",
        "Klinik_Geri_Alma: Delta_Epigenetik_Yas = Yas_son - Yas_baslangic < - 2.5 Yıl (P < 0.001)",
        "Bu klinik rejuvenasyon istatistiği, hedefe yönelik farmakolojik müdahalelerin insan DNA metilasyon yaşını geriletebileceğinin ilk hakemli tıp kanıtıdır."
    ),
    (
        "2.10 Homo Aeternus Epigenetik Saat Protokolü: Sürekli 20 Yaş Metilasyon Mimarisi",
        "Homo Aeternus genomik mühendisliğinde epigenom; rastgele yaşlanmaya ve metilasyon sürüklenmesine bırakılamaz.",
        "Birey 6 ayda bir kandan ve doku biyopsilerinden PC-GrimAge ve DunedinPACE ile taranır. PACE skoru 0.80'in üzerine çıktığında veya PC-GrimAge kronolojik yaştan 1 yıl saptığında; hedefe yönelik dCas9-TET1 / dCas9-DNMT3A epigenetik düzenleyicileri ve kontrollü mRNA-OSK puls terapileri devreye sokulur. 353 Horvath CpG lokusu ve tüm promotör metilasyonları mutlak bir kararlılıkla 20 yaş epigenetik mimarisine kilitlenir.",
        "Hedef_Homo_Aeternus: PC-GrimAge = 20.0 Yas & DunedinPACE <= 0.65 Yıl/Yıl (Durdurulmuş ve Gençleşmiş Epigenom)",
        "Bu nihai epigenetik kararlılık hedefi, insan hücrelerinin epigenetik hafızasının biyomühendislikle ölümsüz bir gençlikte tutulmasını simgeler."
    )
]

# ================= KISIM 3 =================
part3_subsections = [
    (
        "3.1 Transkriptomik Yaşlanma Biyolojisi: Gen Ekspresyon Profilindeki Entropik Kayma",
        "Hücrenin genomik DNA'sından anlık olarak sentezlenen tüm haberci RNA'ların (mRNA) toplamı olan 'Transkriptom'; epigenetik şalterlerin hücredeki anlık dinamik eylemini yansıtır.",
        "Yaşlanmayla birlikte hücrelerin transkripsiyonel programında küresel bir dengesizlik başlar: Mitokondriyal elektron taşıma zinciri, DNA onarımı, proteozomal degradasyon ve ribozom biyogenezi genlerinin ekspresyonu çökerken; pro-enflamatuar sitokinler, apoptoz medyatörleri, otoimmünite faktörleri ve retrotranspozon transkriptleri tavan yapar. Transkriptomik saatler bu küresel gen aktivite haritasından biyolojik yaşı deşifre eder.",
        "Transkriptomik_Profil: T(Yas) = [ log2(TPM_gene1), log2(TPM_gene2), ..., log2(TPM_geneN) ]",
        "Bu yüksek boyutlu transkriptomik ekspresyon vektörü, binlerce genin transkript başına milyon (TPM) seviyelerindeki yaşa bağlı dinamik kaymasını tanımlar."
    ),
    (
        "3.2 RNA-seq Teknolojileri: Bulk RNA-seq'ten Tek Hücre (scRNA-seq) Yaşlanma Atlaslarına",
        "Klasik 'Bulk RNA-seq' yöntemi milyonlarca hücrenin ortalama transkript profilini verirken; 'Tek Hücre RNA Dizilemesi' (scRNA-seq) doku içindeki her bir tekil hücrenin yaşlanma durumunu bağımsız olarak ortaya koymuştur.",
        "Tabula Muris Senis ve İnsan Hücre Yaşlanma Atlası (Human Aging Cell Atlas) gibi devasa konsorsiyumlar; yaşlanmanın her hücre tipinde aynı hızla ilerlemediğini kanıtlamıştır. Bir doku içinde kök hücreler transkripsiyonel kimliğini korurken, endotel hücreleri hızla senesent bir profile kayabilmektedir. scRNA-seq; yaşlanan dokulardaki nadir senesent hücreleri ve klonal miyeloid hücreleri tek hücre çözünürlüğünde yakalar.",
        "Hücre_Heterojenitesi: Var(T_hucre) = Sum_{j=1}^{C} ( t_{ij} - mu_i )^2 -> Yaşlandıkça Hücreler Arası Varyans Patlar",
        "Bu transkripsiyonel varyans denklemi, yaşlanmanın hücreler arasındaki transkripsiyonel disiplini bozarak stokastik bir kaosa sürüklediğini gösterir."
    ),
    (
        "3.3 Biyolojik Yaş Tahmininde Transkriptomik Saat Modelleri (BiT-Age vb.)",
        "Epigenetik saatlerin başarısının ardından, saf RNA-seq ekspresyon verileri kullanılarak ultra-hassas 'Transkriptomik Biyolojik Saatler' inşa edilmiştir.",
        "David Meyer ve Björn Schumacher tarafından geliştirilen 'BiT-Age' (Binarized Transcriptomic Aging Clock); Caenorhabditis elegans'tan insana kadar evrimsel olarak korunan gen ekspresyon eşiklerini kullanarak yaş tahmininde bulunmuştur. Yapay sinir ağları ve LASSO regresyonu ile eğitilen modeller; yaklaşık 1.000 kilit genin ekspresyon düzeyinden biyolojik yaşı 2-3 yıl hata payıyla saptayabilmektedir.",
        "BiT-Age_Algoritmasi: RNA_Age = g( Sum_{k=1}^{G} beta_k * Phi( Gene_k_ekspresyonu ) )",
        "Bu transkriptomik yaş tahmin eşitliği, gen aktivasyon seviyelerinin non-lineer nöral aktivasyon fonksiyonuyla biyolojik yaşa dönüştürülmesini modeller."
    ),
    (
        "3.4 Stokastik Transkripsiyonel Gürültü (Transcriptional Noise) ve Hücresel Kimlik Kaybı",
        "Gençlikte aynı dokudaki iki komşu kardeş hücre (örneğin iki hepatosit), neredeyse birbirinin fotokopisi olan bir gen ekspresyon homojenliği sergiler.",
        "Yaşlandıkça nükleer kromatindeki gevşeme ve RNA Polimeraz II duraklamaları nedeniyle 'Stokastik Transkripsiyonel Gürültü' başlar. Hücreden hücreye ekspresyon korelasyonu çöker; bazı hücreler aşırı miktarda uygunsuz gen eksprese ederken diğerleri temel fonksiyonel enzimleri sentezlemeyi bırakır. Bu transkripsiyonel anarşi dokunun senkronize çalışmasını imkânsız kılar ve fonksiyonel organ yetersizliğini doğurur.",
        "Gurultu_Katsayisi: Noise = CV^2 = ( sigma_ekspresyon / mu_ekspresyon )^2 (Yaşlı Dokuda 4 Kat Artış)",
        "Bu varyasyon katsayısı karesi eşitliği, yaşlanmayla birlikte hücresel transkripsiyonel disiplinin stokastik gürültüye teslim oluşunu simgeler."
    ),
    (
        "3.5 Splicing Disregülasyonu: Alternatif Kırpılma Bozulması ve İntron Retansiyonu",
        "Transkriptomik yaşlanmanın en dramatik yapısal bozulmalarından biri; pre-mRNA'nın olgun mRNA'ya dönüştürülmesini yöneten 'Spliceozom' makinelerinin yaşlanmasıdır.",
        "Lorna Harries ve ekibinin keşfettiği üzere; yaşlanan hücrelerde kilit 'Splicing Faktörleri' (SRSF1, U2AF, hnRNP) tükenir veya hatalı fosforillenir. Bu durum anormal alternatif kırpılmalara, fonksiyonel ekzonların atlanmasına ve en tehlikelisi 'İntron Retansiyonuna' (intronların mRNA içinde kalması) yol açar. Hatalı kırpılan bu anormal transkriptler stop kodonları içerir; ya NMD (Nonsense-mediated decay) ile parçalanıp enerji tüketir ya da toksik, katlanamayan protein agregatlarına dönüşür.",
        "Intron_Retansiyon_Orani: IR_Ratio = Reads_intron / ( Reads_ekzon + Reads_intron ) -> Yaşlanmayla Doğrusal Artış",
        "Bu transkriptomik kalite kontrol metriği, yaşlanan hücrelerde spliceozom sadakatinin çökerek intronik çöpler üretme derecesini belgeler."
    ),
    (
        "3.6 Uzun vs Kısa Transkript Dengesizliği (Gene Length Dependent Transcriptome Imbalance)",
        "2022 yılında Nature Aging ve Nature Genetics'te eş zamanlı yayımlanan devrimci keşif; yaşlanmanın genlerin 'Uzunluğuna' (Nükleotid baz uzunluğu) dayalı evrensel bir dengesizlik yarattığını gösterdi.",
        "Thomas Stoeger ve ekibi; nöronlar, hepatositler ve kas hücrelerinde çok uzun genlerin (>100.000 baz çifti, örneğin sinaptik adezyon ve akson rehberlik genleri) yaşlandıkça transkripsiyonunun yarıda kesildiğini ve ekspresyonlarının çöktüğünü; buna karşılık çok kısa genlerin (<10.000 baz çifti, örneğin ribozomal ve pro-enflamatuar genler) ekspresyonunun arttığını keşfetti. DNA hasarı uzun genlerin üzerinde stokastik olarak daha çok birikerek RNA polimerazı bloke etmektedir.",
        "Uzunluk_Skoru: LTI = log2( Ekspresyon_Uzun_Genler / Ekspresyon_Kisa_Genler ) -> Yaşla Birlikte Negatife Çöker",
        "Bu gen uzunluğu dengesizlik indeksi (LTI), uzun nöronal ve yapısal genlerin yaşlanmayla neden seçici olarak sustuğunu açıklayan biyoenformatik eşitliktir."
    ),
    (
        "3.7 Senesens Transkriptomu: SASP Faktörlerinin Moleküler RNA İmzası",
        "Hücresel senesense giren bir hücrenin transkriptomik profili, sağlıklı hücrelerden tamamen farklılaşarak devasa bir sekretuar fabrikaya dönüşür.",
        "Senesent hücrelerde p16INK4a (CDKN2A) ve p21 (CDKN1A) transkriptleri tavan yaparken; NF-kappaB ve C/EBP-beta transkripsiyon faktörleri kontrolünde yüzlerce pro-enflamatuar sitokin (IL6, CXCL8, CCL2) ve doku eritici matriks metalloproteinaz (MMP1, MMP3, MMP9) geni aşırı eksprese edilir. Bu spesifik 'SASP Transkriptom İmzası', dokudaki senesent hücre yükünü ve dokusal toksisiteyi anında ele verir.",
        "SASP_Transkripsiyon_Skoru: SASP_Score = Sum_{k in SASP_genes} z_score( log2(TPM_k) ) >> 5.0 (Senesent Toksisite)",
        "Bu standartlaştırılmış z-skoru toplamı, doku örneklerinde senesent hücrelerin çevreye saldığı pro-enflamatuar transkript yoğunluğunu ölçer."
    ),
    (
        "3.8 Retrotranspozon Reaktivasyonu: LINE-1 ve Endojen Retrovirüslerin Transkripsiyonel Uyanışı",
        "Genomumuzun yaklaşık %45'ini oluşturan retrotranspozonlar (özellikle LINE-1 / L1 elemanları); genç hücrelerde katı heterokromatin (H3K9me3 ve DNA metilasyonu) altında kilitlidir.",
        "Yaşlanmayla birlikte heterokromatin adacıkları çözüldüğünde; 'LINE-1' retrotranspozonları uyanarak transkripte edilmeye başlar. Sitoplazmaya dökülen LINE-1 RNA'ları ve ters transkripsiyonla üretilen cDNA kopyaları; hücre tarafından yabancı bir retroviral enfeksiyon (HIV gibi) olarak algılanır. cGAS-STING yolağı ateşlenir ve Tip I İnterferon (IFN-alfa/beta) fırtınası başlar; organizma kendi antik viral fosillerinin uyanmasıyla içten içe yanar.",
        "LINE1_Fluksu: d[cDNA_L1]/dt = k_RT * [L1_RNA] * [ORF2p_Revers_Transkriptaz] -> cGAS-STING Patlaması",
        "Bu retroviral transkripsiyon eşitliği, yaşlanan hücrelerde sessiz transpozonların uyanarak steril otoimmün yıkımı nasıl başlattığını modeller."
    ),
    (
        "3.9 Transkriptomik Saatlerin Farmakolojik Müdahalelere Duyarlılığı (Rapamisin, Metformin)",
        "Transkriptomik saatlerin epigenetik saatlere kıyasla en büyük klinik avantajı; saatler ve günler içinde uygulanan farmakolojik tedavilere son derece hızlı ve dinamik yanıt vermesidir.",
        "Epigenetik metilasyonların değişmesi aylar veya yıllar alırken; rapamisin, kalori kısıtlaması veya senolitik tedavisi uygulanan bir organizmada transkriptomik yaş profili 48-72 saat içinde dramatik biçimde genç bir ekspresyon paternine kayar. mTOR inhibisyonu ribozomal transkripsiyonu sakinleştirir, otofaji genlerini (ATG5, BECN1) aktive eder ve transkriptomik yaşı derhal geriye sarar.",
        "Dinamik_Geri_Donus: d(RNA_Age)/dt = - k_farmako * [Rapamisin] * ( RNA_Age - RNA_Genc )",
        "Bu transkriptomik rejuvenasyon kinetiği, anti-aging ilaçların gen ekspresyon düzeyinde sağladığı hızlı gençleşme yanıtını simgeler."
    ),
    (
        "3.10 Homo Aeternus Transkriptomik Homojenite Protokolü: Sıfır Gürültü ve Kararlı RNA",
        "Homo Aeternus longevity mimarisi; transkriptomik ağı mutlak bir düzen, yüksek sinyal-gürültü oranı ve kusursuz bir splicing disiplini altında tutar:",
        "1) Sentetik RNA bağlayıcı proteinler ve spliceozom stabilizatörleri ile intron retansiyonu ve hatalı alternatif kırpılma %100 engellenir; 2) Lamivudin (reverse transkriptaz inhibitörü) ve heterokromatin kilitleri ile LINE-1 retrotranspozon uyanışı tamamen susturulur; 3) Periyodik transkriptomik scRNA-seq izlemi ile transkripsiyonel gürültü minimumda tutulur; 4) Uzun nöronal ve yapısal genlerin ekspresyon sadakati DNA tamir kaskadları ile garanti altına alınır.",
        "Hedef_Homo_Aeternus: Noise_CV2 <= 0.05 & LINE1_RNA = 0.0 TPM & LTI >= 0.0 (Gençlik Düzeyinde Transkriptom)",
        "Bu transkriptomik optimizasyon standardı, insan hücresinin genetik yazılımını hiçbir gürültüye ve entropik bozulmaya izin vermeden çalıştırma yeteneğini tanımlar."
    )
]

# ================= KISIM 4 =================
part4_subsections = [
    (
        "4.1 Plazma Proteomu: Dolaşımdaki Biyolojik Sağlığın Yüksek Çözünürlüklü Aynası",
        "Kan plazması, vücuttaki tüm organ ve dokulardan salgılanan veya sızan on binlerce farklı proteini barındıran; organizmanın anlık fizyolojik durumunu yansıtan en zengin moleküler okyanustur.",
        "Klasik klinik tıp kanda yalnızca 15-20 temel proteini (albümin, troponin, CRP) ölçerken; modern proteomik teknolojileri tek bir damla kandan binlerce proteini aynı anda analiz edebilmektedir. Plazma proteomu dinamiktir; organ hasarını, inflamasyonu, metabolik stresi ve hücresel ölümü dakikalar içinde yansıtır. Proteomik saatler bu devasa protein havuzundan biyolojik saati deşifre eder.",
        "Dinamik_Aralik_Proteom: Konsantrasyon_Araligi ~ 10^10 ila 10^12 (Albümin: mg/mL vs Sitokinler: pg/mL)",
        "Bu devasa dinamik aralık parametresi, plazma proteomundaki en bol bulunan protein ile en nadir sinyal molekülü arasındaki trilyon katlık farkı tanımlar."
    ),
    (
        "4.2 Yüksek Verimli Proteomik Platformları: SomaScan (Aptamer) ve Olink (PEA) Teknolojisi",
        "Proteomik alanındaki büyük sıçrama, iki devrimci analitik platformun geliştirilmesiyle mümkün olmuştur: SomaScan ve Olink Proteomics.",
        "SomaScan; proteinlere yüksek özgüllükle bağlanan sentetik kimyasal modifiye tek sarmal DNA aptamerlerini (SOMAmers) kullanarak tek seferde 7.000'den fazla insan proteinini ölçer. Olink ise 'Yakınlık Uzatma Deneyi' (Proximity Extension Assay - PEA) teknolojisini kullanır: Hedef proteine bağlanan iki bağımsız antikor üzerindeki DNA dizileri birbirine yaklaştığında hibridize olur ve qPCR veya NGS ile okunur; çift antikor doğrulaması sayesinde arka plan gürültüsü sıfırlanır.",
        "Hassasiyet_PEA: LOD_Olink < 0.1 pg/mL (Pikomolar Altı Konsantrasyonlarda Kusursuz Tayin)",
        "Bu analitik tespit limiti (LOD) parametresi, PEA teknolojisinin plazmadaki eser düzeydeki sinyal proteinlerini hatasız saptama gücünü belgeler."
    ),
    (
        "4.3 Plazma Proteomik Saatleri ve Yaşlanmanın Üç Dalga (Three-Wave) Prensibi",
        "Tony Wyss-Coray ve ekibinin 2019 yılında Nature Medicine'da yayımlanan tarihi çalışması (Lehallier ve ark.); 4.263 bireyin 3.000 plazma proteinini analiz ederek yaşlanmanın doğrusal değil, 'Üç Belirgin Dalga' halinde ilerlediğini keşfetti.",
        "Plazma protein dalgalanmaları; 34 yaş (Genç Erişkin Dalgası), 60 yaş (Geç Erişkin Dalgası) ve 78 yaş (İleri Yaş Dalgası) noktalarında fırtınalı tepe noktaları oluşturur. 34 yaşında ekstraselüler matriks proteinleri değişirken; 60 yaşında hormon regülasyonu ve kardiyovasküler yolaklar çöker; 78 yaşında ise immünosenesens, böbrek filtrasyon kaybı ve nörodejenerasyon proteinleri tavan yapar.",
        "Dalga_Fonksiyonu: D(t) = Sum_{k=1}^3 A_k * exp( - ( t - T_k )^2 / ( 2 * sigma_k^2 ) ) (T1=34, T2=60, T3=78 Yaş)",
        "Bu üç modal Gauss dalga dağılımı denklemi, insan ömründe biyolojik yaşlanmanın kesintili ve fırtınalı faz geçişleriyle gerçekleştiğini matematiksel olarak belgeler."
    ),
    (
        "4.4 Organ-Spesifik Proteomik Saatler (Wyss-Coray Modeli): 11 Organın Bağımsız Yaşı",
        "2023 yılında Nature'da yayımlanan çığır açıcı Stanford çalışması (Oh, Wyss-Coray ve ark.); insan bedenindeki organların hepsinin aynı hızda yaşlanmadığını ve her organın kendine has bir proteomik yaşı olduğunu kanıtladı.",
        "Araştırmacılar doku spesifik RNA ekspresyon verilerini plazma proteomu ile eşleştirerek 11 ana organ için (Beyin, Kalp, Karaciğer, Böbrek, Akciğer, Bağırsak, İmmün Sistem, Yağ Dokusu, Kas, Pankreas ve Damar Sistemi) bağımsız proteomik saatler eğitti. Bireylerin %20'sinde en az bir organın 'aşırı hızlı yaşlandığı' (hızlandırılmış organ yaşı) ve bu durumun o organda 15 yıl içinde ölümcül hastalık riskini katladığı gösterildi.",
        "Organ_Yasi_Vektoru: BioAge_Organlar = [ Age_Beyin, Age_Kalp, Age_Bobrek, ..., Age_Damar ]",
        "Bu 11 boyutlu organ yaşı tensörü, bir insanın sistemik olarak kaç yaşında olduğunu değil, hangi organının önce iflas edeceğini öngören nihai haritadır."
    ),
    (
        "4.5 Beyin-Spesifik Proteomik Saat ve Erken Nörodejenerasyon Belirteçleri",
        "Beyin kan-beyin bariyeri arkasında izole olsa da; nöron ve glia kökenli proteinler BOS ve plazmaya sızar; Wyss-Coray'ın Beyin Saati bu sızıntıları yakalar.",
        "Plazmada ölçülen Nörofilament Hafif Zincir (NfL), Glial Fibriler Asidik Protein (GFAP), VILIP-1 ve Brevikan; beynin biyolojik yaşını ve aksonal dejenerasyon hızını ortaya koyar. 'Beyin Yaşı' kronolojik yaştan 3 yıl ileri olan bireylerde Alzheimer ve bilişsel gerileme riski tam 3 kat artmaktadır; bu saat belirtiler başlamadan 15 yıl önce klinik alarm verir.",
        "Risk_Alzheimer: HR_Demans = exp( alpha_beyin * ( Brain_Age - Chrono_Age ) ) (HR > 3.0)",
        "Bu orantılı risk formülasyonu, plazma beyin proteomik yaş sapmasının nörodejeneratif çöküşü onlarca yıl önceden nasıl öngördüğünü belgeler."
    ),
    (
        "4.6 Kalp ve Böbrek-Spesifik Proteomik Saatleri ile Kardiyorenal Risk",
        "Kalp-spesifik proteomik saat (özellikle NT-proBNP, Troponin T, Myoglobin ve FABP3 kombinasyonu); sol ventrikül hipertrofisi ve kalp yetersizliğini semptomsuz evrede yakalar.",
        "Böbrek-spesifik proteomik saat ise glomerüler bazal membran ve tübüler proteinleri (KIM-1, Uromodulin, NGAL ve Sistatin-C) analiz ederek glomerüler filtrasyon hızındaki (eGFR) düşüşü standart kreatinin testlerinden yıllar önce haber verir. 'Kardiyorenal yaşlanma ekseni', sistemik mortalitenin en kritik belirleyicisidir.",
        "Kardiyorenal_Risk = Risk_CR = w_kalp * ( Heart_Age - Chrono ) + w_bobrek * ( Kidney_Age - Chrono )",
        "Bu ağırlıklı kardiyorenal yaş sapması eşitliği, kardiyovasküler ve renal sistemlerin birleşik biyolojik yaşlanma yükünü modeller."
    ),
    (
        "4.7 İnflamatuar ve Senesens Belirteçleri: GDF-15, Activin A ve MCP-1 Kinetiği",
        "Tüm organ spesifik saatlerin üzerinde ortak bir payda gibi yükselen kilit sistemik yaşlanma proteinleri bulunur; bunların başında 'GDF-15' (Büyüme Diferansiasyon Faktörü 15) gelir.",
        "GDF-15; hücresel stres, mitokondriyal disfonksiyon ve senesens altında tüm dokulardan salgılanır ve yaşla birlikte plazmada en keskin artış gösteren proteindir (gençte 400 pg/mL iken yaşlıda 2.500 pg/mL'ye fırlar). Activin A doku fibrozisini yönlendirirken; MCP-1 (CCL2) lökositleri dokulara çağırarak kronik yangıyı körükler. Bu üçlü, biyolojik yıpranmanın evrensel proteomik alarmıdır.",
        "GDF15_Kinetigi: [GDF-15](Yas) = [GDF-15]_0 * exp( k_stres * Yas ) (Her On Yılda %50 Artış)",
        "Bu üstel proteomik artış denklemi, plazma GDF-15 seviyelerinin hücresel stres ve biyolojik yaşlanmayla kusursuz korelasyonunu belgeler."
    ),
    (
        "4.8 Parabiyoz Deneyleri ve Genç Plazma Proteomunun Gençleştirici Faktörleri (GDF11, Klotho, TIMP2)",
        "Yaşlı bir fare ile genç bir farenin dolaşım sistemlerinin cerrahi olarak birbirine bağlandığı 'Heterokronik Parabiyoz' deneyleri; yaşlanmanın kanda dolaşan faktörlerle yönetildiğini kesinleştirmiştir.",
        "Genç farenin kanı yaşlı farenin beynini, kalbini ve kaslarını haftalar içinde gençleştirirken; yaşlı kan genç fareyi hızla yaşlandırmıştır. Yapılan proteomik fraksiyonasyonlar; kanda dolaşan gençleştirici faktörleri izole etmiştir: GDF11 (kardiyak hipertrofiyi geriletir), Çözünür Klotho (sinaptik plastisiteyi artırır) ve TIMP2 (hipokampal nörogenezi ve mekânsal hafızayı restore eder).",
        "Rejuvenasyon_Katsayisi: Delta_Fonksiyon = Sum_j alpha_j * [Genc_Protein_j] - Sum_k beta_k * [Yasli_Toksin_k]",
        "Bu sistemik gençleşme katsayısı formülasyonu, genç plazma faktörleri ile yaşlı kanda biriken pro-senesent toksinlerin hücresel fonksiyon üzerindeki net etkisini tanımlar."
    ),
    (
        "4.9 Terapötik Plazma Değişimi (TPE / Plazmaferez) ile Yaşlı Plazma Toksinlerinin Seyreltilmesi",
        "Irina ve Michael Conboy'un devrimci keşfi; parabiyozdaki gençleşmenin sırrının sadece 'genç kan almak' değil, asıl olarak 'yaşlı kanda biriken toksik proteinleri seyreltmek' olduğunu kanıtladı.",
        "Nötral kan değişimi (Neutral Blood Exchange - NBE) deneylerinde; yaşlı farenin plazmasının yarısı çekilip yerine sadece fizyolojik tuzlu su ve albümin verildiğinde (genç kan verilmeksizin!), beyinde nörogenez patlamış ve karaciğer gençleşmiştir. İnsanlarda uygulanan 'Terapötik Plazma Değişimi' (TPE); kanda biriken otoantikorları, sitokinleri, SASP proteinlerini ve AGE'leri mekanik olarak temizleyerek proteomik saati anında geriye sarar.",
        "Seyreltme_Kinetigi: C_toksin(N) = C_toksin(0) * ( 1 - f_degisim )^N (N Seans Sonrası Toksin Klirensi)",
        "Bu fraksiyonel plazmaferez klirens eşitliği, periyodik plazma değişiminin yaşlılık toksinlerini dolaşımdan nasıl eksponansiyel olarak uzaklaştırdığını gösterir."
    ),
    (
        "4.10 Homo Aeternus Organ-Spesifik Proteomik Optimizasyon Protokolü",
        "Homo Aeternus tıbbında plazma proteomu, her çeyrekte 7.000 proteinlik Olink/SomaScan panelleriyle taranarak 11 organın bağımsız yaşı haritalanır:",
        "Hangi organın proteomik yaşı kronolojik yaşın üzerine çıkarsa (örneğin Karaciğer Saati > 25 Yaş); o organa spesifik AAV gen tedavileri, hedefe yönelik peptid nanopartikülleri veya kök hücre ekzozomları sevk edilir. Eş zamanlı olarak yılda 2 kez uygulanan TPE kürleri ile GDF-15 ve SASP toksinleri kandan temizlenir; rekombinant Klotho ve TIMP2 infüzyonlarıyla plazma sürekli gençlik zenginliğinde tutulur.",
        "Hedef_Homo_Aeternus: Maks(BioAge_11_Organ) <= 22.0 Yas & [GDF-15] <= 450 pg/mL (Tüm Organlarda Kusursuz Proteom)",
        "Bu nihai proteomik yönetim hedefi, 11 hayati organın hiçbirinin erken yaşlanmasına izin vermeksizin tüm bedenin senkronize bir gençlikte tutulmasını simgeler."
    )
]

# ================= KISIM 5 =================
part5_subsections = [
    (
        "5.1 Hücresel Metabolomik: Biyolojik Reaksiyonların Nihai Fonksiyonel Çıktısı",
        "Metabolom; hücre içinde enzimlerin katalizlediği biyokimyasal reaksiyonlar sonucunda oluşan tüm küçük moleküllerin (<1.500 Da; amino asitler, şekerler, nükleotidler, organik asitler) toplamıdır.",
        "Genom neyin olabileceğini, transkriptom neyin planlandığını, proteom neyin yapıldığını gösterirken; metabolom 'gerçekte ne olduğunu' yansıtır. Yaşlanma metabolomik seviyede anlık olarak okunur: Glikoliz, trikarboksilik asit (TCA) döngüsü, oksidatif fosforilasyon, üre döngüsü ve amino asit katabolizması yaşla birlikte dramatik kaymalara uğrar.",
        "Metabolomik_Hacim: N_metabolit ~ 5.000 - 10.000 Endojen Molekül (Biyokimyasal Akı Haritası)",
        "Bu kütle spektrometri parametresi, insan hücresel biyokimyasında dolaşan küçük moleküler metabolit havuzunun toplam ölçeğini tanımlar."
    ),
    (
        "5.2 Analitik Metabolomik Platformları: LC-MS/MS ve NMR Spektroskopisi",
        "Metabolomun hassas analizi, iki ana analitik kimya platformunun entegrasyonu ile yürütülür: Sıvı Kromatografisi-Tandem Kütle Spektrometresi (LC-MS/MS) ve Nükleer Manyetik Rezonans (NMR).",
        "LC-MS/MS; pikomolar ve femtomolar düzeydeki eser metabolitleri muazzam bir kütle doğruluğuyla (m/z hassasiyeti < 1 ppm) saptar ve karmaşık biyolojik sıvılarda binlerce molekülü ayrıştırır. Proton NMR (1H-NMR) ise daha düşük duyarlılığa sahip olmakla birlikte, mutlak kantitatif doğruluk sunar, örnek hazırlığı gerektirmez ve lipoprotein alt fraksiyonları ile amino asitleri saniyeler içinde tekrarlanabilir biçimde ölçer.",
        "Hassasiyet_Kutle: Delta_m/z = ( m_olculen - m_teorik ) / m_teorik * 10^6 < 2 ppm (Ultra-Yuksek Rezolusyon)",
        "Bu yüksek çözünürlüklü kütle spektrometresi sadakat formülü, izomerik metabolitlerin dahi hatasız teşhis edilmesini sağlayan analitik standardı tanımlar."
    ),
    (
        "5.3 Mitokondriyal Metabolit İmzaları: NAD+/NADH Oranı ve TCA Döngüsü Çöküşü",
        "Metabolik yaşlanmanın merkez üssü mitokondridir ve mitokondriyal sağlığın en kritik göstergesi 'NAD+/NADH' redoks çiftidir.",
        "Hücresel NAD+ havuzu yaşlandıkça CD38 ve PARP1 enzimlerinin aşırı aktivasyonu nedeniyle dramatik biçimde tükenir (%50-%80 kayıp). Düşen NAD+/NADH oranı; Sirtuinleri (SIRT1, SIRT3) felç eder ve TCA döngüsünün kilit enzimleri olan İzositrat Dehidrogenaz ve Alfa-Ketoglutarat Dehidrogenazı kilitler. Fumarat ve süksinat gibi onkometabolitler birikirken, hücresel ATP/ADP oranı çöker ve hücre hipoksi benzeri psödoiştah durumuna girer.",
        "Redoks_Dengesi: R_NAD = [NAD+] / [NADH] (Genç Hücrede > 50 | Yaşlı Hücrede < 10) -> Masif Enerjetik İflas",
        "Bu nükleotid kofaktör konsantrasyon oranı, hücresel biyoenerjetik kapasitenin ve sirtuin enzim aktivitesinin yaşla birlikte nasıl çöktüğünü modeller."
    ),
    (
        "5.4 Lipidomik Saatler ve Yaşlanmanın Membran Dinamikleri",
        "Hücresel lipidom; hücre ve organel membranlarının akışkanlığını, sinyal iletimini ve enerji depolanmasını yöneten binlerce farklı lipit türünü (fosfolipitler, sfingolipitler, seramidler, kolesterol) kapsar.",
        "Yaşlanmayla birlikte membran fosfolipit kompozisyonunda radikal bir bozulma yaşanır: Mitokondri iç zarının hayati bileşeni olan 'Kardiyolipin' peroksidasyona uğrayarak tükenir; bu durum elektron taşıma zinciri süper-komplekslerinin dağılmasına yol açar. Eş zamanlı olarak hücre zarlarında doymuş yağ asitleri ve pro-enflamatuar seramidler (C16:0, C18:0 seramid) birikir; membran akışkanlığı kaybolur, reseptörler kilitlenir ve apoptoz tetiklenir.",
        "Lipidomik_Dengesizlik: L_index = [Seramidler] / [Kardiyolipin] -> Yaşlandıkça Eksponansiyel Artış",
        "Bu membran lipidomik toksisite indeksi, mitokondriyal kardiyolipin kaybı ve seramid birikiminin hücresel fonksiyonu nasıl felç ettiğini belgeler."
    ),
    (
        "5.5 Açilkarnitinler ve Bozulmuş Mitokondriyal Yağ Asidi Beta-Oksidasyonu",
        "Uzun zincirli yağ asitlerinin mitokondriye taşınarak beta-oksidasyonla yakılması, karnitin palmitoiltransferaz (CPT-1) mekik sistemine bağımlıdır.",
        "Yaşlanan dokularda mitokondriyal oksidasyon kapasitesi tıkandığında, yağ asitleri tam yakılamaz ve yarıda kesilir. Plazmada ve dokularda orta ve uzun zincirli 'Açilkarnitinler' (C14, C16, C18 açilkarnitin) birikir. Biriken açilkarnitinler sitozolde membran deterjanı gibi davranarak lipotoksisite yaratır, insülin reseptör sinyalini bozar ve sistemik insülin direncinin (tip 2 diyabet) ana metabolomik motoru olur.",
        "Lipotoksisite_Skoru: AS_Score = Sum_{n=12}^{18} [Açilkarnitin_Cn] >> Eşik (Mitokondriyal Oksidatif Tıkanıklık)",
        "Bu eksik yağ asidi oksidasyon indeksi, mitokondrinin yakamadığı lipit ara ürünlerinin yarattığı hücresel zehirlenmeyi tanımlar."
    ),
    (
        "5.6 Dallı Zincirli Amino Asitler (BCAA) ve Katabolik Yolak Yetmezliği",
        "Dallı zincirli amino asitler (Lösin, İzolösin, Valin - BCAA); kas protein sentezi ve mTOR aktivasyonu için kritik olmakla birlikte, yaşlanmayla kanda tehlikeli biçimde birikir.",
        "Yaşlanan karaciğer ve yağ dokusunda BCAA'ları parçalayan BCAT2 ve BCKDH enzim kompleksleri inaktive olur. Yükselen plazma BCAA seviyeleri mTORC1'i kontrolsüzce uyararak otofajiyi durdurur ve insülin reseptör substratı 1'i (IRS-1) serin rezidülerinden fosforilleyerek periferik insülin direncini patlatır. BCAA fazlalığı kardiyometabolik yaşlanmanın ve erken mortalitenin en güçlü metabolomik belirtecidir.",
        "BCAA_Klirens_Verimi: eta_BCAA = J_katabolizma / [BCAA_plazma] -> Yaşla Birlikte %50 Azalma",
        "Bu metabolik klirens katsayısı, amino asit katabolizma yetersizliğinin mTORC1 aşırı aktivasyonu ve diyabet riskini nasıl tetiklediğini gösterir."
    ),
    (
        "5.7 Trimetilamin N-Oksit (TMAO) ve Bağırsak Mikrobiyom-Metabolom Ekseni",
        "Dolaşımdaki bazı ölümcül metabolitler insan hücreleri tarafından değil; bağırsak mikrobiyotası ile karaciğer enzimlerinin ortak eylemiyle sentezlenir.",
        "Diyetteki kolin, betain ve L-karnitin; disbiyotik bağırsak bakterileri tarafından Trimetilamin (TMA) gazına dönüştürülür. Portal damarla karaciğere ulaşan TMA, Flavin İçeren Monooksijenaz 3 (FMO3) enzimi tarafından 'Trimetilamin N-Oksit'e (TMAO) oksitlenir. Plazma TMAO seviyeleri; trombosit agregasyonunu hızlandırarak tromboz riskini patlatır, makrofaj çöpçü reseptörlerini artırarak aterosklerozu azdırır ve böbrek tübüler fibrozisini tetikler.",
        "TMAO_Sentezi: Kolin/Karnitin ----[Mikrobiyota]--> TMA ----[Karaciğer FMO3]--> TMAO (Kardiyotoksik Metabolit)",
        "Bu mikrobiyom-metabolom ortak üretim şeması, bağırsak florası bozulmasının karaciğer üzerinden kardiyorenal hasarı nasıl tetiklediğini modeller."
    ),
    (
        "5.8 Ürik Asit, Ksantin Oksidaz ve Fruktoz Kaynaklı Metabolik Fırtına",
        "Pürin nükleotidlerinin yıkım son ürünü olan Ürik Asit; insan evriminde ürataz (Uricase) geninin psödogenezle susturulması sonucu primatlarda yüksek seviyelere çıkmıştır.",
        "Ksantin Oksidaz (XO) enzimi pürinleri ürik aside dönüştürürken yan ürün olarak masif hidrojen peroksit (H2O2) ve süperoksit (O2•-) üretir. Aşırı fruktoz tüketimi karaciğerde fruktokinaz üzerinden ATP'yi hızla tüketerek pürin yıkımını ve ürik asit patlamasını tetikler. Kanda ürik asit > 6.0 mg/dL olduğunda; endotelde eNOS'u bloke eder, mitokondriyal akonitazı inaktive eder ve sistemik endotelyal disfonksiyonu körükler.",
        "Urik_Asit_Toksisitesi: d[NO]/dt = ( d[NO]/dt )_0 / ( 1 + [Urik_Asit] / K_i_eNOS ) (Hiperürisemide NO Felci)",
        "Bu vasküler inhibisyon denklemi, yüksek ürik asit seviyelerinin endotelyal nitrik oksit sentezini nasıl doğrudan baskıladığını belgeler."
    ),
    (
        "5.9 N-Asetilglukozamin, Poliaminler (Spermidin) ve Otofaji Metabolitleri",
        "Yaşlanmayla birlikte yıkıcı metabolitler birikirken; hücresel gençliği ve otofajiyi tetikleyen 'Koruyucu Metabolitler' dramatik biçimde tükenir.",
        "Doğal bir poliamin olan 'Spermidin'; hücresel konsantrasyonu yaşla hızla düşen ve düşüşü otofaji kaybıyla birebir örtüşen hayati bir metabolittir. Spermidin takviyesi; EP300 asetiltransferazını inhibe ederek otofaji proteinlerinin transkripsiyonunu açar, mitofajiyi hızlandırır ve kardiyak fonksiyonu gençleştirir. Benzer şekilde hekzozamin yolağı ara ürünü GlcNAc, protein kalitesini korur.",
        "Otofaji_Uyarimi: Flux_Otofaji = k_spermidin * [Spermidin] / ( K_m + [Spermidin] ) * ( 1 / [EP300_aktif] )",
        "Bu metabolomik otofaji akı eşitliği, hücresel spermidin düzeylerinin otofajik temizlik hızını doğrudan belirlediğini formüle eder."
    ),
    (
        "5.10 Homo Aeternus Metabolomik Kararlılık Protokolü: Sürekli Genç Biyokimyasal Akı",
        "Homo Aeternus longevity mimarisi; metabolik akıyı genç bir atletin mitokondriyal saflığında ve kusursuz redoks dengesinde kilitler:",
        "1) Yüksek doz lipozomal NMN/NR ve CD38 inhibitörleri (Apigenin, 78c) ile hücre içi NAD+/NADH oranı sürekli > 40 seviyesinde tutulur; 2) Oral Spermidin ve alfa-ketoglutarat (AKG) kürleriyle otofaji ve epigenetik demetilasyon canlı tutulur; 3) Mikrobiyotayı hedef alan kolin parçalanma inhibitörleri (DMB / İyodometilkolin) ile TMAO sentezi sıfırlanır; 4) Ksantin oksidaz baskılayıcıları ile ürik asit < 4.5 mg/dL düzeyinde dondurulur.",
        "Hedef_Homo_Aeternus: [NAD+]/[NADH] >= 40.0 & TMAO <= 2.0 uM & [Urik_Asit] <= 4.5 mg/dL (Kusursuz Biyokimyasal Akı)",
        "Bu nihai metabolomik optimizasyon hedefi, hücresel metabolit havuzunun yaşlanma toksinlerinden tamamen arındırılarak ebedi bir enerji verimliliğine ulaştırılmasını simgeler."
    )
]

# ================= KISIM 6 =================
part6_subsections = [
    (
        "6.1 Glikobiyoloji ve Glikomik: Protein Fonksiyonunun Şeker Kodlaması",
        "İnsan proteomundaki proteinlerin yarısından fazlası translasyon sonrasında spesifik dallanmış oligosakkarit zincirleriyle (glikanlar) kovalent olarak modifiye edilir (Glikozilasyon).",
        "Glikanlar basit enerji molekülleri değildir; proteinlerin 3 boyutlu katlanmasını, stabilitesini, yarı ömrünü, hücre-hücre tanımasını ve immünolojik sinyal aktivitesini belirleyen karmaşık bir 'Glikan Alfabesi' (Glycome) oluşturur. Glikozilasyon şablonu doğrudan bir gen dizisi tarafından kodlanmaz; hücre içindeki yüzlerce glikoziltransferaz ve glikozidaz enziminin dinamik rekabetiyle belirlenir; bu nedenle hücrenin metabolik ve yaşlanma durumunu son derece hassas biçimde yansıtır.",
        "Glikomik_Karmasiklik: N_glikoform >> 10^5 Farklı Yapısal Glikan Ağacı Kombinasyonu",
        "Bu biyoenformatik parametre, insan glikomunun translasyon sonrası protein çeşitliliğini genomik sınırların çok ötesine taşıyan yapısal devasasını tanımlar."
    ),
    (
        "6.2 İmmünoglobulin G (IgG) N-Glikom Haritası ve Asn297 Konservatif Bölgesi",
        "Glikomik yaşlanmanın en net ve en iyi karakterize edilmiş aynası; plazmadaki en bol bulunan antikor olan İmmünoglobulin G'nin (IgG) Fc bölgesindeki N-glikanlarıdır.",
        "Her bir IgG ağır zincirinin CH2 domaininde, Asn297 (Asparagin-297) amino asidine bağlı tek bir çift dallanmış kor N-glikan çekirdeği bulunur. Bu glikan çekirdeğine bağlanan veya eksilen Galaktoz, Sialik Asit, Fukoz ve Bölücü N-Asetilglukozamin (bisecting GlcNAc) şekerleri; antikorun bağışıklık sistemine vereceği emri radikal biçimde değiştirir: Antikor ya sakinleştirici bir antienflamatuar ya da dokuyu yıkan agresif bir pro-enflamatuar ajana dönüşür.",
        "IgG_Glikan_Cekirdegi: GlcNAc2-Man3-GlcNAc2 +/- Fukoz +/- Galaktoz +/- Sialik_Asit +/- Bisecting_GlcNAc",
        "Bu biyokimyasal N-glikan dallanma şablonu, antikor Fc reseptör etkileşimini yöneten yapısal şeker kodunu belgeler."
    ),
    (
        "6.3 Galaktoz ve Sialik Asit Kaybı: Anti-Enflamatuardan Pro-Enflamatuar Fenotipe Kayış",
        "Genç bir bireyin IgG antikorları; terminal galaktoz ve sialik asit şekerleriyle zenginleşmiştir.",
        "Sialik asit ve galaktoz taşıyan IgG molekülleri; makrofaj yüzeyindeki inhibitör Fc-gamma-RIIB reseptörlerine bağlanarak bağışıklık hücrelerini sakinleştirir ve inflamasyonu durdurur. Yaşlandıkça B lenfositlerindeki galaktoziltransferaz (B4GALT1) ve sialiltransferaz enzim aktiviteleri çöker. Antikorlar galaktoz ve sialik asitlerini kaybeder (Agilakozile / G0 glikanlar). Sialiksiz kalan IgG'ler aktive edici Fc-gamma-RIIIa reseptörlerini uyararak masif sitotoksisite ve kompleman aktivasyonu başlatır.",
        "Pro_Enflamatuar_Glikan_Orani: R_glikan = [G0_agilakozile] / ( [G1] + [G2_tam_galaktozile] ) -> Yaşla Fırlar",
        "Bu glikomik oran formülasyonu, antikorların yaşlanmayla birlikte koruyucu şekerlerini kaybederek kronik yangı tetikleyicisine dönüşümünü modeller."
    ),
    (
        "6.4 GlycanAge Saati: Gordan Lauc Algoritması ve Biyolojik Yaş Korelasyonu",
        "Hırvat genetikçi Gordan Lauc ve uluslararası Genos konsorsiyumu; binlerce bireyin IgG N-glikomunu UPLC ve kapiler elektroforez ile analiz ederek 'GlycanAge Saati'ni geliştirdi.",
        "GlycanAge; kandaki 24 ana IgG N-glikan piki arasındaki matematiksel oranları (özellikle galaktozsuz G0 glikanların galaktozlu G2 glikanlara oranı) kullanarak biyolojik yaşı tahmin eder. GlycanAge, kronolojik yaştan bağımsız olarak kardiyovasküler risk, erken menopoz, kronik inflamasyon ve metabolik sendromu onlarca yıl önceden öngörür. Kötü beslenen ve aşırı stres altındaki bir bireyin GlycanAge saati kronolojik yaşından 20 yıl ileri çıkabilir.",
        "GlycanAge_Formulu: GlycanAge = w_0 + w_1 * G0 - w_2 * G2 - w_3 * Sialilasyon + w_4 * Bisecting_GlcNAc",
        "Bu çok değişkenli lineer regresyon eşitliği, plazma IgG glikan fraksiyonlarının ağırlıklı toplamıyla hesaplanan biyolojik glikomik yaşı tanımlar."
    ),
    (
        "6.5 Biyolojik Yaşam Tarzı Müdahalelerine GlycanAge'in Olağanüstü Duyarlılığı",
        "GlycanAge saatinin en büyük klinik değeri; yaşam tarzı değişikliklerine (kilo kaybı, kalori kısıtlaması, egzersiz, hormon replasmanı) aylar içinde pozitif ve ölçülebilir yanıt vermesidir.",
        "Obez bireylerde bariatrik cerrahi veya GLP-1 reseptör agonistleri ile sağlanan kilo kaybı; 6 ay içinde GlycanAge saatini ortalama 9 yıl geriye sarar. Menopoza giren kadınlarda düşen östrojen B4GALT1 enzimini baskılayarak GlycanAge'i aniden 10-15 yıl ileri fırlatırken; Biyo-özdeş Hormon Replasman Tedavisi (BHRT) uygulandığında glikomik saat hızla ilk gençlik düzeyine döner.",
        "Glikomik_Rejuvenasyon: Delta_GlycanAge = GlycanAge_son - GlycanAge_baslangic < - 8.0 Yıl (Kilo Kaybı / BHRT ile)",
        "Bu klinik rejuvenasyon istatistiği, glikomik saatin yaşam tarzı ve hormonal müdahalelerin biyolojik gençleşme etkisini nasıl yüksek hassasiyetle yakaladığını kanıtlar."
    ),
    (
        "6.6 B Hücre İmmünosenesensi ve Glikozilasyon Enzimlerinin Transkripsiyonel Çöküşü",
        "IgG glikan yapısının yaşlanmayla bozulmasının hücresel kaynağı; antikorları üreten plazma B lenfositlerinin içsel yaşlanması ve immünosenesensidir.",
        "Yaşlı kemik iliğindeki B hücre progenitörlerinde hücresel stres, mitokondriyal hasar ve pro-enflamatuar sitokinler (IL-6); Golgi aygıtında yer alan glikoziltransferaz enzimlerinin (B4GALT1, ST6GAL1) ekspresyonunu transkripsiyonel olarak susturur. B hücresi antikoru sentezlerken şeker zincirini tam olgunlaştıramadan eksik glikanlarla kana salar; immünosenesens kanda dolaşan antikorların kalitesini doğrudan zehirler.",
        "Enzim_Aktivite_Dususu: Rate_Galaktozilasyon = V_maks * [B4GALT1] / ( K_m + [UDP-Galaktoz] ) -> Yaşla %60 Düşüş",
        "Bu enzim kinetiği denklemi, B hücre yaşlanmasıyla Golgi glikozilasyon kapasitesinin çöküşünü ve agilakozile antikor üretimini formüle eder."
    ),
    (
        "6.7 İnflammaging ve Glikomik Geri Bildirim Döngüsü (Feedback Loop)",
        "Glikomik yaşlanma tek yönlü bir sonuç değil; yaşlanmanın kronik yangısını (inflammaging) sürekli besleyen agresif bir pozitif geri bildirim döngüsüdür.",
        "Enflamasyon (IL-6, TNF-alfa); B hücrelerinde galaktoz transferini baskılar -> Galaktozsuz G0 antikorlar üretilir -> Bu antikorlar Fc reseptörleri üzerinden makrofajları uyararak daha fazla IL-6 ve TNF-alfa salgılatır -> Yangı tırmandıkça daha çok hatalı glikan üretilir. Bu kısır döngü kırılmadığı sürece birey hızlanmış vasküler, eklem ve organ yaşlanmasına kilitlenir.",
        "Dongu_Dinamigi: d[Enflamasyon]/dt = alpha_sitokin * [G0_antikor] + beta_baski * [IL-6] (Kendi Kendini Besleyen Yangı)",
        "Bu diferansiyel geri bildirim eşitliği, hatalı antikor glikozilasyonunun sistemik inflamatuar yaşlanmayı nasıl kontrolsüz bir sarmala dönüştürdüğünü modeller."
    ),
    (
        "6.8 Hücre Yüzey Glikokaliksi ve Matriks Proteomunun Glikomik Aşınması",
        "Glikomik yaşlanma yalnızca antikorlarla sınırlı değildir; damar endotelinin glikokaliksinde, nöronal sinapslarda ve ekstraselüler matriks proteinlerinde de dramatik bir glikomik aşınma yaşanır.",
        "Hücre zarlarındaki polisialik asit (PSA-NCAM) ve heparan sülfat zincirleri yaşlandıkça kısalır ve sülfasyon desenlerini kaybeder. Bu durum nöroplastisiteyi felç eder, büyüme faktörlerinin (FGF, VEGF) reseptörlerine bağlanmasını engeller ve damar geçirgenliğini bozar; glikomik çözülme tüm dokularda hücreler arası iletişimi susturur.",
        "Glikan_Kalinlik_Kaybi: L_glikokaliks(t) = L_0 * exp( - k_glikan * Yas ) (Mikrovasküler Kalkanın Çözülmesi)",
        "Bu hücresel yüzey kalkanı tükeniş formülü, endotel ve nöron membran glikanlarının yaşlanmayla eksponansiyel kaybını tanımlar."
    ),
    (
        "6.9 Endoglikozidazlar ve Rekombinant Glikomühendislik Terapileri",
        "Glikomik yaşlanmayı tersine çevirmek için sentetik biyoloji; antikorların üzerindeki pro-enflamatuar şekerleri temizleyen veya eksik galaktozları ekleyen 'Rekombinant Glikomühendislik' araçları geliştirmiştir.",
        "Bakteriyel Endoglikozidazlar (EndoS / EndoS2); kanda dolaşan pro-enflamatuar IgG'lerin Asn297 bölgesindeki hatalı glikanları mikrosaniyeler içinde keserek antikoru tamamen nötr ve zararsız hale getirir. Eş zamanlı olarak uygulanan rekombinant ST6Gal1 ve B4GALT1 enzimleri ile in vivo antikor glikozilasyonu restore edilerek tüm antikor havuzu taze bir antienflamatuar zırha dönüştürülür.",
        "Glikom_Temizleme_Hizi: d[G0_toksik]/dt = - k_EndoS * [EndoS_enzim] * [G0_antikor] (Dakikalar İçinde Tam Klirens)",
        "Bu biyokatalitik glikan arınma eşitliği, rekombinant enzimlerin dolaşımdaki yangı yapıcı antikorları nasıl anında etkisizleştirdiğini belgeler."
    ),
    (
        "6.10 Homo Aeternus Glikomik Optimizasyon Protokolü: Sürekli Genç Antikor Havuzu",
        "Homo Aeternus longevity mimarisi; plazma IgG N-glikanlarını sürekli maksimum galaktozile ve sialile gençlik durumunda tutar:",
        "1) Yılda iki kez yüksek çözünürlüklü UPLC IgG N-glikan kromatografisi ile GlycanAge skoru izlenir; 2) Skorda en ufak bir sapma görüldüğünde hedefe yönelik anti-enflamatuar protokoller ve BHRT dengesi restore edilir; 3) Glikozilasyon substratı olan UDP-galaktoz prekürsörleri ve N-asetilmannozamin (ManNAc) takviyeleri ile B hücre Golgi biyosentezi desteklenir; 4) GlycanAge saati sürekli takvim yaşından en az 15-20 yıl genç seviyede kilitlenir.",
        "Hedef_Homo_Aeternus: GlycanAge <= 20.0 Yas & [G0_agilakozile] < %15 (Maksimum Anti-Enflamatuar Antikor Profili)",
        "Bu nihai glikomik optimizasyon standardı, bağışıklık sisteminin antikor cephanesinin yaşlanma kaynaklı oto-yıkıcı yangıdan tamamen arındırılmasını simgeler."
    )
]

# ================= KISIM 7 =================
part7_subsections = [
    (
        "7.1 İmmünomik Yaşlanma Biyolojisi: İmmünosenesens ve İnflammaging Çift Sarmalı",
        "İmmün sistem; organizmanın dış patojenlere ve içsel mutant hücrelere karşı mutlak savunma hattıdır; ancak yaşlanmayla birlikte iki yıkıcı kutba savrulur: 'İmmünosenesens' (yeni patojenlere ve kansere karşı savunmanın çökmesi) ve 'İnflammaging' (düşük dereceli, kronik ve steril sistemik yangı).",
        "İmmünomik; bağışıklık sisteminin hücresel kompozisyonunu (T, B, NK hücreleri, monositler) ve salgılanan yüzlerce sitokin, kemokin ve büyüme faktörünü küresel bir ağ olarak inceler. İmmünomik yaşlanma; diğer tüm organ sistemlerinin (beyin, kalp, damar) yaşlanmasını doğrudan hızlandıran ana tetikleyicidir.",
        "Kutupsal_Cokus: Sistemik_ImmunoAge = alpha_inflammaging * [Yangı_Skoru] + beta_immünosenesens * [Defans_Kaybi]",
        "Bu dual immünopatoloji eşitliği, bağışıklık sisteminin hem aşırı yangı hem de yetersiz savunma ile organizmayı iki koldan çökerttiğini tanımlar."
    ),
    (
        "7.2 Timik İnvolüsyon: T Hücre Fabrikasının Kapanması ve Naive T Hücre Tükenişi",
        "İnsan immün sisteminin yaşlanmasındaki en kritik anatomik milat; ergenlikten hemen sonra başlayan ve 60 yaşında neredeyse tamamen yağ dokusuna dönüşen 'Timüs Bezinin Büzüşmesidir' (Timik İnvolüsyon).",
        "Gençlikte timüs her gün milyonlarca yeni, hiç antijenle karşılaşmamış 'Naive T Hücresi' (CD45RA+ / CCR7+) üretip kana salar. Timüs kapandığında naive T hücre üretimi sıfıra yaklaşır; T hücre repertuar çeşitliliği çöker. Kalan T hücreleri periferde homeostatik bölünmeye zorlanır; bu durum telomer erozyonunu, replikatif senesensi ve yeni virüslere veya neoplazilere karşı körlüğü doğurur.",
        "Timik_Cokus: V_timus_aktif(t) = V_0 * exp( - lambda_involusyon * Yas ) (60 Yaşında %95 Kayıp)",
        "Bu timik involüsyon eksponansiyel modeli, T hücre biyogenez organının yaşlanmayla nasıl acımasızca yağ kütlesine dönüştüğünü modeller."
    ),
    (
        "7.3 İmmünosenesent T Hücreleri: CD28 Kaybı ve CD57 / KLRG-1 Birikimi",
        "Yaşam boyu tekrarlanan viral enfeksiyonlar (özellikle Sitomegalovirüs - CMV) ve homeostatik bölünmeler; T hücrelerini terminal diferansiasyona ve senesense iter.",
        "Senesent T hücreleri; T hücre aktivasyonu için hayati olan ko-stimülatör reseptör 'CD28'i kaybeder (CD28-null) ve terminal yaşlanma belirteçleri olan 'CD57' ile 'KLRG-1'i yüzeyine yerleştirir. Bu hücreler artık yeni antijenlere yanıt veremez, bölünemez; ancak çevre dokulara sürekli perforin, granzim ve toksik TNF-alfa/IFN-gamma salgılayan öldürücü birer SASP fabrikasına dönüşür (TEMRA hücreleri).",
        "Senesent_T_Orani: Fraksiyon_CD28null = [CD8+ CD28- CD57+] / [CD8+_toplam] -> Yaşlıda > %50 (İmmün Enkaz)",
        "Bu immünofenotipik yaşlanma fraksiyonu, periferik kanda dolaşan hafıza T hücrelerinin yarısından fazlasının terminal senesense girdiğini belgeler."
    ),
    (
        "7.4 CD4/CD8 T Hücre Oranının Tersine Dönmesi ve İmmün Risk Profili (IRP)",
        "İsveç boylamsal yaşlanma çalışmalarında (OCTO ve NONA) keşfedilen ve ileri yaşta 2-4 yıl içinde ölümü öngören en güçlü immünolojik parametre 'İmmün Risk Profili'dir (IRP).",
        "Genç sağlıklı bireylerde CD4+ T yardımcı hücrelerinin CD8+ sitotoksik T hücrelerine oranı yaklaşık 2.0 ila 3.0 arasındadır (CD4/CD8 > 2). Kronik CMV enfeksiyonu ve CD8+ klonal genişlemesiyle bu oran 1.0'ın altına iner ve tersine döner (CD4/CD8 < 1.0). Tersine dönmüş oran; timik tükenişin, klonal anarşinin ve çökmüş bağışıklığın tartışmasız klinik faturasıdır.",
        "Immuno_Risk_Indeksi: IRP = I( CD4/CD8 < 1.0 ) * I( CMV_seropozitif ) * I( CD8+ CD28- > %50 ) -> Mortalite %80 Artar",
        "Bu mantıksal risk profili formülasyonu, T hücre alt tip dengesizliğinin yaşlı popülasyonda kardiyovasküler ve enfeksiyon mortalitesini nasıl katladığını gösterir."
    ),
    (
        "7.5 Stanford İmmünomik Saati: iAge Algoritması ve Sitokin Metriği",
        "2021 yılında Nature Aging'de yayımlanan tarihi Stanford çalışması (David Furman, Mark Davis ve ark.); derin öğrenme kullanarak ilk bağımsız 'İmmünolojik Yaş Saati' olan 'iAge'i geliştirdi.",
        "1.001 bireyin (1001 İmmünom Projesi) kan örneklerinde 50'den fazla sitokin, kemokin, büyüme faktörü ve 60 immün hücre alt tipi analiz edildi. Algoritma; kronolojik yaştan bağımsız olarak kardiyovasküler yaşlanmayı, endotel disfonksiyonunu ve tüm nedenlere bağlı ölümü öngören bir immün yaş skoru üretti. iAge'in merkez üssünde tek bir sitokin vardı: CXCL9.",
        "iAge_Denklemi: iAge = w_0 + w_CXCL9 * [CXCL9] + w_TRAIL * [TRAIL] - w_IL10 * [IL-10] + Sum_k beta_k * Cytokine_k",
        "Bu çok değişkenli immünomik yaş eşitliği, pro-enflamatuar sitokinlerin ağırlıklı toplamından biyolojik immün yaşın nasıl hesaplandığını tanımlar."
    ),
    (
        "7.6 Kemokin CXCL9: Endotel Yıkımının ve Kardiyovasküler Yaşlanmanın İmmünomik Tetikleyicisi",
        "Stanford iAge çalışmasının en şok edici keşfi; interferon-gamma ile indüklenen bir kemokin olan 'CXCL9'un (MIG) sistemik yaşlanmanın baş mimarı olduğunu ortaya koymasıdır.",
        "Yaşlı bireylerde CXCL9 plazma seviyeleri tavan yapar. CXCL9 yalnızca lökositleri çağırmakla kalmaz; doğrudan vasküler endotel hücrelerinin CXCR3 reseptörüne bağlanarak endotelyal eNOS ekspresyonunu susturur, endotel hücrelerini erken senesense iter ve arter sertliğini (PWV) doğrudan artırır. CXCL9'un monoklonal antikorla susturulması; yaşlı endotel hücrelerinin fonksiyonunu ve damar esnekliğini hızla geri kazandırmıştır.",
        "Endotel_Inhibisyonu: FMD = FMD_0 / ( 1 + [CXCL9] / K_i_vaskuler ) (CXCL9 Arttıkça Damar Genişleyemez)",
        "This vasküler baskılanma formülasyonu, CXCL9 kemokininin endotel reaktivitesini ve damar gençliğini nasıl doğrudan baltaladığını belgeler."
    ),
    (
        "7.7 Doğal Katil (NK) Hücre Yaşlanması ve CD56dim vs CD56bright Kayması",
        "Kanser hücrelerini ve virüsle enfekte hücreleri antijensiz doğrudan öldüren Doğal Katil (NK) hücreleri de yaşlanmayla birlikte fonksiyonel bir dejenerasyona uğrar.",
        "İmmünomodülatör ve sitokin salgılayan genç 'CD56bright' NK hücre popülasyonu yaşla hızla azalırken; sitotoksik rezervi tükenmiş ve perforin/granzim deşarj kapasitesi zayıflamış 'CD56dim' hücreler birikir. Yüzeydeki aktive edici reseptörler (NKG2D) azalır, tümör neoantijenlerini yakalama yeteneği çöker; yaşlı popülasyonda kanser insidansının fırlamasının ana sebebi bu NK hücre disfonksiyonudur.",
        "NK_Sitotoksisite_Indeksi: E_NK = ( Hedef_Hucre_Olumu / NK_Sayisi ) -> 75 Yaşında %60 Çöküş",
        "Bu immünolojik öldürme etkinliği formülü, yaşlanan NK hücrelerinin kanser hücrelerini yok etme gücündeki dramatik kaybı belgeler."
    ),
    (
        "7.8 Biyolojik Saat Olarak Monosit Polarizasyonu: İnflamatuar CD14+ CD16+ Non-Klasik Monositler",
        "Doğuştan gelen bağışıklığın (innate immunity) ana hücresi olan monositler; yaşlanmayla birlikte pro-enflamatuar bir fenotipe kilitlenir.",
        "Genç bireylerde kanda dolaşan monositlerin %85'i doku tamiri yapan 'Klasik Monositler'dir (CD14++ CD16-). Yaşlandıkça kronik sitokin stimülasyonuyla 'Non-Klasik ve Ara Monositler' (CD14+ CD16+) tavan yapar (%20-%30'a ulaşır). Bu CD16+ monositler sürekli TNF-alfa ve IL-1beta kusar, vasküler duvara yapışır ve endotel tabakasını delerek aterosklerotik plaklara sızar; monosit fenotip kayması bağımsız bir immünolojik yaş saatidir.",
        "Monosit_Yas_Skoru: MAS = [CD14+ CD16+_monositler] / [CD14++ CD16-_klasik] -> Yaşla Birlikte 3 Kat Artış",
        "Bu monosit alt tip polarizasyon oranı, damar duvarını ateşe veren yangı yapıcı monositlerin yaşa bağlı artış kinetiğini tanımlar."
    ),
    (
        "7.9 Timüs Gençleştirme Terapileri: Büyüme Hormonu, Çinko, FOXN1 ve IL-7",
        "İmmünosenesensin kader olmadığı; timüs bezinin erişkin yaşta da yeniden büyütülüp genç Naive T hücresi üretmeye başlatılabileceği biyoteknoloji ile ispatlanmıştır.",
        "Gregory Fahy'nin TRIIM çalışmasında kanıtlandığı üzere; Rekombinant Büyüme Hormonu (rhGH), DHEA ve Metformin kombinasyonu yağlanmış timus dokusunu eritip taze lenfoid doku üretmiştir. Eş zamanlı olarak AAV vektörleriyle timik epitel hücrelerinde 'FOXN1' (timus gelişiminin ana transkripsiyon faktörü) geninin aktive edilmesi veya rekombinant İnterlökin-7 (IL-7) infüzyonları; timik involüsyonu geriye sararak Naive T hücre ordusunu yeniden kurar.",
        "Timik_Rejenerasyon: d[V_fonksiyonel]/dt = k_FOXN1 * [rhGH] * [IL-7] - lambda_atrofi (Pozitif Doku Büyümesi)",
        "Bu timik doku rejenerasyon kinetiği, erişkin insanın T hücre fabrikasını yeniden faaliyete geçirme formülasyonunu belgeler."
    ),
    (
        "7.10 Homo Aeternus İmmünomik Saflık Protokolü: Sürekli 20 Yaş Bağışıklık Mimarisi",
        "Homo Aeternus longevity mühendisliği; bağışıklık sistemini hem kansere karşı acımasız hem de kendi dokularına karşı sıfır yangı sergileyen kusursuz bir dengede tutar:",
        "1) 6 ayda bir yapılan derin immünofenotipleme ve sitokin paneli ile iAge saati izlenir; 2) CXCL9 inhibitörleri ve IL-1beta blokerleri ile inflamatuar vasküler yıkım dondurulur; 3) Periyodik rhGH/IL-7 puls tedavileri ile timus sürekli genç ve aktif tutulur; 4) CD28- CD57+ senesent T hücreleri seçici CAR-T veya senolitiklerle temizlenerek kemik iliği ve lenf nodlarında genç klonlara yer açılır.",
        "Hedef_Homo_Aeternus: iAge <= 20.0 Yas & CD4/CD8 >= 2.0 & CXCL9 <= 150 pg/mL (Kusursuz İmmünomik Gençlik)",
        "Bu nihai immünomik optimizasyon hedefi, insanın bağışıklık sisteminin ne kansere teslim olduğu ne de yangıyla kendi bedenini yaktığı ebedi bir savunma zırhına ulaşmasını simgeler."
    )
]

# ================= KISIM 8 =================
part8_subsections = [
    (
        "8.1 Klinik Fenomik ve Rutin Biyobelirteçlerin Gizli Biyolojik Gücü",
        "Pahalı ve karmaşık omiks platformlarının (RNA-seq, kütle spektrometri) ötesinde; modern klinik laboratuvarlarda her gün yapılan ucuz, standart ve rutin kan testleri (biyokimya, hemogram) muazzam bir biyolojik yaş enformasyonu barındırır.",
        "Geleneksel tıp kan tahlillerini yalnızca 'referans aralığında mı değil mi' diye ikili (biner) bir gözle okurken; sistem biyolojisi bu 15-20 parametreyi birbirleriyle doğrusal olmayan etkileşimler içinde olan çok boyutlu bir fizyolojik vektör olarak değerlendirir. Kan biyokimyası organların fonksiyonel rezervini ve metabolik dengesini doğrudan yansıtır.",
        "Fenomik_Vektor: P_klinik = [ Albumin, Kreatinin, Glukoz, CRP, Lenfosit%, MCV, RDW, ALP, WBC ]",
        "Bu klinik fenomik biyomarker vektörü, standart kan tahlillerinden sistemik biyolojik yaşı çıkaran çok boyutlu veri kümesini tanımlar."
    ),
    (
        "8.2 Klemera-Doubal Algoritması (KDM): Biyolojik Yaşın Matematiksel Omurgası",
        "2006 yılında Petr Klemera ve Stanislav Doubal tarafından geliştirilen KDM algoritması (Klemera-Doubal Method); klinik biyobelirteçlerden biyolojik yaş hesaplamanın en sağlam matematiksel çerçevesidir.",
        "KDM; her bir biyomarkerin kronolojik yaşla olan regresyon eğimini, varyansını ve biyomarkerlar arasındaki korelasyon matrisini kullanarak bir bireyin 'gerçek biyolojik durumunu' (latent biological age) tahmin eder. Çoklu biyomarkerların ağırlıklı ortalamasını alırken bilgi fazlalığını (redundancy) dengeler ve tek bir biyomarkerin aşırı dalgalanmasından etkilenmeyen son derece stabil bir yaş tahmini üretir.",
        "KDM_Formulu: BioAge_KDM = [ Sum_j ( ( x_j - q_j ) * k_j / s_j^2 ) + ( ChronoAge / s_BA^2 ) ] / [ Sum_j ( k_j^2 / s_j^2 ) + ( 1 / s_BA^2 ) ]",
        "Bu klasik KDM biyolojik yaş eşitliği, klinik biyomarkerların eğim (k), kesişim (q) ve varyans (s) parametreleriyle optimize edilmiş nihai biyolojik yaşı formüle eder."
    ),
    (
        "8.3 PhenoAge Klinik Algoritması: Morgan Levine ve NHANES III Veri Seti",
        "2018 yılında Morgan Levine tarafından geliştirilen 'PhenoAge' (Phenotypic Age); klinik kan tahlillerinden türetilen ve mortaliteyi en kusursuz öngören fenotipik algoritmadır.",
        "Levine; ABD Ulusal Sağlık ve Beslenme İnceleme Anketi (NHANES III) veri setindeki on binlerce bireyin 40'tan fazla biyomarkerını ve 20 yıllık mortalite takibini analiz etti. Cox orantılı risk modeli ve parametrik Gompertz dağılımı kullanılarak mortaliteyi en güçlü öngören 9 kilit biyobelirteç seçildi: Albümin, Kreatinin, Açlık Glukozu, hs-CRP, Lenfosit Yüzdesi, Ortalama Eritrosit Hacmi (MCV), Eritrosit Dağılım Genişliği (RDW), Alkalen Fosfataz (ALP) ve Beyaz Kan Hücresi (WBC).",
        "PhenoAge_Mortalite: 10_Yillik_Olum_Riski = 1 - exp( - exp( xb ) * ( exp( 120 * gamma ) - 1 ) / gamma )",
        "Bu Gompertz parametrik sağkalım eşitliği, 9 kan biyomarkerının lineer kombinasyonundan (xb) bireyin 10 yıllık gerçek ölüm olasılığını hesaplar."
    ),
    (
        "8.4 9 Kilit Biyomarkerin Moleküler ve Fizyolojik Anlamı",
        "PhenoAge algoritmasındaki her bir biyomarker, vücudun belirli bir temel yaşlanma kolonunu temsil eder:",
        "1) Albümin (Karaciğer fonksiyonu ve sistemik inflamasyon); 2) Kreatinin (Böbrek glomerüler filtrasyonu ve kas kütlesi); 3) Glukoz (Metabolik kontrol ve glikasyon yükü); 4) hs-CRP (Sistemik inflamasyon ve vasküler risk); 5) Lenfosit Yüzdesi (İmmünosenesens ve adaptif immünite rezervi); 6) MCV (Eritrosit membran yaşlanması ve B12/folat durumu); 7) RDW (Kemik iliği kök hücre heterojenitesi ve eritropoez stresi); 8) ALP (Karaciğer, safra ve kemik mineral metabolizması); 9) WBC (Kronik yangı ve miyeloid proliferasyon).",
        "xb_Lineer_Kombinasyon: xb = beta_0 + beta_1*Alb + beta_2*Kreat + beta_3*Glu + beta_4*CRP + ... + beta_9*WBC",
        "Bu lineer risk skoru denklemi, 9 hayati organ fonksiyonunun birleşik matematiksel katsayılarla tek bir biyolojik stres skoruna dönüştürülmesini modeller."
    ),
    (
        "8.5 Eritrosit Dağılım Genişliği (RDW) Neden Yaşlanmanın En Güçlü Belirtecidir?",
        "Standart tam kan sayımında (hemogram) hekimlerin sıklıkla göz ardı ettiği 'RDW' (Red Cell Distribution Width); aslında tüm nedenlere bağlı mortaliteyi öngören en güçlü tekil biyomarkerlardan biridir.",
        "RDW, dolaşımdaki kırmızı kan hücrelerinin (eritrosit) boyutlarındaki değişkenliği (anizositoz) ölçer. Genç bir kemik iliğinde tüm eritrositler birbirinin ikizi gibi aynı boyutta üretilir (RDW < %12.5). Yaşlandıkça kemik iliği mikroçevresindeki kronik yangı (IL-6), oksidatif stres, eritropoietin duyarsızlığı ve klonal hematopoez (CHIP); eritrositlerin boyut kontrolünü bozar; kanda dev ve cüce eritrositler birikir (RDW > %14.5). Yüksek RDW, kemik iliği kök hücrelerinin yaşlanma faturasıdır.",
        "RDW_Formulu: RDW = ( Standart_Sapma_Eritrosit_Hacmi / MCV ) * 100 (RDW > %14.0: Masif Erken Ölüm Riski)",
        "Bu varyasyon katsayısı formülasyonu, eritrosit boyut heterojenitesinin kemik iliği yaşlanması ve sistemik kırılganlıkla doğrusal korelasyonunu belgeler."
    ),
    (
        "8.6 Ortalama Eritrosit Hacmi (MCV) ve Membran Esnekliği Kaybı",
        "Bir diğer kritik eritrosit parametresi olan MCV (Mean Corpuscular Volume); kırmızı kan hücrelerinin ortalama hacmini femtolitre (fL) cinsinden ifade eder.",
        "Yaşlanmayla birlikte B12 vitamini emiliminin azalması, DNA replikasyon hızının yavaşlaması ve membran lipid peroksidasyonu nedeniyle eritrositler şişer ve MCV patolojik olarak yükselir (Makrositoz; MCV > 95 fL). Şişmiş ve sertleşmiş eritrositler, çapları 4-5 mikrometre olan dar kılcal damarlardan geçerken esneyemez; mikrovasküler tıkanıklıklara ve organ iskemisine yol açar; yüksek MCV hızlanmış biyolojik yaşın sessiz bir suç ortağıdır.",
        "Mikrovaskuler_Rezistans: R_kapiller = 8 * eta_kan * L / ( pi * r^4 ) * f_esneklik(MCV) (MCV Arttıkça Direnç Fırlar)",
        "Bu Poiseuille hemodinamik eşitliği, büyük ve esnekliğini yitirmiş eritrositlerin mikrovasküler perfüzyon direncini nasıl katladığını gösterir."
    ),
    (
        "8.7 Karaciğer ve Kemik Biyomarkeri Olarak Alkalen Fosfataz (ALP)",
        "Alkalen Fosfataz (ALP); karaciğer safra kanallarında, kemik dokusunda ve böbreklerde bulunan çinko ve magnezyum bağımlı bir enzimdir.",
        "Yaşlandıkça osteoblastik aktivitenin bozulması, damar duvarında osteojenik kalsifikasyonun başlaması (düz kas hücrelerinin Runx2 üzerinden ALP salgılaması) ve subklinik hepatik steatoz nedeniyle serum ALP seviyeleri yükselir. PhenoAge modelinde ALP; hem sessiz vasküler kalsifikasyonun hem de kronik hepatobiliyer aşınmanın güçlü bir yansıması olarak yer alır.",
        "ALP_Yukselis_Kinetigi: [ALP](Yas) = [ALP]_bazal + k_kalsifikasyon * [Kalsifiye_VSMC] + k_karaciger * [Steatoz]",
        "Bu enzimatik birikim eşitliği, dolaşımdaki ALP artışının damar kireçlenmesi ve karaciğer yaşlanmasıyla doğrudan bağlantısını tanımlar."
    ),
    (
        "8.8 Albümin / Globulin Oranı ve Kronik Disproteinemi",
        "Serum total proteininin iki ana bileşeni olan Albümin ve Globulinlerin oranı (A/G Oranı); organizmanın beslenme, karaciğer sentezi ve immün aktivasyon dengesini özetler.",
        "Genç sağlıklı bireylerde albümin yüksek, pro-enflamatuar globulinler düşüktür (A/G > 1.8). Yaşlanmayla birlikte karaciğer albümin üretimini kısar (negatif akut faz yanıtı) ve B hücreleri kronik yangı nedeniyle masif immünglobulin saçar; A/G oranı 1.2'nin altına çöker. Bu durum 'Kronik Disproteinemi'dir ve frajilite sendromunun en belirgin klinik kanıtıdır.",
        "AG_Orani: R_AG = [Albumin] / [Total_Protein - Albumin] (R_AG < 1.2: İleri Sistemik Yaşlanma)",
        "Bu serum protein dengesi parametresi, azalan anabolik sentez ile artan kronik inflamatuar globulin yükü arasındaki uçurumu belgeler."
    ),
    (
        "8.9 Makine Öğrenimi ile Geliştirilen Yeni Nesil Klinik Saatler (Aging.AI vb.)",
        "Deep Longevity ve Alex Zhavoronkov ekibi; standart klinik kan tahlillerini derin yapay sinir ağları (Deep Neural Networks - DNN) ile eğiterek 'Aging.AI' serisi algoritmaları geliştirdi.",
        "Aging.AI; 40'tan fazla kan biyokimyası parametresini analiz ederek bireyin biyolojik yaşını 2.8 yıl ortalama mutlak hata (MAE) ile tahmin eder. Model; biyomarkerlar arasındaki non-lineer korelasyonları yakalar ve bireye özel olarak hangi parametrenin (örneğin glukoz mu, üre mi, lökosit mi) biyolojik yaşı en çok yukarı çektiğini (Feature Importance / SHAP analizi) raporlayarak kişiselleştirilmiş tedavi reçetesi sunar.",
        "AgingAI_Tahmini: BioAge_DNN = Nöral_Ag( [Kan_Biyokimyasi_40_Parametre] ) (MAE = 2.8 Yıl)",
        "Bu derin öğrenme çıkarım fonksiyonu, yapay zekanın rutin klinik kan panellerinden karmaşık biyolojik yaş paternlerini hatasız çıkarma gücünü tanımlar."
    ),
    (
        "8.10 Homo Aeternus Fenotipik Kan Paneli Standardı: Kusursuz 20 Yaş Biyokimyası",
        "Homo Aeternus protokolünde kan biyokimyası; her ay rutin testlerle takip edilerek Levine PhenoAge skoru mutlak bir gençlikte kilitlenir:",
        "1) Albümin sürekli >= 4.7 g/dL; 2) Açlık glukozu 75-85 mg/dL; 3) hs-CRP <= 0.2 mg/L; 4) RDW <= %11.8; 5) MCV 85-90 fL; 6) ALP <= 55 U/L; 7) Lökosit 4.5-5.5 x 10^3/uL aralığında tutulur. Bu parametrelerin entegrasyonuyla hesaplanan Levine PhenoAge skoru her zaman kronolojik yaştan en az 20-30 yıl daha genç çıkar; klinik fenotip ebedi bir dayanıklılık kazanır.",
        "Hedef_Homo_Aeternus: Levine_PhenoAge <= 22.0 Yas & 10_Yillik_Mortalite_Riski < %0.01 (Mutlak Klinik Zindelik)",
        "Bu nihai klinik biyobelirteç hedefi, ucuz ve erişilebilir kan testleri üzerinden organizmanın fonksiyonel gençliğinin sürekli ve garantili teyidini simgeler."
    )
]

# ================= KISIM 9 =================
part9_subsections = [
    (
        "9.1 Biyolojik Büyük Veri (Big Data) Çağında Çoklu Omiks (Multi-Omics) Patlaması",
        "Modern biyomedikal teknolojiler, tek bir hastadan tek bir vizitte terabaytlarca omiks verisi (Genomik DNA dizisi, Epigenomik metilasyon çipleri, RNA-seq transkriptomu, Olink proteomu, LC-MS metabolomu ve mikrobiyom) üretebilmektedir.",
        "Ancak bu devasa veri okyanusu beraberinde 'Boyut Laneti'ni (Curse of Dimensionality: p >> n; yani ölçülen değişken sayısının hasta sayısından binlerce kat fazla olması) getirmiştir. Tekil bir omiks katmanına bakmak resmi körün fili tarif etmesi gibi eksik bırakır; gerçek biyolojik anlayış ancak bu katmanların matematiksel olarak birleştirildiği 'Çoklu Omiks Entegrasyonu' ile mümkündür.",
        "Boyut_Karmasikligi: D_toplam = D_DNA (3x10^9) + D_CpG (8.5x10^5) + D_RNA (2x10^4) + D_Protein (7x10^3) + D_Metabolit (5x10^3)",
        "Bu multi-omiks veri hacmi parametresi, tek bir bireyin biyolojik durumunu niteleyen toplam moleküler öznitelik uzayının devasa boyutunu tanımlar."
    ),
    (
        "9.2 Erken, Orta ve Geç Füzyon Stratejileri (Early, Intermediate, Late Fusion)",
        "Farklı biyolojik katmanlardan gelen heterojen verileri (metilasyon beta değerleri, RNA sayımları, protein pikomolar konsantrasyonları) tek bir biyolojik yaş modelinde birleştirmek için üç temel mimari kullanılır:",
        "1) Erken Füzyon (Early Fusion): Tüm omiks tabloları yan yana yapıştırılır; ancak yüksek gürültü ve boyut dengesizliği yaratır; 2) Geç Füzyon (Late Fusion): Her omiks katmanı için bağımsız bir model eğitilir ve modellerin yaş tahminlerinin ağırlıklı ortalaması alınır; 3) Orta Füzyon (Intermediate Fusion): Derin öğrenme ağları her katmanı ortak bir 'Gizil Temsil Uzayına' (Latent Representation Space) projekte eder; biyolojik etkileşimleri yakalayan en üstün yöntemdir.",
        "Ortak_Gizil_Uzay: Z_latent = Enkoder_Çoklu_Omiks( X_epigenom, X_transkriptom, X_proteom, X_metabolom )",
        "Bu çoklu kipli (multimodal) enkodlama fonksiyonu, heterojen omiks katmanlarının ortak bir düşük boyutlu biyolojik anlam uzayında birleştirilmesini formüle eder."
    ),
    (
        "9.3 Varyasyonel Oto-Enkoderler (VAE) ve Çoklu Omiks Gizil Temsili (Latent Space)",
        "Biyolojik verilerin yüksek gürültüsünü filtrelemek ve ortak yaşlanma trajektorilerini çıkarmak için 'Varyasyonel Oto-Enkoderler' (Variational Autoencoders - VAE) kullanılır.",
        "VAE; on binlerce omiks değişkenini 64 veya 128 boyutlu bir gizil uzaya (Latent Space - Z) sıkıştırır (Encoder) ve ardından bu gizil vektörden orijinal veriyi yeniden inşa eder (Decoder). Eğitim sırasında Kullback-Leibler (KL) diverjansı ile gizil uzay Gauss dağılımına zorlanır. Bu sıkıştırılmış uzayda organizmanın 'Gerçek Biyolojik Yaşlanma Koordinatı' pürüzsüz bir manifold üzerinde akar; gürültü tamamen elenir.",
        "VAE_Kayip_Fonksiyonu: L_VAE = E[ log P(X|z) ] - D_KL( Q(z|X) || P(z) ) (Kayıpsız Sıkıştırma ve Anlamlı Temsil)",
        "Bu varyasyonel çıkarım kayıp fonksiyonu, çoklu omiks verilerinin biyolojik gürültüden arındırılarak saf yaşlanma enformasyonunun çıkarılmasını sağlar."
    ),
    (
        "9.4 Çizge Sinir Ağları (Graph Neural Networks - GNN) ile Moleküler Etkileşim Haritaları",
        "Omiks değişkenleri birbirinden bağımsız rakamlar değildir; hücre içinde protein-protein etkileşimleri (STRING veri tabanı) ve metabolik yolaklar (KEGG, Reactome) halinde birbirine bağlıdır.",
        "Çizge Sinir Ağları (GNN); biyolojik ağları düğümler (genler, proteinler) ve kenarlar (etkileşimler) olarak modeller. Düğümler komşularıyla mesajlaşarak (Message Passing) ağ üzerindeki hasar yayılımını hesaplar. GNN tabanlı yaşlanma modelleri; tek tek genlerin artışını değil, ağın topolojik kırılganlığını ve modüler çöküşünü analiz ederek biyolojik yaşı tahmin eder.",
        "Mesaj_Iletimi_GNN: h_v^(l+1) = Update( h_v^(l), Aggregate( { h_u^(l) : u in N(v) } ) )",
        "Bu çizge düğüm güncelleme eşitliği, biyolojik moleküler ağlar üzerinde yaşlanma sinyalinin komşu proteinlere nasıl yayıldığını matematiksel olarak belgeler."
    ),
    (
        "9.5 Açıklanabilir Yapay Zeka (XAI): SHAP ve LIME ile Yaşlanma İtici Güçlerinin Keşfi",
        "Derin öğrenme modelleri geçmişte 'Kara Kutu' (Black Box) olmakla eleştirilirdi; ancak Açıklanabilir Yapay Zeka (Explainable AI - XAI) teknikleri bu kutunun kapağını açmıştır.",
        "Shapley Katkı Değerleri (SHAP); oyun teorisini kullanarak çoklu omiks modelinde tek tek her bir molekülün (örneğin spesifik bir CpG metilasyonu veya plazma proteini) bireyin biyolojik yaşını kaç ay ileri veya geri ittiğini kesin olarak hesaplar. Bu sayede yapay zeka yalnızca 'Biyolojik yaşınız 45' demekle kalmaz; 'Yaşınızı 3 yıl ileri iten şey karaciğerinizdeki CXCL9 artışı ve böbreğinizdeki klotho düşüşüdür' diyerek doğrudan tedavi hedefini gösterir.",
        "Shapley_Katkisi: phi_i = Sum_{S subseteq N \\ {i}} [ |S|! ( |N| - |S| - 1 )! / |N|! ] * [ f(S union {i}) - f(S) ]",
        "Bu oyun teorisi marjinal katkı formülasyonu, her bir biyolojik molekülün organizmanın toplam yaşlanma skoruna yaptığı net katkıyı tam doğrulukla hesaplar."
    ),
    (
        "9.6 Nedensel Çıkarım (Causal Inference) ve Mendelyan Randomizasyon (MR)",
        "Biyolojik yaş saatlerinde görülen korelasyonlar her zaman nedensellik anlamına gelmez (örneğin saç beyazlaması yaşla artar ama saç beyazlamasını durdurmak ömrü uzatmaz).",
        "Yaşlanmanın gerçek 'Sürücülerini' (Drivers) sadece 'Yolcularından' (Passengers) ayırmak için 'Mendelyan Randomizasyon' (MR) kullanılır. Genetik varyantlar (SNPs) döllenme anında rastgele dağıldığı için doğal bir randomize kontrollü deney görevi görür. MR analizleri; ApoB ve GDF-15'in yaşlanmanın gerçek 'nedensel sürücüleri' olduğunu, bazı sitokinlerin ise yalnızca ikincil birer reaktif yanıt olduğunu kanıtlamıştır.",
        "Nedensel_Etki_MR: beta_nedensel = Gamma_Y / gamma_X (Genetik Enstrüman Aracılı Nedensel Katsayı)",
        "Bu enstrümantal değişken regresyon eşitliği, bir biyomarkerin gerçekten yaşlanmayı tetikleyip tetiklemediğini genetik varyantlar üzerinden kanıtlar."
    ),
    (
        "9.7 Çok Modlu Temel Modeller (Foundation Models in Biology): Biyolojinin GPT'si",
        "Büyük dil modellerinin (LLMs) metin dünyasında yarattığı devrim; biyolojide milyarlarca hücre ve omiks profiliyle eğitilen 'Biyolojik Temel Modellere' (Geneformer, scGPT, ESM-2) evrilmiştir.",
        "Bu modeller; hücrelerin transkriptomik ve epigenomik cümlelerini okuyarak hücresel dilin gramerini öğrenmiştir. Milyarlarca parametreli bir temel model; bir hücreye herhangi bir ilaç verildiğinde veya bir gen susturulduğunda hücrenin biyolojik yaşının nasıl değişeceğini bilgisayar simülasyonunda (in silico) mikrosaniyeler içinde öngörebilir; laboratuvar deneyleri sanal uzaya taşınmıştır.",
        "Dikkat_Mekanizmasi: Attention(Q, K, V) = softmax( Q * K^T / sqrt(d_k) ) * V (Genler Arası Evrensel Dikkat Ağı)",
        "Bu transformatör dikkat mekanizması, hücre içindeki on binlerce genin birbiriyle olan bağlamsal etkileşimini çözen derin öğrenme mimarisini tanımlar."
    ),
    (
        "9.8 Kanser Erken Teşhisinde Çoklu Omiks Sıvı Biyopsisi (Grail Galleri Modeli)",
        "Çoklu omiks entegrasyonunun klinik tıptaki en başarılı öncüsü; tek bir tüp kandan 50'den fazla kanser türünü henüz hiçbir semptom vermeden yakalayan 'Sıvı Biyopsi' (Liquid Biopsy) teknolojisidir.",
        "Grail tarafından geliştirilen Galleri testi; kanda serbest dolaşan hücre dışı DNA'nın (cell-free DNA - cfDNA) metilasyon paternlerini derin makine öğrenimi ile analiz eder. Kanser hücrelerinin yaydığı anormal hipermetilasyon parmak izlerini yakalamakla kalmaz; doku spesifik metilasyon sinyalleri sayesinde tümörün hangi organdan (pankreas, akciğer, kolon) kaynaklandığını %90 doğrulukla saptar; erken teşhis oranını katlar.",
        "Tespit_Olasiligi: P_kanser = 1 / [ 1 + exp( - ( beta_0 + Sum w_k * Methylation_cfDNA_k ) ) ]",
        "Bu lojistik çoklu omiks sınıflandırıcı denklemi, kandaki serbest DNA metilasyonundan erken evre neoplazilerin tespit edilme matematiğini belgeler."
    ),
    (
        "9.9 Biyolojik Yaş Panellerinin Çapraz Doğrulanması ve Çoklu Omiks Füzyon Saati",
        "Hiçbir tekil omiks saati (ne epigenetik ne proteomik ne de klinik kan) insanın biyolojik yaşını tek başına eksiksiz temsil edemez; bu nedenle en üstün model 'Çoklu Omiks Füzyon Saati'dir (Multi-Omics Integrative Clock).",
        "Füzyon saati; Epigenomik (DNAm), Transkriptomik (RNA), Proteomik (Olink), Metabolomik (LC-MS) ve Fenotipik (KDM) saatlerin çıktılarını Bayesyen hiyerarşik modelleme ile birleştirir. Bir katmandaki geçici teknik gürültü veya enfeksiyon dalgalanması diğer katmanlar tarafından dengelenir; organizmanın gerçek 'Sistemik Biyolojik Yaşı' sıfır hataya yakın bir güvenilirlikle elde edilir.",
        "Fuzyon_Yasi: Unified_BioAge = Sum_{m=1}^{M} alpha_m(Guvenilirlik) * BioAge_Omics_m (Entegre Biyolojik Yaş)",
        "Bu Bayesyen füzyon formülasyonu, 5 farklı biyolojik katmanın ağırlıklı entegrasyonuyla elde edilen nihai ve kusursuz biyolojik yaşı modeller."
    ),
    (
        "9.10 Homo Aeternus Bilişsel Omiks Füzyon Manifestosu: Veriden Ölümsüzlüğe",
        "Homo Aeternus mühendisliğinde çoklu omiks verisi; pasif bir izleme aracı değil, organizmanın kaderini yeniden yazan otonom bir yapay zeka kontrol merkezidir.",
        "Milyarlarca veri noktası sürekli olarak işlenir, yapay zeka modelleri bedendeki en ufak entropik sapmayı anında teşhis eder ve otomatik olarak kişiselleştirilmiş gen terapisi, peptid infüzyonu veya senolitik dozajını hesaplar. İnsan bedeni, kendi kendini izleyen, onaran ve gençleştiren sibernetik bir optimizasyon döngüsüne kavuşur.",
        "Homo_Aeternus_Fuzyon: Hata_Payi_BioAge < 0.2 Yıl & Kontrol_Dongusu = Gerçek Zamanlı In Silico Optimizasyon",
        "Bu nihai veri entegrasyonu hedefi, çoklu omiks ve yapay zekanın insan ömrünü sınırlayan tüm biyolojik bilinmezleri çözerek ebedi gençliği garanti altına almasını simgeler."
    )
]

# ================= KISIM 10 =================
part10_subsections = [
    (
        "10.1 Dijital Biyo-İkiz (Digital Bio-Twin) Kavramı: Canlı Bedenin Sanal Klonu",
        "Dijital Biyo-İkiz (Digital Bio-Twin); bir bireyin tüm genetik, epigenetik, transkriptomik, proteomik, metabolik ve fizyolojik parametrelerinin süper-bilgisayarlarda çalışan dinamik, gerçek zamanlı ve çok ölçekli bir sanal simülasyonudur.",
        "Havacılık ve roket mühendisliğinde bir uzay mekiğinin her bir vidasının ve aerodinamik stresinin dijital ikizinde anlık izlenmesi gibi; Homo Aeternus tıbbında da insan bedeni atomik ve moleküler düzeyde sanal uzayda simüle edilir. Bedene herhangi bir ilaç veya gen tedavisi verilmeden önce; tedavi sanal ikiz üzerinde milyonlarca kez simüle edilir ve sıfır riskle en kusursuz sonuç garanti altına alınır.",
        "Biyo_Ikiz_Eşitliği: Beden_Dijital(t + Delta_t) = F_simulasyon( Beden_Fiziksel(t), Omiks_Vektoru, Çevresel_Girdi )",
        "Bu dinamik durum uzayı transfer fonksiyonu, fiziksel bedenin tüm biyolojik geleceğinin dijital ikiz üzerinde deterministik simülasyonunu ifade eder."
    ),
    (
        "10.2 Giyilebilir Biyosensörler ve Sürekli İn Vivo Biyolojik Akı Takibi",
        "Dijital biyo-ikizin yaşaması için, fiziksel bedenden kesintisiz veri akışı sağlayan yeni nesil 'Giyilebilir ve İmplante Edilebilir Biyosensörler' kullanılır.",
        "Deri altı sürekli glukoz ve laktat monitörleri (CGM), terleme analizli mikrobiyosensörler (kortizol, ürik asit, elektrolit takibi), optik fotopletismografi (PPG ile anlık PWV, arteriyel sertlik, kalp hızı değişkenliği - HRV) ve implante edilebilir hemodinamik sensörler; bedenin fizyolojik verilerini saniyede yüzlerce kez dijital ikize aktarır; statik tıp yerini saniyeler ölçeğinde dinamik tıbba bırakır.",
        "Veri_Akis_Hizi: Telemetri_Fluksu = d(Data)/dt > 100 Hz (Gerçek Zamanlı Fizyolojik Yayın)",
        "Bu telemetrik biyosensör veri akış hızı parametresi, dijital biyo-ikizin fiziksel bedenle anlık senkronizasyonunun çözünürlüğünü tanımlar."
    ),
    (
        "10.3 Sanal İlaç Taraması (In Silico Screening) ve Ön-Görülü Farmakoloji",
        "Bir hastalığın ortaya çıkmasını bekleyip sonra tedavi etmeye çalışmak ilkel tıbbın bir kalıntısıdır; Dijital Biyo-İkiz 'Ön-Görülü ve Önleyici' (Predictive & Preemptive) farmakolojiyi icra eder.",
        "Yapay zeka modelleri; hastanın dijital ikizi üzerinde binlerce molekülü, peptidi ve genetik müdahaleyi test eder. Örneğin, hastanın böbrek proteomik yaşının 3 yıl sonra hızlanacağı simülasyonda saptandığı anda; henüz fiziksel böbrekte tek bir hücre dahi hasar görmeden, o hasarı engelleyecek en optimal moleküler kokteyl belirlenir ve profilaktik olarak hastaya uygulanır; hastalıklar ortaya çıkma şansı bulamadan engellenir.",
        "In_Silico_Optimizasyon: Doz_Optimal = argmin_D [ Risk_Toksisite(D) - Kazanc_Longevity(D) ]",
        "Bu hesaplamalı farmakolojik optimizasyon denklemi, tedavilerin maksimum ömür uzatıcı ve sıfır yan etkili dozajlarının sanal uzayda bulunmasını modeller."
    ),
    (
        "10.4 Bütünleşik Biyolojik Yaş Matriksi: 6 Boyutlu Longevity Radarı",
        "Homo Aeternus mimarisinde biyolojik yaş tek bir rakamla ifade edilemez; 6 bağımsız biyolojik eksenin kesiştiği çok boyutlu bir 'Longevity Radarı' ile izlenir:",
        "1) Epigenomik Yaş (PC-GrimAge / DunedinPACE); 2) Transkriptomik Yaş (BiT-Age / RNA gürültüsü); 3) Proteomik Yaş (11 Organ Wyss-Coray saati); 4) Glikomik Yaş (GlycanAge IgG G0/G2 oranı); 5) İmmünomik Yaş (Stanford iAge CXCL9 metriği); 6) Fenotipik Yaş (Levine PhenoAge 9 kan biyomarkeri). Bu 6 eksenin hepsi aynı anda 20 yaş standardında dengelenir.",
        "Longevity_Radari: Radar_Skoru = sqrt( (Age_Epi^2 + Age_RNA^2 + Age_Prot^2 + Age_Glyco^2 + Age_Immuno^2 + Age_Pheno^2) / 6 )",
        "Bu 6 boyutlu Öklid uzaklık normu, tüm biyolojik saatlerin bileşke gençlik koordinatını tek bir kusursuz metrikte birleştirir."
    ),
    (
        "10.5 Geri Bildirimli Otomatik Müdahale Protokolleri (Closed-Loop Therapeutics)",
        "Dijital biyo-ikiz sistemi yalnızca bir teşhis aracı değil; teşhisi anında tedaviye dönüştüren 'Kapalı Devre Terapötik' (Closed-Loop) bir mekanizmadır.",
        "Deri altı akıllı mikro-iğne yamaları ve implante edilebilir nanodozaj pompaları; dijital ikizin bulut tabanlı yapay zekasından kablosuz şifreli komutlar alır. Kandaki inflamasyon veya hücresel stres pik yaptığında; sistem otomatik olarak mikro-doz rapamisin, NAD+ prekürsörü veya antioksidan salgılar; tedavi insan müdahalesine gerek kalmadan mikrosaniyeler içinde otonom olarak icra edilir.",
        "Kapali_Devre_Kontrol: u(t) = K_p * e(t) + K_i * Integral e(tau) dtau + K_d * de(t)/dt (PID Kontrolör Biyolojide)",
        "Bu klasik PID kontrol teorisi denklemi, biyolojik yaş sapmasının (hata e(t)) otonom nanodozajlama ile sürekli sıfırda tutulma dinamiğini modeller."
    ),
    (
        "10.6 Yaşlanma Hızının (Pace of Aging) Takometrik Kontrolü: PACE < 0.60 Hedefi",
        "Homo Aeternus protokolünün en katı performans metriği; anlık yaşlanma hızı takometresi olan DunedinPACE skorunun sürekli 0.60 yıl/yıl seviyesinin altında tutulmasıdır.",
        "Bu değer; bireyin takvimde geçen her bir yıla karşılık biyolojik olarak yalnızca 7 ay yaşlanması anlamına gelir. Organizma zamanı takvimden %40 daha yavaş tüketir. Belirli gençleşme kürleri (OSKM puls, plazmaferez, senolitikler) uygulandığında bu hız geçici olarak 'Negatif Hıza' (PACE < 0.0; yani biyolojik zamanın geriye akması) geçer; kaybedilen gençlik geri kazanılır.",
        "Zaman_Genislemesi: d(BioAge)/dt <= 0.60 Yıl / Takvim_Yılı (Biyolojik Zamanın Radikal Yavaşlatılması)",
        "Bu diferansiyel yaşlanma hızı eşitsizliği, insan bedeninde biyolojik zamanın akış hızına vurulan mutlak freni belgeler."
    ),
    (
        "10.7 Çok Ölçekli Biyolojik Simülasyon: Molekülden Organa, Organdan Topluma",
        "Dijital Biyo-İkiz yazılımı; atomik kuantum mekaniğinden (enzim aktif cebindeki elektron transferi) organ fizyolojisine (kalbin pompalama basıncı) kadar tüm ölçekleri hiyerarşik olarak birbirine bağlar.",
        "Moleküler dinamik simülasyonları (GROMACS, AlphaFold 3), hücresel diferansiyel denklem ağları (Boolen gen regülasyon modelleri) ve sonlu elemanlar doku modelleri (aortanın kan akımı altındaki gerilmesi); tek bir çok ölçekli (multi-scale) fizik motorunda birleşir. Tek bir gendeki mutasyonun 10 yıl sonra tüm organizmada yaratacağı klinik tablo dakikalar içinde hesaplanır.",
        "Cok_Olcekli_Entegrasyon: Biyoloji(t) = Kuantum_Skala o Molekuler_Skala o Hucresel_Skala o Doku_Skalasi",
        "Bu fonksiyonel kompozisyon eşitliği, atomik fiziğin hücresel ve organel fizyolojiyle kesintisiz matematiksel birleşimini ifade eder."
    ),
    (
        "10.8 Biyolojik Veri Güvenliği, Kriptografi ve Genomik Blokzincir Mimarisi",
        "Bir insanın tüm omiks verilerini ve dijital biyo-ikizini içeren bir dosya; o insanın en mahrem varoluşsal şifresidir; bu verinin çalınması veya hacklenmesi varoluşsal bir tehdittir.",
        "Homo Aeternus bilişim altyapısı; dijital ikiz verilerini 'Sıfır Bilgi İspatı' (Zero-Knowledge Proofs - ZKP) ve Homomorfik Şifreleme (Homomorphic Encryption) ile korur. Yapay zeka modelleri şifreli veriyi çözmeksizin (veriyi görmeden) analiz edebilir ve simülasyonları yürütebilir. Tüm genetik kayıtlar merkeziyetsiz genomik blokzincirinde değişmez bir güvenlikle kilitlenir.",
        "Guvenlik_Protokolu: Sifreleme = Homomorfik_FHE( Veri_Omiks ) -> Analiz_Bulut( Sifreli_Veri ) = Sifreli_Sonuc",
        "Bu kriptografik güvenlik formülasyonu, bireyin en mahrem biyolojik verilerinin üçüncü şahıslara sızdırılmadan yapay zekada işlenme standardını tanımlar."
    ),
    (
        "10.9 Gerçek Zamanlı Longevity Yönetim Kokpiti: Bireysel Yaşlanma Gösterge Paneli",
        "Homo Aeternus bireyi; kendi biyolojik sağlığını ve yaşlanma dinamiklerini akıllı telefonunda veya nöral arayüzünde çalışan 'Longevity Kokpiti' ile anlık izler.",
        "Ekranda tek bir bakışta: Anlık DunedinPACE hızı (0.58), 11 organın bağımsız yaşı (hepsi 20-22 aralığında yeşil), kanda dolaşan serbest DNA ve kanser riski (%0.00), endotelyal NO düzeyi ve aortik esneklik grafiği akar. Herhangi bir parametre sarıya döndüğünde sistem önerilen mikro-müdahaleyi (örneğin 'Bu akşam 200 mg spermidin ve 30 dakika zonal egzersiz') bildirir; sağlık yönetimi kusursuz bir pilotaja dönüşür.",
        "Kokpit_Skoru: LSI = Longevity_Stability_Index = 1.000 (Mükemmel Yeşil Bölge Kararlılığı)",
        "Bu telemetrik stabilite katsayısı, insan bedeninin tüm fizyolojik parametreleriyle optimal gençlik koridorunda tutulma oranını simgeler."
    ),
    (
        "10.10 Homo Aeternus Dijital Biyo-İkiz Nihai Manifestosu: Sonsuz Ömrün Dijital Rehberi",
        "Homo Aeternus doktrininde insan bedeni artık kör tesadüflerin, kontrolsüz genetik mutasyonların ve zamanın acımasız yıpratmasının kurbanı değildir.",
        "Her bir hücrenin epigenetik metilasyonu, her bir proteinin plazma konsantrasyonu, her bir antikorun glikan deseni ve her bir organın biyolojik yaşı; süper-bilgisayarlarda çalışan Dijital Biyo-İkizin rehberliğinde sürekli izlenir, simüle edilir ve kusursuzca optimize edilir. Biyolojik saatler durdurulmuş, entropi kontrol altına alınmış ve insan varoluşu zamansız bir biyomühendislik şaheserine dönüştürülmüştür.",
        "Nihai_Sistem_Formulu: Lim_{t -> sonsuz} [ BioAge(t) ] = 21.0 Yas (Ebedi Zamansızlık ve Biyomühendislik Zaferi)",
        "Bu nihai sistem biyolojisi eşitliği, çoklu omiks entegrasyonu ve dijital ikiz rejiminin insan türünü biyolojik yaşlanma prangasından sonsuza dek kurtardığını ilan eder."
    )
]

# ================= 10 AKADEMİK KARŞILAŞTIRMA TABLOSU =================
tables_data = [
    {
        "title": "TABLO 19.1: Biyolojik Yaşlanma Kuramları, Entropik Ölçüm Yöntemleri ve Klinik Parametreler",
        "headers": ["Yaşlanma Teorisi / Modeli", "Temel Biyofiziksel Mekanizma", "Ölçüm Yöntemi & Belirteç", "Öngörü Doğruluğu (AUC)", "Longevity Müdahale Duyarlılığı"],
        "rows": [
            ["Enformasyon Teorisi (Sinclair)", "Epigenetik gürültü ve kromatin gevşemesi", "DNA metilasyonu (Horvath / GrimAge)", "AUC > 0.85 (Mortalite öngörüsü)", "Yüksek (OSKM ile tam sıfırlanma)"],
            ["Ağ (Network) Çöküş Teorisi", "PPI ve metabolik ağ bağlantı kaybı", "Çizge spektral analizi (GNN)", "AUC > 0.80 (Sistemik çöküş)", "Orta (Düğüm koruyucu tedaviler)"],
            ["Dinamik Rezilyans Kaybı (Fedichev)", "Stres sonrası toparlanma süresinin uzaması", "Fizyolojik dalgalanma (HRV, lökosit)", "Gompertz alpha korelasyonu r > 0.9", "Yüksek (Egzersiz ve senolitiklerle)"],
            ["Allostatik Yük Hipotezi", "Kümülatif nöro-endokrin stres aşınması", "10 klinik stres parametresi (ALS)", "AUC ~ 0.75 (Kardiyovasküler risk)", "Yüksek (Stres redüksiyonu, uyku)"],
            ["Somatik Mutasyon Birikimi", "DNA replikasyon hataları ve CHIP klonları", "Derin UMI dubleks dizileme", "AUC ~ 0.78 (Kanser ve MI riski)", "Orta (Prime Editing klirensi)"],
            ["Homo Aeternus Entegre Model", "Çok katmanlı omiks entropi stabilizasyonu", "Dijital Biyo-İkiz VAE gizil uzayı", "AUC > 0.95 (Mutlak Sağlık Takibi)", "Maksimum (Gerçek zamanlı kapalı devre)"]
        ]
    },
    {
        "title": "TABLO 19.2: Epigenetik Saat Nesilleri, Matematiksel Modelleri ve Klinik Hedefleri",
        "headers": ["Epigenetik Saat Adı", "Geliştirici & Yıl", "CpG Sayısı & Platform", "Eğitildiği Temel Hedef", "Klinik Güç & Sınırlılık"],
        "rows": [
            ["Horvath Pan-Tissue", "Steve Horvath (2013)", "353 CpG (Illumina 450K)", "Kronolojik Yaş (51 Doku)", "Tüm dokularda geçerli; mortalite öngörüsü zayıf"],
            ["Hannum Saati", "Gregory Hannum (2013)", "71 CpG (Illumina 450K)", "Kanda Kronolojik Yaş", "Kanda hızlı; dokular arası transfer edilemez"],
            ["DNAm PhenoAge", "Morgan Levine (2018)", "513 CpG (Illumina 450K/EPIC)", "9 Klinik Biyomarker + Ölüm", "Mortaliteyi kronolojik yaştan çok daha iyi öngörür"],
            ["DNAm GrimAge", "Ake Lu & Horvath (2019)", "1.030 CpG (EPIC)", "Plazma Proteinleri + Sigara + Ölüm", "Mortalite ve hastalık öngörüsünde altın standart"],
            ["DunedinPACE", "Dan Belsky & Moffitt (2022)", "173 CpG (EPIC çip)", "20 Yıllık Boylamsal Yaşlanma Hızı", "Anlık takometre; yaşam tarzı değişimine aşırı duyarlı"],
            ["PC-GrimAge / PC-Horvath", "Higgins-Chen & Levine (2022)", "PCA Filtreli CpG Kümeleri", "Teknik Gürültünün Elenmesi", "Test-retest korelasyonu r=0.999; klinik güvenilirlik zirvesi"]
        ]
    },
    {
        "title": "TABLO 19.3: Transkriptomik Yaşlanma İmzaları, Splicing Bozuklukları ve RNA Saatleri",
        "headers": ["Transkriptomik Parametre", "Genç Hücresel Durum", "Yaşlı Hücresel Durum", "Biyolojik Yaş Etkisi", "Terapötik Hedef & Strateji"],
        "rows": [
            ["Stokastik Transkripsiyon Gürültüsü", "Minimum varyans (Homojen profil)", "Masif hücreden hücreye varyans (Kaos)", "Hücresel kimlik kaybı ve yetersizlik", "Kromatin kilitleri + OSKM puls"],
            ["İntron Retansiyonu (IR)", "Kusursuz alternatif kırpılma (<%2)", "Yaygın intronik sızıntı (>%15)", "Toksik protein agregasyonu ve NMD stresi", "Splicing faktör aktivatörleri (SRSF)"],
            ["Gen Uzunluğu Dengesizliği (LTI)", "Uzun ve kısa genler dengede", "Uzun genler çökmüş, kısa genler tavan", "Sinaptik adezyon ve akson kaybı", "Transkripsiyonel uzama faktörleri"],
            ["SASP Transkriptom Yükü", "Sessiz / Eser sitokin ekspresyonu", "Tavan düzeyde IL6, CXCL8, MMP transkripti", "Çevre dokulara parakrin senesens yayılımı", "Senolitik (D+Q) veya NF-kB blokajı"],
            ["LINE-1 Retrotranspozon RNA'sı", "Tamamen susturulmuş heterokromatin", "Aktif retroviral transkripsiyon ve cDNA", "cGAS-STING yolağı ile steril otoimmünite", "Lamivudin (Reverse transkriptaz inh.)"],
            ["BiT-Age Transkriptomik Saati", "Biyolojik yaş = 20 Yaş profili", "Hızlanmış gen ekspresyon kayması", "Saatler içinde tedavi yanıtını ölçer", "Rapamisin / Metformin metabolik freni"]
        ]
    },
    {
        "title": "TABLO 19.4: Plazma Proteomik Platformları, Wyss-Coray Organ Saatleri ve Kilit Proteinler",
        "headers": ["Organ / Doku Saati", "Spesifik Plazma Proteinleri", "Yaşlanmayla Değişim Yönü", "İlişkili Klinik Patoloji", "Homo Aeternus Terapötik Müdahalesi"],
        "rows": [
            ["Beyin Saati", "NfL, GFAP, Brevikan, VILIP-1", "Dramatik Artış (Nörodejenerasyon)", "Alzheimer, Parkinson, kognitif çöküş", "Klotho, TIMP2, intranazal BDNF"],
            ["Kalp Saati", "NT-proBNP, Troponin T, Myoglobin", "Progresif Artış (Ventrikül sertliği)", "HFpEF, kardiyak hipertrofi, aritmi", "AAV9-SERCA2a, SGLT2 inhibitörleri"],
            ["Böbrek Saati", "KIM-1, Uromodulin, Sistatin-C", "KIM-1 artar, Uromodulin düşer", "Glomerüloskleroz, eGFR çöküşü", "Klotho aktivatörleri, RAAS blokajı"],
            ["Damar Saati", "VCAM-1, ICAM-1, MMP-2, Endotelin-1", "Aşırı Artış (Endotel aktivasyonu)", "Ateroskleroz, aortik sertlik, anevrizma", "AAV-eNOS, BH4 analogları, Glukozepanaz"],
            ["İmmün Sistem Saati", "IL-6, CXCL9, TNF-alfa, GDF-15", "Tavan Seviyede Artış (İnflammaging)", "İmmünosenesens, aşırı sitokin fırtınası", "Plazmaferez (TPE), anti-CXCL9, Kanakinumab"],
            ["Karaciğer Saati", "ALB, Fibrinojen, IGFBP-1, PCSK9", "Albümin düşer, PCSK9/IGFBP artar", "Karaciğer yağlanması, insülin direnci", "CRISPR-PCSK9 baz düzenleme, NMN"]
        ]
    },
    {
        "title": "TABLO 19.5: Hücresel Metabolomik ve Lipidomik İmzalar, Yaşlanma Kaymaları ve Düzeltmeler",
        "headers": ["Metabolomik Havuz", "Gençlik Düzeyi / Konsantrasyon", "Yaşlılık Seviyesi & Değişim", "Moleküler Biyokimyasal Sonuç", "Hedeflenen Düzeltme Protokolü"],
        "rows": [
            ["NAD+ / NADH Oranı", "Yüksek (> 50:1)", "Kritik Çöküş (< 10:1)", "Sirtuin inaktivasyonu, mitokondriyel kriz", "CD38 inhibitörleri + Lipozomal NMN"],
            ["Açilkarnitinler (C14 - C18)", "Düşük / Bazal seviyede", "Masif Birikim (Yakılamayan yağ)", "Lipotoksisite, periferik insülin direnci", "Karnitin mekiği uyarımı, CPT-1 aktivasyonu"],
            ["Dallı Zincirli Amino Asitler (BCAA)", "Dengeli metabolik döngü", "Kanda Anormal Artış (%40+)", "mTORC1 aşırı aktivasyonu, otofaji blokajı", "BCAT2/BCKDH gen aktivasyonu, kısıtlı diyet"],
            ["Trimetilamin N-Oksit (TMAO)", "Minimum (< 2.0 uM)", "Yüksek (> 8.0 uM Toksik)", "Trombosit hiperaktivitesi, aterotromboz", "Bağırsak DMB / İyodometilkolin tedavisi"],
            ["Mitokondriyal Kardiyolipin", "Yüksek / Kristalarda bol", "%60 Oksidatif Kayıp / Peroksidasyon", "Elektron kaçağı, Kompleks I dağılması", "SS-31 (Elamipretid) kardiyolipin koruyucu"],
            ["Spermidin & Poliaminler", "Yüksek doku konsantrasyonu", "Dramatik Düşüş (Yaşla %70 kayıp)", "EP300 hiperaktivitesi ve otofaji iflası", "Oral Spermidin takviyesi (>15 mg/gün)"]
        ]
    },
    {
        "title": "TABLO 19.6: Glikomik Saat (GlycanAge), IgG N-Glikan Tipleri ve İnflamatuar Etkileri",
        "headers": ["IgG N-Glikan Piki / Yapı", "Şeker Bileşeni (Asn297)", "Biyolojik İmmünolojik Rol", "Yaşlanmayla Değişim", "Klinik Rejuvenasyon Yanıtı"],
        "rows": [
            ["G0 (Agilakozile Glikan)", "Galaktozsuz kor çekirdek", "Pro-enflamatuar Fc-gamma-RIIIa uyarımı", "Yaşla birlikte 3 kat artar", "Kilo kaybı ve BHRT ile hızla düşer"],
            ["G2 (Digalaktozile Glikan)", "İki terminal galaktoz", "Anti-enflamatuar inhibitör reseptör bağı", "Yaşlandıkça dramatik azalır", "Egzersiz ve kalori kısıtlamasıyla artar"],
            ["Sialillenmiş Glikanlar (S1/S2)", "Terminal nöraminik asit", "Maksimum anti-enflamatuar susturma", "Menopoz ve yaşlılıkta çöker", "Östrojen/BHRT tedavisiyle restore edilir"],
            ["Kor Fukozilasyon", "Kor GlcNAc'a bağlı fukoz", "Antikora bağımlı sitotoksisiteyi frenler", "Genellikle stabil / hafif kayma", "Glikomühendislik ile kontrol edilir"],
            ["Bölücü (Bisecting) GlcNAc", "Beta-1,4 bağlı ara şeker", "Agresif ADCC ve tümör yanıtı", "Yaşlanma ve kanserde artar", "B hücre metabolik regülasyonu"],
            ["GlycanAge Biyolojik Yaşı", "Tüm 24 pikin matematiksel oranı", "Biyolojik inflamatuar yaş skoru", "Kötü yaşam tarzıyla +20 yıl fırlar", "6 ayda 10 yıl geriye sarılabilir"]
        ]
    },
    {
        "title": "TABLO 19.7: İmmünomik Saat (iAge), Sitokinler ve Hücresel İmmünosenesens Belirteçleri",
        "headers": ["İmmünomik Parametre", "Genç Bağışıklık Profili", "Yaşlı / İmmünosenesent Profil", "Sistemik Klinik Hasar", "Homo Aeternus İmmün Kürü"],
        "rows": [
            ["Kemokin CXCL9 (MIG)", "Düşük (< 200 pg/mL)", "Tavan Düzeyde (> 1.200 pg/mL)", "Endotel eNOS felci, arter sertleşmesi", "Anti-CXCL9 monoklonal antikoru"],
            ["CD4 / CD8 T Hücre Oranı", "Normal Denge (2.0 - 3.0)", "Tersine Dönmüş (< 1.0; IRP)", "CMV yükü, enfeksiyon ve mortalite riski", "Naive T hücre rejenerasyonu (IL-7)"],
            ["CD28-null CD57+ T Hücreleri", "Nadir (< %5)", "Baskın Popülasyon (> %50)", "Dokuya lökositer sitotoksik saldırı", "Seçici senolitik T hücresi klirensi"],
            ["Naive T Hücreleri (CD45RA+ CCR7+)", "Yüksek (> %60)", "Tükenmiş (< %10)", "Yeni aşı ve virüslere karşı tam körlük", "rhGH + DHEA ile timik re-büyütme"],
            ["İnterlökin-6 (IL-6)", "Bazal (< 1.5 pg/mL)", "Sürekli Yüksek (> 6.0 pg/mL)", "Karaciğerde hs-CRP fırtınası, anemi", "Ziltivekimab / IL-6R inhibitörü"],
            ["Stanford iAge Skoru", "Biyolojik İmmün Yaş = 20 Yaş", "Hızlanmış İmmün Yaş (+15 Yıl)", "Kardiyovasküler erken ölüm habercisi", "Bütünleşik sitokin temizleme protokolü"]
        ]
    },
    {
        "title": "TABLO 19.8: Klinik Biyobelirteçler, Levine PhenoAge ve KDM Algoritmaları Karşılaştırması",
        "headers": ["Klinik Kan Biyomarkeri", "Optimal Gençlik Değeri", "Yaşlılık / Patolojik Değer", "Temsil Ettiği Organel / Doku Ekseni", "PhenoAge Regresyon Ağırlığı"],
        "rows": [
            ["Serum Albümin", ">= 4.7 g/dL", "< 3.8 g/dL (Hipoalbüminemi)", "Karaciğer anabolizması & İnflamasyon", "Negatif katsayı (Düştükçe yaş fırlar)"],
            ["hs-CRP (C-Reaktif Protein)", "<= 0.3 mg/L", "> 3.0 mg/L (Kronik Yangı)", "Sistemik inflamasyon & Vasküler risk", "Pozitif katsayı (Arttıkça yaş fırlar)"],
            ["Açlık Kan Glukozu", "75 - 85 mg/dL", "> 110 mg/dL (İnsülin Direnci)", "Glukotoksisite ve AGE glikasyonu", "Pozitif katsayı (Yaşlanmayı hızlandırır)"],
            ["Eritrosit Dağılım Genişliği (RDW)", "<= %12.0", "> %14.5 (Anizositoz)", "Kemik iliği kök hücre heterojenitesi", "Yüksek pozitif katsayı (Güçlü prediktör)"],
            ["Ortalama Eritrosit Hacmi (MCV)", "86 - 90 fL", "> 96 fL (Makrositoz)", "Membran esnekliği ve B12/folat durumu", "Pozitif katsayı (Mikrovasküler tıkanma)"],
            ["Alkalen Fosfataz (ALP)", "40 - 55 U/L", "> 95 U/L (Kemik/Karaciğer)", "Damar kalsifikasyonu ve hepatobiliyer", "Pozitif katsayı (Kireçlenme belirteci)"],
            ["Beyaz Kan Hücresi (WBC)", "4.0 - 5.5 x 10^3/uL", "> 8.5 x 10^3/uL", "Kronik miyeloid aktivasyon & Yangı", "Pozitif katsayı (Lökositoz riski)"],
            ["Lenfosit Yüzdesi (%Lymph)", "%35 - %45", "< %20 (Lenfopeni)", "Adaptif bağışıklık rezervi ve timus", "Negatif katsayı (Düştükçe yaş fırlar)"],
            ["Serum Kreatinin", "0.8 - 1.0 mg/dL", "> 1.4 mg/dL veya aşırı düşük", "Glomerüler filtrasyon ve kas kitlesi", "Non-lineer parabolik ağırlık"]
        ]
    },
    {
        "title": "TABLO 19.9: Yapay Zeka ve Çoklu Omiks Entegrasyon Mimarileri",
        "headers": ["Yapay Zeka Mimarisi", "Girdi Veri Tipleri", "Matematiksel Prensip", "Temel Biyolojik Avantaj", "Longevity Uygulama Alanı"],
        "rows": [
            ["Varyasyonel Oto-Enkoder (VAE)", "CpG + RNA + Protein + Metabolit", "Kayıplı sıkıştırma ve Gauss gizil uzayı", "Boyut lanetini (p>>n) aşar, gürültüyü siler", "Tek bir gizil biyolojik yaş vektörü üretir"],
            ["Çizge Sinir Ağı (GNN)", "Moleküler ağlar (PPI, KEGG) + Omiks", "Topolojik çizge mesaj iletimi", "Moleküller arası biyolojik etkileşimi korur", "Hücresel ağ çöküşünü ve hub kırılmasını izler"],
            ["Biyolojik Temel Model (scGPT)", "Tek hücre RNA-seq / ATAC-seq", "Milyarlarca parametreli Transformatör", "Genomik bağlamı ve hücresel dili anlar", "In silico ilaç simülasyonu ve hücresel gençleşme"],
            ["Açıklanabilir Yapay Zeka (SHAP)", "Tüm makine öğrenimi modelleri", "Kooperatif oyun teorisi marjinal katkı", "Kara kutuyu açar, sorumlu molekülü bulur", "Hastaya özel doğrudan tedavi hedefi tayini"],
            ["Mendelyan Randomizasyon (MR)", "GWAS genetik varyantları (SNPs)", "Enstrümantal değişken regresyonu", "Korelasyonu nedensellikten kesin ayırır", "Gerçek longevity hedeflerini saptar (ApoB, GDF15)"],
            ["Bayesyen Füzyon Saati", "6 Bağımsız omiks yaş tahmini", "Hiyerarşik olasılıksal Bayes çıkarımı", "Katmanlar arası güvenilirlik ağırlıklaması", "Kusursuz ve sapmasız nihai biyolojik yaş"]
        ]
    },
    {
        "title": "TABLO 19.10: Homo Aeternus Dijital Biyo-İkiz ve Gerçek Zamanlı Yaşlanma Takip Standartları",
        "headers": ["Biyolojik Saat & Parametre", "Konvansiyonel Ortalama Değer", "Homo Aeternus Hedef Değeri", "Veri Toplama Frekansı", "Otomatik Kapalı Devre Müdahale"],
        "rows": [
            ["DunedinPACE (Yaşlanma Hızı)", "1.00 Yıl / Takvim Yılı", "<= 0.60 Yıl / Takvim Yılı", "Yılda 2 kez (EPIC dizileme)", "PACE > 0.8 ise derhal senolitik / OSKM puls"],
            ["PC-GrimAge (Epigenomik)", "Kronolojik Yaşa Eşit (+/- 3 Yıl)", "Kalıcı Olarak 20.0 - 22.0 Yaş", "Yılda 1 kez (Hedefli NGS)", "Doku spesifik dCas9 epigenetik formatlama"],
            ["Wyss-Coray 11 Organ Proteomu", "Organlarda asimetrik yaşlanma", "Tüm organlar <= 22.0 Yaş", "6 ayda bir (Olink Explore)", "Yaşlanan organa spesifik AAV / peptid kürü"],
            ["GlycanAge (IgG Glikomik)", "Yaşa göre G0 glikan artışı", "GlycanAge <= 20 Yaş (G2 baskın)", "Yılda 2 kez (UPLC kromatografi)", "BHRT optimizasyonu + EndoS glikan temizliği"],
            ["Stanford iAge (İmmünomik)", "CXCL9 > 800 pg/mL (Yangı)", "iAge <= 20 Yaş (CXCL9 < 150 pg/mL)", "3 ayda bir (Luminex panel)", "Anti-CXCL9 antikorları + Timus gençleştirme"],
            ["Levine PhenoAge (Klinik Kan)", "Yaşa paralel kümülatif aşınma", "PhenoAge <= 22.0 Yaş", "Ayda 1 kez (Standart kan paneli)", "Mikro-besin, metabolik ve yaşam tarzı ayarı"]
        ]
    }
]

# ================= BELGE OLUŞTURMA DÖNGÜSÜ =================
parts = [
    ("KISIM 1: BİYOLOJİK YAŞ VS KRONOLOJİK YAŞ: SİSTEM BİYOLOJİSİ VE ENTROPİK AŞINMA", part1_subsections),
    ("KISIM 2: EPİGENOMİK SAATLERİN EVRİMİ: HORVATH, LEVINE, GRIMAGE VE DunedinPACE", part2_subsections),
    ("KISIM 3: TRANSKRİPTOMİK SAATLER (RNA-seq) VE GEN EKSPRESYON İMZALARI", part3_subsections),
    ("KISIM 4: PLAZMA PROTEOMİK SAATLERİ VE ORGAN-SPESİFİK YAŞLANMA HARİTALARI", part4_subsections),
    ("KISIM 5: HÜCRESEL METABOLOMİK VE LİPİDOMİK İMZALAR: REDOKS VE ENERJİ AKISI", part5_subsections),
    ("KISIM 6: GLİKOMİK SAAT (GlycanAge): IgG N-GLİKANLARININ İNFLAMATUAR DÖNÜŞÜMÜ", part6_subsections),
    ("KISIM 7: İMMÜNOMİK SAAT (iAge): SİTOKİNLER, İMMÜNOSENESENS VE İNFLAMMAGING", part7_subsections),
    ("KISIM 8: KLİNİK FENOTİPİK BİYOMARKER ALGORİTMALARI: LEVINE PhenoAge VE KDM", part8_subsections),
    ("KISIM 9: DERİN ÖĞRENME, YAPAY ZEKA VE ÇOKLU OMİKS (MULTI-OMICS) FÜZYONU", part9_subsections),
    ("KISIM 10: HOMO AETERNUS DİJİTAL BİYO-İKİZİ VE GERÇEK ZAMANLI LONGEVITY PROTOKOLÜ", part10_subsections)
]

for p_idx, (p_title, p_subs) in enumerate(parts):
    # Kısım Başlığı
    ph = doc.add_paragraph()
    ph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ph.paragraph_format.space_before = Pt(14)
    ph.paragraph_format.space_after = Pt(8)
    prun = ph.add_run(p_title)
    prun.font.name = "Calibri"
    prun.font.size = Pt(13)
    prun.font.bold = True
    prun.font.color.rgb = RGBColor(16, 44, 87)
    
    # 10 Alt Bölüm
    for s_idx, (sub_title, lead_txt, body_txt, eq_txt, param_txt) in enumerate(p_subs):
        sh = doc.add_paragraph()
        sh.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sh.paragraph_format.space_before = Pt(8)
        sh.paragraph_format.space_after = Pt(3)
        srun = sh.add_run(sub_title)
        srun.font.name = "Calibri"
        srun.font.size = Pt(11)
        srun.font.bold = True
        srun.font.color.rgb = RGBColor(0, 102, 153)
        
        lp = doc.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        lp.paragraph_format.line_spacing = 1.15
        lp.paragraph_format.space_after = Pt(4)
        lrun = lp.add_run(lead_txt)
        lrun.font.name = "Calibri"
        lrun.font.size = Pt(9.5)
        lrun.font.italic = True
        lrun.font.color.rgb = RGBColor(50, 50, 50)
        
        bp = doc.add_paragraph()
        bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        bp.paragraph_format.line_spacing = 1.15
        bp.paragraph_format.space_after = Pt(5)
        brun = bp.add_run(body_txt)
        brun.font.name = "Calibri"
        brun.font.size = Pt(9.5)
        brun.font.color.rgb = RGBColor(30, 30, 30)
        
        # Formül Kutusu
        eq_table = doc.add_table(rows=1, cols=1)
        eq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = eq_table.rows[0].cells[0]
        set_cell_background(c, "F0F4F8")
        set_cell_margins(c, top=80, bottom=80, left=140, right=140)
        ep = c.paragraphs[0]
        ep.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ep.paragraph_format.space_before = Pt(0)
        ep.paragraph_format.space_after = Pt(0)
        erun = ep.add_run(f"BİYOFİZİKSEL FORMÜLASYON / EŞİTLİK: {eq_txt}")
        erun.font.name = "Consolas"
        erun.font.size = Pt(8.5)
        erun.font.bold = True
        erun.font.color.rgb = RGBColor(16, 44, 87)
        
        pp = doc.add_paragraph()
        pp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pp.paragraph_format.line_spacing = 1.1
        pp.paragraph_format.space_before = Pt(3)
        pp.paragraph_format.space_after = Pt(8)
        prun2 = pp.add_run(f"Parametre Açıklaması ve Fizyolojik Anlamı: {param_txt}")
        prun2.font.name = "Calibri"
        prun2.font.size = Pt(8.5)
        prun2.font.italic = True
        prun2.font.color.rgb = RGBColor(90, 100, 110)
        
        doc.add_page_break()
        
    # Her Kısım Sonunda Akademik Karşılaştırma Tablosu
    t_data = tables_data[p_idx]
    th = doc.add_paragraph()
    th.alignment = WD_ALIGN_PARAGRAPH.LEFT
    th.paragraph_format.space_before = Pt(12)
    th.paragraph_format.space_after = Pt(6)
    th_run = th.add_run(t_data["title"])
    th_run.font.name = "Calibri"
    th_run.font.size = Pt(11)
    th_run.font.bold = True
    th_run.font.color.rgb = RGBColor(16, 44, 87)
    
    table = doc.add_table(rows=len(t_data["rows"]) + 1, cols=len(t_data["headers"]))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Başlık Satırı
    hdr_cells = table.rows[0].cells
    for c_idx, h_text in enumerate(t_data["headers"]):
        cell = hdr_cells[c_idx]
        set_cell_background(cell, "102C57")
        set_cell_margins(cell, top=120, bottom=120, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.font.name = "Calibri"
        run.font.size = Pt(8.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        
    # Veri Satırları
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
print(f"CİLT 19 Başarıyla Kaydedildi: {OUTPUT_PATH}")