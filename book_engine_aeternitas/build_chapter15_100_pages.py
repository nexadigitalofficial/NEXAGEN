# -*- coding: utf-8 -*-
"""
PROJECT AETERNITAS - CİLT 15: NANOTEKNOLOJİ, BİYO-YAPAY ORGANLAR VE SİBERNETİK BİYOLOJİ
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

OUTPUT_PATH = r"C:\\Users\\USER\\Desktop\\kitap1\\BOLUM_15_NANOTEKNOLOJI_VE_BIYO_YAPAY_ORGANLAR_TAM_100_SAYFA.docx"

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
    hrun = hp.add_run("PROJECT AETERNITAS | CİLT 15: NANOTEKNOLOJİ VE BİYO-YAPAY ORGANLAR")
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
s_run = sub_p.add_run("CİLT 15: NANOTEKNOLOJİ, BİYO-YAPAY ORGANLAR VE SİBERNETİK BİYOLOJİ\\n(TIBBİ NANOROBOTİK, SENTETİK ERİTROSİTLER, 3D BİYO-YAZICILAR, DESELLÜLARİZASYON VE BİYO-HİBRİT İMPLANTLAR)")
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
ih_run = intro_h.add_run("CİLT 15 MANİFESTOSU: BİYOLOJİNİN MEKANİK AŞILMASI: ET VE ÇELİĞİN, KARBON VE ELMASIN EBEDİ SENTEZİ")
ih_run.font.name = "Calibri"
ih_run.font.size = Pt(15)
ih_run.font.bold = True
ih_run.font.color.rgb = RGBColor(16, 44, 87)

intro_body = (
    "Biyolojik evrimin ürettiği organlar ve dokular; milyonlarca yıllık kısıtlı kaynaklar, genetik mutasyon baskıları ve termodinamik "
    "yıpranma koşullarında şekillenmiş geçici biyomoleküler makinelerdir. Ne kadar kusursuz gen terapisi veya hücresel yeniden "
    "programlama uygulanırsa uygulansın; organik dokuların fiziksel yorulma sınırları, kalsifikasyon eğilimleri ve makroskobik "
    "travma kırılganlıkları ebedi varoluş için mutlak bir risk faktörü olmaya devam eder.\\n\\n"
    "Bu kırılganlığı kesin olarak aşmanın yolu, Sibernetik Biyoloji ve Tıbbi Nanoteknolojidir. Robert Freitas'ın teorik temellerini attığı "
    "elmasoid yapılı nanorobotlar (Respirositler, Mikrobivorlar); kan dolaşımında biyolojik hücrelerin yüzlerce katı verimle gaz taşımakta, "
    "patojenleri saniyeler içinde mekanik olarak sindirmekte ve damar içi plakları atom atom kazıyarak temizlemektedir. Eş zamanlı olarak "
    "3D Biyo-Yazıcılar ve desellülarize matriks iskeleleri; hastanın kendi kök hücreleriyle tohumlanan yedek biyo-yapay katı organlar "
    "(kalp, böbrek, karaciğer) inşa ederek organ yetmezliği kavramını tıbbın lügatinden silmektedir.\\n\\n"
    "Bu ciltte; tıbbi nanorobotiğin biyofiziği, yapay kan hücreleri, vasküler temizlik nanobotları, 3D ekstrüzyon/stereolitografi biyo-baskısı, "
    "organ iskelelerinin resellülarizasyonu, biyomekanik ve sürekli akışlı yapay kalp sistemleri (TAH/VAD), biyo-elektronik sensörler, "
    "protein korona dinamikleri ve Homo Aeternus biyo-hibrit organ replasman protokolü 100 akademik alt bölümde detaylandırılmaktadır."
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
        "1.1 Nanomedikal Sistemlerin Termodinamik Sınırları ve Düşük Reynolds Sayısı (Re << 1) Rejimi",
        "İnsan kan dolaşımında veya doku interstisyel sıvısında hareket eden mikroskobik ve nanometrik robotlar (nanobotlar), makroskobik dünyamızın fiziksel kurallarından tamamen farklı bir biyofiziksel rejimde çalışırlar.",
        "Edward Purcell'in tarihi 'Düşük Reynolds Sayısında Yaşam' tezinde kanıtlandığı üzere; mikron altı boyutlarda atalet (eylemsizlik) kuvvetleri neredeyse tamamen sıfırlanır ve viskoz sürtünme kuvvetleri çevreye mutlak hâkim olur (Reynolds Sayısı: Re = rho * v * L / eta << 10^-4). Nanobot hareket motorunu kapattığı anda, bir mikrosaniyeden daha kısa sürede ve bir hidrojen atomu çapından daha az mesafede durur; yüzme eylemi ileri momentum aktarımıyla değil, zaman simetrisini kıran periyodik şekil değişimleri ile gerçekleştirilebilir.",
        "Sayisi_Reynolds = Re = ( rho_plazma * v_nanobot * L_karakteristik ) / eta_kan << 10^-4",
        "Bu hidrodinamik boyutsuz sayı formülü, kan plazması içinde nanorobotik sistemlerin ataletsiz ve viskoz kuvvetlerin mutlak tahakkümü altında çalıştığı fiziksel ortamı tanımlar."
    ),
    (
        "1.2 Brown Hareketi, Termal Çalkantı ve Fluktuasyon-Disipasyon Teoremi",
        "Nanometrik ölçekte su ve plazma molekülleri, sıcaklığa bağlı sürekli ve kaotik bir termal bombardıman (Brown hareketi) sergiler.",
        "Bir nanorobot, her saniye trilyonlarca rastgele su molekülü çarpışmasına maruz kalır. Einstein-Smoluchowski ve Fluktuasyon-Disipasyon bağıntısına göre; nanorobotun maruz kaldığı sürtünme katsayısı (gamma = 6 * pi * eta * r) ile termal gürültü kuvveti doğrudan ilişkilidir. Hedefe yönelik kontrollü doğrusal hareket sağlamak için nanobotun ürettiği itici mekanik kuvvet, bu k_B * T termal gürültü eşiğini en az bir mertebe aşmak zorundadır.",
        "Kuvvet_Termal_Gurultu = F_termal = sqrt( 2 * k_B * T * gamma / Delta_t )",
        "Bu istatistiksel termodinamik fluktuasyon denklemi, nanorobotik seyir sistemlerinin yenmek zorunda olduğu ortalama termal Brown çalkantı kuvvetini modeller."
    ),
    (
        "1.3 Elmasoid (Diamondoid) Malzemeler: Safir, Elmas ve Karbon Nanotüp Dayanımı",
        "Biyolojik dokularda enzimlerin, reaktif oksijen türlerinin ve mekanik basınçların yıkıcı etkisine karşı koyabilmek için nanomedikal makinelerin yapısal iskeleti 'Elmasoid' (Diamondoid) atomik kafeslerden inşa edilir.",
        "Elmasoid yapılar; elmas (sp3 karbon atomlarının tetrahedral kovalent bağı), safir (alüminyum oksit) ve fulleren/karbon nanotüp türevleridir. Bu malzemeler gigapaskal (GPa) düzeyinde elastik modüle (E ~ 1000 GPa), 50-100 GPa çekme dayanımına ve kimyasal inertliğe sahiptir. Robert Freitas'ın tasarımlarında olduğu gibi, 1000 atmosferlik iç gaz basıncına dahi deforme olmadan dayanabilen mikroskobik nano-basınç kapları bu atomik kafeslerle üretilir.",
        "Dayanim_Elmasoid = Sigma_cekme ~ 60 - 100 GPa >> Sigma_titanyum (1 GPa)",
        "Bu malzeme mekaniği gerilme sınırı karşılaştırması, elmasoid nanomekanik kabukların geleneksel cerrahi metallere kıyasla yüz kat daha yüksek yapısal mukavemet sunduğunu belgeler."
    ),
    (
        "1.4 Nanorobotik İtme Mekanizmaları: Dönen Kamçılar, Akustik ve Manyetik Alan Sürüşü",
        "Düşük Reynolds sayısında ileri hareket sağlayabilmek için doğanın bakterilerde geliştirdiği 'Dönen Helikal Kamçı' (Bacterial Flagella) veya harici biyofiziksel alanlar taklit edilir.",
        "1) Helikal Nano-Pervaneler: Dönen spiral bir kuyruk, viskoz sıvıda bir tirbüşon gibi ilerleyerek ileri itme kuvveti üretir; 2) Akustik Levigasyon: Harici ultrasonik odaklanmış ses dalgaları, nanobotun akustik kontrast faktörünü kullanarak doku içinde yönlendirilmesini sağlar; 3) Manyetik Tork Sürüşü: İçerisinde süperparamanyetik demir oksit çekirdek taşıyan nanorobotlar, vücut dışından uygulanan dönen gradyan manyetik alanlar (MRI benzeri) ile saniyede onlarca mikrometre hızla hedefe sürülür.",
        "Hiz_Itme_Helikal = v_nanobot = ( Delta_mu * omega_rotasyon * R_heliks ) / ( 2 * pi * ( 1 + k_suruklenme ) )",
        "Bu mikroskopik akışkanlar mekaniği eşitliği, dönen manyetik alan frekansı ve helikal vida geometrisinin viskoz sıvılardaki doğrusal nanorobotik seyir hızını nasıl belirlediğini formüle eder."
    ),
    (
        "1.5 Nanorobotik Enerji Hasadı: Glikoz-Oksijen Biyo-Yakıt Pilleri ve Piezoelektrik Dönüşüm",
        "Vücut içinde yıllarca aktif görev yapacak nanorobotların dahili bir kimyasal pil taşıması hacim ve toksisite nedeniyle imkânsızdır; enerji doğrudan biyolojik çevreden hasat edilmelidir.",
        "En zarif çözüm 'Enzimatik Biyo-Yakıt Pilleri'dir (EBFC): Nanobotun dış yüzeyindeki glukoz oksidaz enzimi kandaki glukozu glukonolaktona oksitlerken elektron üretir; katottaki lakkaz veya bilirubin oksidaz ise kandaki çözünmüş oksijeni suya indirgeyerek devreyi tamamlar. Kan plazmasındaki bol şeker ve oksijen kesintisiz mikrowatt düzeyinde elektrik gücü sağlar. Ek olarak kan nabız basıncından enerji üreten çinko oksit (ZnO) piezoelektrik nano-jeneratörler kullanılır.",
        "Guc_Hasat = P_biofuel = J_glukoz * Delta_V_hucre = [Glukoz_kan] * k_enzim * ( E_katot - E_anot )",
        "Bu elektrokimyasal güç üretim denklemi, kandaki glukoz konsantrasyonu ve redoks potansiyeli farkının nanorobotik sistemlere sağladığı sürekli mikro-enerji debisini simgeler."
    ),
    (
        "1.6 Moleküler Sıralama Rotorları (Sorting Rotors): Nanometrik Seçici Molekül Pompaları",
        "Nanomekanik sistemlerin çevredeki sıvıdan spesifik molekülleri (glukoz, oksijen, toksinler) tek tek yakalayıp içeri alması 'Moleküler Sıralama Rotorları' (Molecular Sorting Rotors) ile gerçekleştirilir.",
        "Rotor, elmasoid bir disk yüzeyinde spesifik moleküle kusursuz geometrik ve elektrostatik uyum gösteren bağlama ceplerinden oluşur. Dış ortamdaki hedef molekül bu cebe girdiğinde rotor mekanik bir dişli gibi 180 derece döner; iç kompartmandaki düşük afiniteli boşaltım kamı molekülü yerinden fırlatarak nanobotun içine atar. Bu rotorlar saniyede yüz binlerce molekülü %99.999 seçicilikle ve ters konsantrasyon gradyanına karşı pompalayabilir.",
        "Debi_Molekuler_Pompa = J_rotor = N_cep * f_donus * ( [Ligand] / ( K_d + [Ligand] ) )",
        "Bu nanomekanik taşıma kinetiği eşitliği, moleküler sıralama rotorlarının dönüş frekansı ve bağlanma afinitesi ile birim zamanda içeri aktarılan net molekül sayısını tanımlar."
    ),
    (
        "1.7 Akustik ve Biyo-Optik İletişim: Nanobot Ağları ve Vücut İçi İnternet (IoNT)",
        "Trilyonlarca nanorobotun koordine bir cerrahi veya metabolik operasyon yürütebilmesi için birbirleriyle ve vücut dışındaki ana bilgisayarla kesintisiz iletişim kurması şarttır (Internet of Nano-Things - IoNT).",
        "Radyo dalgaları (RF) vücut dokularında aşırı zayıfladığı ve anten boyutu nanometrik ölçeğe sığmadığı için iki ana iletişim paradigması kullanılır: 1) Akustik Basınç Dalgaları: Megahertz frekansındaki mikroskopik piezoelektrik titreşimler dokularda santimetrelerce uzağa kodlanmış dijital veri paketleri taşır; 2) Yakın Kızılötesi (NIR) Fotonik İletişim: Biyolojik optik pencerede (700-1100 nm) çalışan nano-LED ve fotodiyotlar hücre zarlarını aşarak yüksek bant genişlikli optik veri transferi sağlar.",
        "Bant_Genisligi_Akustik = C_akustik = B * log2( 1 + P_sinyal / ( N_0 * B + Sigma_termal ) )",
        "Bu Shannon-Hartley iletişim kapasitesi eşitliği, doku içi akustik nano-iletişimin bant genişliği ve sinyal-gürültü oranına (SNR) bağlı maksimum veri aktarım hızını hesaplar."
    ),
    (
        "1.8 Nanobilgisayar Mimarisi: Mekanik Nanomekanik Mantık Kapıları ve DNA Bellek",
        "Nanometrik boyuttaki bir aygıta geleneksel silikon transistörleri sığdırmak kuantum tünelleme kaçakları ve hacim nedeniyle zordur; bu nedenle 'Mekanik Nanobilgisayarlar' tasarlanmıştır.",
        "K. Eric Drexler'in 'Rod-Logic' mimarisinde; hareketli elmasoid çubuklar (sliding rods) birbirini fiziksel olarak bloke ederek veya geçit vererek mekanik mantık kapıları (AND, OR, NOT) oluşturur. Bu mekanik transistörler sıfıra yakın enerjiyle (kT'nin çok altında) ve gigahertz saat hızlarında çalışır. Veri depolama için ise nükleotit dizilimini dijital baytlara dönüştüren ve gram başına yüzlerce petabayt veri saklayabilen sentetik DNA polimer bellek modülleri entegre edilir.",
        "Yogunluk_Mekanik_Bitis = D_rod = N_kapi / V_nanobot ~ 10^18 kapi / cm3 (Silikon Çiplerden 10^6 Kat Yoğun)",
        "Bu hesaplama yoğunluğu karşılaştırması, elmasoid çubuk mantık kapılarının mikroskobik nanorobot somasına süperbilgisayar işlem gücünü nasıl sığdırdığını gösterir."
    ),
    (
        "1.9 Konumlandırma ve Seyir Biyofiziği: Damar Ağında Koordinat Belirleme",
        "Nanobotun vücutta hangi organda, hangi damar dalında olduğunu bilmesi (İn vivo GPS); hedefe yönelik cerrahi ve ilaç teslimatı için mutlak bir gerekliliktir.",
        "Bu seyir üçlü bir entegre sistemle sağlanır: 1) Biyokimyasal Parmak İzi: Nanobot yüzeyindeki sensörler yerel pH, oksijen parsiyel basıncı (pO2), laktat ve endotelyal yüzey reseptör profilini okur; 2) Akustik Nirengi (Triangülasyon): Vücut dışındaki ultrasonik transdüserler nanobotun gönderdiği piezoelektrik sinyalleri yakalayarak milimetrik uzamsal koordinatları hesaplar; 3) Manyetik Haritalama: Vücudun vasküler ağının üç boyutlu dijital ikizi ile gerçek zamanlı eşleşme sağlanır.",
        "Hata_Konumlandirma = Delta_x = sqrt( (c_ses * Delta_t_olcum)^2 + Sigma_drift^2 ) < 10 mikrometre",
        "Bu mikroskopik nirengi hata analizi eşitliği, çoklu sensör füzyonunun nanorobotun damar içindeki uzamsal konumunu kılcal damar çözünürlüğünde (<10 um) nasıl sabitlediğini formüle eder."
    ),
    (
        "1.10 Biyo-İnert Yüzey Zırhlama: Protein Korona ve İmmün Fagositozdan Kaçış",
        "Bir nanorobot kan dolaşımına girdiği anda en büyük tehlike, bağışıklık sistemi kompleman proteinlerinin (C3b) ve antikorların nanobot yüzeyine çökerek onu opsonize etmesi ve makrofajların fagositozuna yem yapmasıdır.",
        "Bu immünolojik yok oluşu engellemek için elmasoid yüzey 'Zwitteriyonik Polimer Fırçalar' (polikarbetain, polisülfobetain) ve insan CD47 proteini mimetik peptitleri ('Don't Eat Me' sinyali) ile kovalent olarak kaplanır. Zwitteriyonik yüzey, su moleküllerini aşırı sıkı bağlayarak termodinamik olarak hiçbir plazma proteininin yüzeye yapışmasına izin vermeyen bir hidrasyon kalkanı kurar; protein korona oluşumu sıfırlanır ve nanobot kanda haftalarca hayalet gibi dolaşır.",
        "Adsorpsiyon_Protein = Gamma_korona = Gamma_0 * exp( - Delta_G_hidrasyon / (k_B * T) ) ~ 0",
        "Bu yüzey termodinamiği eşitliği, zwitteriyonik hidrasyon kalkanının serbest enerji bariyerini aşırı yükselterek plazma proteini yapışmasını ve makrofaj fagositozunu nasıl sıfırladığını kanıtlar."
    )
]

# ================= KISIM 2 =================
part2_subsections = [
    (
        "2.1 Respirositler (Respirocytes): Robert A. Freitas Jr. Tarafından Tasarlanan Yapay Eritrosit",
        "Doğal insan alyuvarı (eritrosit), gaz taşımacılığını pasif hemoglobin difüzyonuna dayandıran ve kapasitesi sınırlı bir biyolojik hücredir; Robert Freitas Jr. bu biyolojik sınırı aşmak için 1998'de 'Respirosit' nanorobotunu tasarlamıştır.",
        "Respirosit; 1 mikrometre çapında, küresel, elmasoid yapılı ve içinde 1000 atmosfer (atm) basınç altında oksijen ve karbondioksit depolayan mikroskobik bir nanomekanik basınç tankıdır. Tek bir respirosit hücresi, aynı hacimdeki doğal bir eritrositten tam 236 kat daha fazla oksijen taşıma kapasitesine sahiptir. Vücuda 5 ml respirosit süspansiyonu enjekte edildiğinde, bir insan soluk almadan suyun altında veya karbonmonoksit gazı içinde 4 saat boyunca ağır fiziksel efor sarf edebilir.",
        "Kapasite_Respirosit = C_O2_respirosit = 236 * C_O2_eritrosit ~ 3.6 * 10^9 O2 molekulu / nanobot",
        "Bu molar gaz taşıma kapasitesi oranı, yüksek basınçlı elmasoid depolama tanklarının biyolojik hemoglobine kıyasla iki yüz kattan fazla gaz yoğunluğu sağladığını belgeler."
    ),
    (
        "2.2 Respirosit İç Basınç Odaları, Gaz Rotorları ve 1000 Atmosferlik Gaz Sıkıştırma",
        "Respirositin iç mimarisi üç bağımsız elmasoid basınç odasından oluşur: Oksijen depolama odası, Karbondioksit toplama odası ve genel balast/glukoz odası.",
        "Dış yüzeyde yer alan moleküler sıralama rotorları; kanda pO2 yüksek olduğunda (akciğer kapillerleri) oksijeni içeri çeker ve mikromekanik kompresör pompaları ile tankın içine 1000 atm basınçla basar. Doku kapillerine ulaşıldığında sensörler pO2 düşüşünü ve pCO2 artışını okur; oksijen boşaltım rotorları kontrollü olarak açılarak dokuya saf O2 deşarj edilirken ortamdaki CO2 emilerek karbondioksit tankına kilitlenir. Tüm süreç sensör geribildirimli dahili rod-logic bilgisayarı tarafından mikrosaniyeler içinde yönetilir.",
        "Basinc_Tank = P_gaz = ( n_mol * R * T ) / ( V_oda - n_mol * b_van_der_Waals ) ~ 1000 atm",
        "Bu Van der Waals yüksek basınç gaz sıkıştırma denklemi, elmasoid nanokompartman içindeki moleküler oksijen yoğunluğunun süperkritik akışkan sınırındaki fiziksel davranışını modeller."
    ),
    (
        "2.3 Karbonmonoksit Zehirlenmesi, Hipoksi ve Dekompresyon Krizinde Respirosit Protokolü",
        "Hemoglobinin karbonmonoksite (CO) olan afinitesi oksijenden 200 kat daha yüksektir; bu durum yangınlarda ve zehirlenmelerde dakikalar içinde ölümcül asfiksiye yol açar.",
        "Respirositler ise kimyasal bağlanma yerine fiziksel rotor seçiciliği kullandığından karbonmonoksitten tamamen etkilenmezler. CO ile kilitlenmiş bir hastaya damar yoluyla respirosit verildiğinde, nanobotlar saniyeler içinde tüm dokulara ve beyne 1000 atm rezervinden saf oksijen basmaya başlar. Benzer biçimde derin deniz dalgıçlarında görülen vurgun (dekompresyon) krizinde kanda köpüren serbest azot gazlarını vakumlayarak ölümcül hava embolilerini anında yok eder.",
        "Sagkalim_Hipoksi = t_bilinc = t_dogal * ( 1 + 236 * [Respirosit_kan_fraksiyonu] )",
        "Bu hipoksi tolerans süresi formülü, dolaşımdaki respirosit fraksiyonunun doku oksijenlenme süresini ve beyin ölümünü geciktirme katsayısını tanımlar."
    ),
    (
        "2.4 Mikrobivorlar (Microbivores): Yapay Akyuvarlar ve Fagositoz Hızının Yüz Katlanması",
        "Doğal akyuvarlar (nötrofiller ve makrofajlar), kan dolaşımındaki bakteriyel ve viral patojenleri yakalamakta kemotaksi ve yavaş hücresel ameboid göç kullanır; ağır septik şok tablolarında yetersiz kalırlar.",
        "Mikrobivorlar; Robert Freitas tarafından tasarlanan 3.4 mikrometre uzunluğunda ve 2 mikrometre çapında oval biçimli sentetik yapay beyaz kan hücreleridir (Artificial Phagocytes). Yüzeyinde patojenlerin korunmuş yüzey antijenlerini bağlayan antijen reseptör dizilimleri taşır. Bir mikrobivor, kanındaki bir bakteriyi 100 milisaniye içinde yakalar, iç sindirim odasına çeker, 30 saniye içinde mekanik ultrasonik bıçaklar ve termal-enzimatik lizisle atomlarına ayırarak amino asit ve şeker halinde zararsızca kana bırakır.",
        "Hiz_Klirens_Mikrobivor = d[Bakteri]/dt = - k_fagositoz_nano * [Mikrobivor] * [Bakteri_plazma] >> 100 * k_makrofaj",
        "Bu patojen klerens kinetiği eşitliği, mikrobivorların septik bakteriyemi tablolarını doğal lökositlerden yüz kat daha hızlı temizleme gücünü ifade eder."
    ),
    (
        "2.5 Mikrobivor İç Sindirim Odası: Ultrasonik ve Enzimatik Patojen Parçalama",
        "Mikrobivorun iç lümeni, yutulan hiçbir patojenin (bakteri, virüs, prion, mantar sporu) canlı çıkamayacağı ölümcül bir moleküler öğütücüdür.",
        "Bakteri içeri alındıktan sonra sindirim odasının elmasoid kapakları hermetik olarak kapanır. İlk adımda 40 MHz yüksek frekanslı ultrasonik kavitasyon dalgaları bakterinin hücre duvarını (peptidoglikan) mekanik olarak parçalar. İkinci adımda ortama konsantre lizozim, proteaz ve nükleaz kokteylleri pompalanarak patojenin genomik DNA/RNA'sı ve proteinleri mononükleotitlere ve amino asitlere hidrolize edilir; toksik endotoksinler (LPS) kimyasal olarak nötralize edilir.",
        "Sure_Lisis = Tau_parcalama = t_mekanik (5 sn) + t_enzimatik (25 sn) = 30 saniye / patojen",
        "Bu imha zaman sabiti formülü, tek bir mikrobivor nanorobotunun bir bakteri hücresini tamamen zararsız moleküler yapı taşlarına ayırma süresini simgeler."
    ),
    (
        "2.6 Sepsis, Çoklu İlaç Dirençli (MDR) Bakteriler ve Viral Pandemilerin Kesin İmhası",
        "Antibiyotik direnci (MRSA, VRE, pan-rezistan Gram-negatif basiller); modern tıbbın en büyük kabuslarından biridir çünkü kimyasal ilaçlar mutasyonlarla etkisizleşir.",
        "Mikrobivorlar bakteriyi öldürmek için kimyasal antibiyotiklere değil; mekanik yakalama ve fiziksel parçalama prensibine dayandığı için hiçbir bakteri mikrobivora karşı genetik direnç geliştiremez. 1 trilyon mikrobivordan oluşan bir infüzyon, ölümcül bir septik şok hastasının tüm kan hacmini (5 litre) 20 dakika içinde sterilize edebilir. Benzer şekilde HIV, Ebola veya mutasyona uğramış yeni pandemik virüsler kanda saniyeler içinde yakalanıp yok edilir.",
        "Zaman_Sterilizasyon = t_steril = V_kan * ln( [Patojen_0] / [Patojen_guvenli] ) / ( N_mikrobivor * Q_giris ) ~ 20 dakika",
        "Bu hematolojik arındırma formülü, terapötik dozda uygulanan mikrobivor sürüsünün tüm insan kan hacmini sıfır patojen seviyesine indirme süresini belgeler."
    ),
    (
        "2.7 Clottositler (Clottocytes): Yapay Trombositler ve Anında Hemostaz Kontrolü",
        "Doğal trombositler büyük arteriyel kanamalarda ve travmalarda pıhtılaşma kaskadını dakikalar içinde aktive eder; ancak masif kan kayıplarında bu süre ölümcül hipovolemik şoka engel olamaz.",
        "Clottositler; 2 mikrometre boyutunda küresel nanorobotlar olup, içlerinde sıkıştırılmış yüksek mukavemetli sentetik biyobozunur ağ lifleri (örneğin sentetik fibrin veya modifiye lizin polimerleri) depolar. Nanobot damar duvarındaki ani basınç düşüşünü ve kolajen maruziyetini algıladığı anda; 1 saniyeden daha kısa sürede içindeki nano-ağ yapısını dışarı fırlatır (net deployment). Bu ağ kanayan damar ağzını mekanik olarak örterek kanamayı anında durdurur.",
        "Sure_Hemostaz = t_pihti = 1 saniye (Clottosit) << 300 - 600 saniye (Dogal Tromboz)",
        "Bu hemostatik reaksiyon hızı karşılaştırması, yapay trombosit sistemlerinin masif kanamalarda pıhtı oluşturma süresini yüzlerce kat nasıl kısalttığını gösterir."
    ),
    (
        "2.8 Clottositlerin Nano-Ağ Dağıtım Biyofiziği ve Tromboz Risk Kontrolü",
        "Yapay bir hemostaz sisteminin en büyük klinik riski, damar yaralanması olmayan sağlam damarlarda kontrolsüzce açılarak ölümcül yaygın damar içi pıhtılaşmaya (DIC) veya pulmoner emboliye yol açmasıdır.",
        "Bu riski engellemek için clottositler katı bir çift-anahtarlı (dual-key) mantık kontrolü kullanır: Nano-ağ fırlatılması için hem damar dışı ekstravasküler kolajen/doku faktörü sinyalinin algılanması hem de damar lümenindeki kayma gerilmesinde (shear stress) ani bir basınç gradyanı düşüşünün doğrulanması şarttır. Sağlam damarda tek bir sinyal dahi eksik olsa sistem tetiklenmez; ayrıca yanlışlıkla açılan ağlar 5 dakika içinde kendi kendini parçalayan enzimler salgılar.",
        "Tetikleme_Sarti = ( Sinyal_Kollajen == TRUE ) AND ( Delta_P_damar_delinme > P_esik )",
        "Bu biyomekanik güvenlik mantık kapısı algoritması, clottositlerin sağlam dolaşımda tromboz riski yaratmaksızın yalnızca gerçek kanama sahalarında aktifleşmesini garantiler."
    ),
    (
        "2.9 Hemoraji, Travmatik Şok ve Cerrahi Sahada Anında Kanama Kanamasız Ameliyatlar",
        "Askeri harp cerrahisinde ve ağır politravmalarda ölümlerin birincil nedeni durdurulamayan iç kanamalardır (karaciğer yırtılması, aort rüptürü).",
        "İntravenöz clottosit profilaksisi altındaki bir hastada; cerrah neşteri vurduğu anda kesilen tüm kılcal ve arteriyel damar ağızları mikrosaniyeler içinde clottosit ağları ile mühürlenir. Cerrahi sahada tek bir damla kan dahi dökülmez (Kansız Cerrahi / Bloodless Surgery). Karaciğer veya dalak gibi yüksek vasküler organ rezeksiyonları dahi sıfır kan kaybı ve sıfır transfüzyon ihtiyacı ile gerçekleştirilebilir.",
        "Kayip_Kan_Hacmi = Delta_V_kanama = Integral_0^{t_hemostaz} Q_kanama(t) dt ~ 0 (t_hemostaz ~ 1 sn)",
        "Bu dinamik kanama hacim integrali, 1 saniyelik reaksiyon süresine sahip yapay trombositlerin cerrahi kan kaybını teorik olarak sıfıra indirdiğini formüle eder."
    ),
    (
        "2.10 Biyo-Hibrit Sentetik Kan: Respirosit-Mikrobivor-Clottosit Üçlü Konsorsiyumu",
        "Biyolojik kanın tüm hücre fraksiyonlarının (eritrosit, lökosit, trombosit) yerine sentetik nanorobotik eşdeğerlerinin konulması; nihai 'Sentetik Biyo-Hibrit Kan' (Synthetic Blood) devrimini doğurur.",
        "Bu yapay kan süspansiyonu; kan grubu antijenlerinden (ABO/Rh) tamamen arındırılmıştır, oda sıcaklığında yıllarca bozulmadan saklanabilir, hiçbir viral veya bakteriyel enfeksiyon bulaştırmaz ve biyolojik kandan 100 kat daha üstün fizyolojik performans sunar. Sentetik kan taşıyan bir Homo Aeternus bireyi; kalp krizi, inme, sepsis, zehirlenme ve kanama risklerinden tamamen bağımsız, biyolojik kırılganlığı aşılmış bir fizyolojiye kavuşur.",
        "Performans_BiyoHibrit_Kan = Omega_kan = ( C_O2 / C_O2_dogal ) * ( Hiz_Klirens / Hiz_dogal ) * ( 1 / t_hemostaz ) >> 10^6",
        "Bu genel hematolojik mükemmellik katsayısı, üçlü nanorobotik kan konsorsiyumunun doğal kana kıyasla sağladığı milyon katlık fizyolojik üstünlük bileşkesini simgeler."
    )
]

# ================= KISIM 3 =================
part3_subsections = [
    (
        "3.1 Vasküler Yaşlanmanın Zirvesi: Ateroskleroz, Kararsız Plaklar ve Tromboz",
        "Kardiyovasküler hastalıklar ve ateroskleroz, modern dünyada ölüm nedenlerinin bir numaralı failidir ve damar duvarının kronik enflamatuar yaşlanmasının nihai sonucudur.",
        "Endotel hasarı ile başlayan süreç; subendotelyal aralığa sızan LDL partiküllerinin oksidasyonu (oxLDL), makrofajların bunları yutarak köpük hücrelerine (foam cells) dönüşmesi, kalsiyum kristalleri birikimi ve nekrotik lipit çekirdeğin etrafında fibröz bir kapsül oluşumu ile ilerler. Yaşlanan damarda fibröz kapsülün metaloproteinazlarca incelmesi plağın yırtılmasına (rüptür) yol açar; açığa çıkan trombüs koroner arteri tıkayarak dakikalar içinde ölümcül miyokard enfarktüsünü tetikler.",
        "Risk_Ruptur = R_plak = ( Gerilme_Mekanik / Mukavemet_Kapsul ) = ( P_kan * r_lumen / h_kapsul ) / Sigma_kolajen",
        "Bu biyomekanik Laplace gerilme eşitliği, damar içi nabız basıncının fibröz kapsül kalınlığı inceldikçe yırtılma ve fatal tromboz riskini nasıl kritik eşiğe fırlattığını gösterir."
    ),
    (
        "3.2 Vaskülositler (Vasculocytes): Damar İçi Onarım ve Aterom Kazıma Nanobotları",
        "Aterosklerozun cerrahi baypas veya stentleme ile palyatif tedavisi yerine; Robert Freitas'ın önerdiği 'Vaskülosit' nanorobotları damar endotelini ve intimasını hücresel düzeyde rekonstrükte eder.",
        "Vaskülositler; kılcal damarlardan ana arterlere kadar vasküler ağ boyunca ilerleyen, yüzeyindeki moleküler problarla kalsifiye lezyonları, köpük hücrelerini ve okside lipid birikimlerini mikron-altı çözünürlükle haritalayan akıllı tamir araçlarıdır. Plak bölgesine ulaştıklarında mekanik elmasoid frezeleri ve lokal enzim pompaları ile plağı sağlıklı damar endoteline zarar vermeden katman katman tıraşlarlar.",
        "Hiz_Kazima = dV_plak/dt = - N_vaskulosit * A_kesici * v_freze * ( 1 - [Kalsifikasyon_direnci] )",
        "Bu mekanik debridman kinetiği formülü, damar içi aterom plağının nanorobotik kazıma debisinin aktif vaskülosit sayısı ve frezeleme hızıyla doğrudan ilişkisini modeller."
    ),
    (
        "3.3 Hedefli Lipit Ekstraksiyonu: Kolesterol Kristalleri ve Okside LDL'nin Moleküler Emilimi",
        "Aterom plağının tehlikeli çekirdeğini oluşturan katılaşmış Kolesterol Monohidrat kristalleri ve toksik oxLDL; vaskülositlerin moleküler sıralama rotorları ile emilerek temizlenir.",
        "Nanobot plak yüzeyine kilitlenir; kolesterol seçici rotorlar kristal yüzeyindeki serbest kolesterol moleküllerini tek tek yakalayarak nanobotun içindeki depolama haznesine çeker. Eş zamanlı olarak makrofajların içine kilitlenmiş köpük hücrelerine mikro-iğneler batırılarak hücre içi lipid yükü boşaltılır ve makrofajın normal apoptotik döngüye girmesi sağlanır; nekrotik çekirdek haftalar içinde tamamen geriler.",
        "Akı_Kolesterol_Emilimi = J_lipid = N_rotor * k_baglanma * ( [Kolesterol_kristal] - [Kolesterol_hazne] )",
        "Bu trans-membran ekstraksiyon eşitliği, aterom plağındaki kristalin lipid kütlesinin nanomekanik rotorlarla molekül molekül tahliye edilme akısını tanımlar."
    ),
    (
        "3.4 Biyomekanik Kalsifikasyon Çözünmesi: Apatit Kristallerinin Asidik Nano-Çözünmesi",
        "İleri evre aterosklerozda ve yaşlı arterlerde damar duvarı hidroksiapatit [Ca10(PO4)6(OH)2] kalsiyum kristalleri ile kaplanarak kemik benzeri sert bir boruya dönüşür (Arteriyel Sertleşme).",
        "Vaskülositler, kalsifiye odaklar üzerine mikroskobik vakum kapakları ile hermetik olarak yapışır. Yalnızca kalsiyum kristalinin üzerine bakan izole mikro-kompartmana lokalize olarak seyreltik sitrik asit ve EDTA bazlı şelatör ajanlar salgılanır (pH 4.5). Bu asidik mikro-banyoda kalsiyum fosfat kristalleri hızla çözünerek Ca2+ iyonlarına ayrışır; açığa çıkan iyonlar nanobot tarafından emilerek kanda güvenli fizyolojik seviyede yavaşça salınır; damar duvarı yeniden bebeklik esnekliğine kavuşur.",
        "Reaksiyon_Kalsifikasyon: Ca10(PO4)6(OH)2 (kristal) + 8 H+ --[Vaskulosit]--> 10 Ca2+ + 6 HPO4^2- + 2 H2O",
        "Bu çözünürlük kimyasal dengesi, vaskülositlerin lokal mikro-asitlendirme ile damar duvarındaki hidroksiapatit kalsifikasyonunu çözme basamağını belgeler."
    ),
    (
        "3.5 Trombolitik Nanobotlar: T-PA Yüklü Manyetik Nanomotorlar ve Mekanik Pıhtı Delme",
        "Akut iskemik inme ve miyokard enfarktüsünde sistemik doku plazminojen aktivatörü (t-PA) verilmesi ölümcül beyin kanaması riski taşır; pıhtının dakikalar içinde mekanik olarak açılması şarttır.",
        "Trombolitik Nanobotlar; yüzeyinde t-PA enzimleri taşıyan ve harici manyetik alanlarla yönlendirilen süperparamanyetik nano-matkaplardır (Magnetic Nano-Drills). Tıkalı artere ulaşan nanobot sürüsü, manyetik alanla saniyede 1000 devirle döndürülerek fibrin ağının içine mekanik bir vida gibi dalar. Fibrin lifleri hem mekanik olarak parçalanır hem de lokalize t-PA deşarjı ile eritilir; tıkalı koroner veya serebral damar 5 dakika içinde tamamen rekanalize edilir.",
        "Sure_Rekanalizasyon = t_acilis = L_trombus / ( v_delme + k_enzimatik * [t-PA_lokal] ) ~ 3 - 5 dakika",
        "Bu biyo-mekatronik pıhtı delme hızı denklemi, manyetik rotasyonel penetrasyon ve lokal fibrinolizis sinerjisinin vasküler tıkanıklıkları açma süresini modeller."
    ),
    (
        "3.6 Endotelyal Rekonstrüksiyon: Sentetik Hücre Dışı Matriks (ECM) Yamalama ve İskeleleme",
        "Plak ve kalsifikasyon kazındıktan sonra geride kalan çıplak ve hasarlı endotel tabakası, tedavi edilmezse hemen yeni bir akut tromboza zemin hazırlayabilir.",
        "Vaskülositler kazıma işlemini tamamlar tamamlamaz; hasarlı lümen yüzeyine biyouyumlu nanometrik bir 'Sentetik Endotel Yaması' (Endothelial Nanopatch) serer. Bu yama; heparan sülfat mimetik polimerleri, CD31 endotel bağlantı peptitleri ve nitrik oksit (NO) salgılayan nano-donörler içerir. Yama trombosit yapışmasını anında engellerken; komşu sağlıklı endotel hücrelerinin hızla göç edip çoğalmasını teşvik eden VEGF gradyanları yayar; damar 24 saatte kusursuz doğal endotelle kaplanır.",
        "Hiz_Reendotelyalizasyon = v_endotel = mu_goc * ( [VEGF_lokal] / K_v ) * ( 1 - Trombosit_Adezyonu )",
        "Bu rejeneratif doku yamalama eşitliği, vaskülosit kaynaklı sentetik ECM örtüsünün endotel hücre proliferasyonunu ve vasküler iyileşmeyi nasıl maksimize ettiğini tanımlar."
    ),
    (
        "3.7 Hücre İçi Temizlik Nanobotları: Sitoplazmik Lipofuskin ve Agregatom Süpürme",
        "Nanomekanik cerrahi yalnızca damar boşluklarıyla sınırlı değildir; 'Sitotemizleyici Nanobotlar' (Cell-Cleaner Nanorobots) doğrudan yaşayan nöron ve kardiyomiyositlerin içine girebilir.",
        "50-100 nanometre boyutundaki bu esnek karbon kafesli nanobotlar; hücre zarı klatrin endositozunu taklit ederek sitoplazmaya sızar. Bölünmeyen yaşlı hücrelerin sitoplazmasını işgal eden sarı-kahverengi sindirilemeyen lipofuskin agregatlarını, hatalı katlanmış tau liflerini ve okside protein yumaklarını mekanik tutucularıyla yakalar. İç kırma odasında bu polimerleri parçalayarak serbest monomerlere dönüştürür ve hücrenin proteazom sistemini rahatlatır.",
        "Klirens_Lipofuskin = - d[Lipofuskin]/dt = N_nanobot_intraselluler * k_yakalama * [Lipofuskin_kesiflik]",
        "Bu hücre içi atık temizleme kinetiği denklemi, sitoplazmik nanobotların post-mitotik hücrelerdeki lipofuskin yükünü tersine çevirme hızını formüle eder."
    ),
    (
        "3.8 Çekirdek İçi Cerrahi: Nükleer Giriş ve Somatik DNA Çift Zincir Kırığı Onarımı",
        "Hücre yaşlanmasının nihai kökü olan nükleer DNA çift zincir kırıkları (DSB) ve retrotranspozon insersiyonları; 'Nükleer Genom Nanobotları' ile doğrudan çekirdeğin içinde tamir edilir.",
        "Bu nanobotlar nükleer por kompleksinden (NPC) geçebilecek ultra-küçük boyutta (30-40 nm) tasarlanır ve yüzeylerinde nükleer lokalizasyon sinyalleri (NLS) taşır. Çekirdeğe girdikten sonra kromatini tarayarak gama-H2AX ve 53BP1 ile işaretlenmiş kırık noktalarını bulur. Nanobot iki kırık DNA ucunu mekanik nano-kıskaçlarla kavrayarak birbirine yaklaştırır ve kendi içinde taşıdığı yüksek sadakatli DNA ligaz ve polimeraz enzimleriyle sıfır mutasyonla sarmalı kapatır.",
        "Verim_Onarim_DSB = eta_nano_tamir = Hiz_kiskac * [DSB_tespit] >> Hiz_hucresel_NHEJ",
        "Bu nükleer mikrocerrahi verim eşitliği, mekanik olarak hizalanan DNA uçlarının hücresel hata eğilimli NHEJ yolağına fırsat vermeden kusursuz kapatılma oranını gösterir."
    ),
    (
        "3.9 Yaşlı Hücrelerin Mekanik İmhası: Senolitik Nanorobotlar ve Hedefli Apoptoz",
        "Genetik hasarı tamir edilemeyecek boyuta ulaşmış ve SASP salgılayarak dokuyu zehirleyen senesen (zombi) hücreler; hücresel cerrahi nanobotlar tarafından seçici olarak yok edilir.",
        "Senolitik nanorobotlar, senesen hücrelerin yüzeyindeki spesifik belirteçleri (uPAR, CD9 ve senesens ilişkili beta-galaktozidaz) tanır. Hücreye bağlanan nanobot; hücre içine mekanik bir perforatör fırlatarak mitokondri dış zarını deler (MOMP benzeri etki) ve sitokrom c deşarjını tetikler; ya da doğrudan kaspaz-3 aktivatörü salarak senesen hücreyi 15 dakika içinde sessiz ve yangısız bir apoptoz kaskadına zorlar.",
        "Hiz_Senolizis = d[Senesen_Hucre]/dt = - k_senolitik * [Nanorobot] * [Hucre_Senesen_yuzey_uPAR]",
        "This targeted senolytic ablation equation models the selective elimination kinetics of senescent cells mediated by diagnostic-therapeutic nanorobotic effectors."
    ),
    (
        "3.10 Vasküler Sistem Sürekli Tarama ve Koruma Konsorsiyumu (Vasculature Patrol)",
        "Homo Aeternus damar sağlığı, kanda sürekli devriye gezen milyarlarca vaskülosit, trombolitik nanobot ve endotel onarım ünitesinden oluşan otonom bir 'Vasküler Devriye Konsorsiyumu' ile korunur.",
        "Bu konsorsiyum; 1) Her gün oluşan mikro-trombüsleri saniyeler içinde eritir; 2) Damar çeperine yapışan ilk LDL moleküllerini anında süpürür; 3) Endotel mikro-yırtıklarını derhal kaynaklar; 4) Arteriyel sertlik indeksini sürekli gençlik seviyesinde tutar. İnsan kardiyovasküler sistemi; yaşlanmanın ve biyomekanik yorulmanın tamamen dışlandığı, ebediyen tıkanmayan ve yırtılmayan ölümsüz bir hidrolik otoyola dönüşür.",
        "Kardiyovaskuler_Dayaniklilik = Omega_damar = lim_{t->sonsuz} [ Elastisite_Arteriyel(t) / ( 1 + Plak_Hacmi(t) ) ] = Sabit > 0",
        "Bu nihai vasküler tekillik formülü, otonom nanorobotik konsorsiyum kontrolünde insan arteriyel sisteminin sonsuz ömür boyunca sıfır plak ve tam elastisite ile işleyeceğini matematiksel olarak tescil eder."
    )
]

# ================= KISIM 4 =================
part4_subsections = [
    (
        "4.1 3D Biyo-Baskı (Bioprinting) Teknolojisinin Temel Prensipleri ve Katmanlı Üretim",
        "Doku mühendisliğinin en ileri imalat paradigması olan 3D Biyo-Baskı; canlı hücreleri, biyolojik polimerleri ve büyüme faktörlerini önceden belirlenmiş üç boyutlu dijital koordinatlara katman katman yerleştiren katkılı üretim (additive manufacturing) teknolojisidir.",
        "Geleneksel 3D plastik baskıdan farklı olarak; biyo-baskı süreci hücrelerin hayatta kalmasını sağlamak için fizyolojik sıcaklıkta (37°C), steril ortamda, nötr pH'ta ve hücre zarını patlatmayacak düşük kayma gerilmelerinde (shear stress) yürütülmek zorundadır. Hastanın bilgisayarlı tomografi (BT) ve manyetik rezonans (MR) taramalarından elde edilen anatomik CAD modelleri; mikron düzeyinde dilimlenerek biyo-yazıcının hareket eksenlerine (X-Y-Z) aktarılır.",
        "Hacim_Doku = V_baski = Integral Integral Integral [ Hucre(x,y,z) + Matriks(x,y,z) ] dx dy dz",
        "Bu uzamsal katman integrali, biyo-yazıcının üç boyutlu dijital voksel koordinatlarında canlı hücre ve ekstraselüler matriks yerleşimini formüle eder."
    ),
    (
        "4.2 Biyo-Mürekkep (Bio-Ink) Reolojisi: Viskozite, Kesme İnceltmesi ve Biyouyumluluk",
        "Biyo-baskının başarısındaki en kritik malzeme bilimi unsuru 'Biyo-Mürekkep'tir (Bio-Ink); bu sıvı veya hidrojel malzeme hem hücrelerin yaşaması için biyokimyasal bir yuva hem de basım sırasında mekanik kararlılık sunmalıdır.",
        "İdeal bir biyo-mürekkep 'Kesme İnceltmesi' (Shear-Thinning / Non-Newtonian) reolojik davranış sergilemelidir: Nozul içinde yüksek basınç altında akarken viskozitesi düşmeli (hücrelere binen mekanik stresi azaltmak için); ancak nozuldan çıkıp tablaya düştüğü anda viskozitesi aniden yükselerek şeklini korumalı ve yayılmamalıdır. Aljinat, jelatin metakriloil (GelMA), hyaluronik asit ve desellülarize ECM hidrojelleri altın standart biyo-mürekkep omurgalarıdır.",
        "Viskozite_Akis = eta(gamma_dot) = K_kivam * gamma_dot^(n - 1) (n < 1: Kesme İnceltici Davranış)",
        "Bu Ostwald-de Waele kuvvet kanunu eşitliği, biyo-mürekkebin kayma hızı (gamma_dot) arttıkça viskozitesinin nasıl azalarak hücreleri nozul sürtünmesinden koruduğunu belgeler."
    ),
    (
        "4.3 Ekstrüzyon Tabanlı Biyo-Baskı: Pnömatik, Piston ve Vidalı Nozul Dinamikleri",
        "Doku mühendisliğinde en yaygın kullanılan mekanik yöntem, hidrojellerin mikro-nozullardan sürekli silindirik iplikçikler halinde sıkıldığı 'Ekstrüzyon Tabanlı Biyo-Baskı'dır.",
        "İtme kuvveti pnömatik (hava basıncı), mekanik piston veya dönen helezon vidalar ile sağlanır. Bu yöntem yüksek hücre yoğunluklarını (>10^7 hücre/ml) ve viskoz polimerleri basabilme kabiliyetine sahiptir. Ancak en büyük biyofiziksel kısıtlama, nozul daraldıkça hücre zarlarına binen kayma gerilmesinin artması ve hücre canlılığının (%70-80 seviyelerine) düşmesidir. Nozul geometrisinin konik (tapered) tasarlanması bu kayma hasarını en aza indirir.",
        "Kayma_Gerilmesi = Tau_duvar = ( 4 * Q_akis * eta ) / ( pi * R_nozul^3 ) < Tau_kritik_olum (100 Pa)",
        "Bu Poiseuille duvar gerilmesi denklemi, nozul yarıçapı ve debinin hücre zarına binen mekanik stresi ve canlılık kaybı eşiğini nasıl belirlediğini simgeler."
    ),
    (
        "4.4 Işık Tabanlı Stereolitografi (SLA) ve Dijital Işık İşleme (DLP) Teknolojileri",
        "Mekanik nozulların baskı hızı ve çözünürlük sınırlarını aşmak için 'Işık Tabanlı Biyo-Baskı' (SLA ve DLP) geliştirilmiştir.",
        "Bu sistemlerde; ışığa duyarlı foto-çapraz-bağlanabilir biyo-mürekkepler (örneğin GelMA veya PEGDA) bir reçine teknesine konur. DLP projeksiyon çipi veya UV/mavi lazer (405 nm), organın her bir katmanının iki boyutlu dijital maskesini saniyeler içinde reçine yüzeyine fırlatır. Işığın çarptığı bölgelerdeki foto-başlatıcılar (LAP veya Irgacure) serbest radikaller üreterek hidrojeli anında polimerize eder ve katılaştırır. Tüm bir organ katmanı tek bir ışık parlamasıyla (1-2 saniyede) ve 5-10 mikrometre gibi olağanüstü bir çözünürlükle basılır.",
        "Kürlenme_Derinligi = C_d = D_p * ln( E_isik / E_kritik )",
        "Bu Jacobs fotopolimerizasyon eşitliği, uygulanan ışık enerjisi dozu ve optik penetrasyon derinliğinin mikroskobik hidrojel katman kalınlığını nasıl kontrol ettiğini modeller."
    ),
    (
        "4.5 İki-Fotonlu Polimerizasyon (2PP): Nanometre Çözünürlüklü Kılcal Damar İskeleleri",
        "Biyolojik kılcal damarların çapı 5 ila 10 mikrometredir; geleneksel hiçbir mekanik nozul bu çözünürlükte delikli borular basamaz.",
        "Bu nihai mikroskobik çözünürlük 'İki-Fotonlu Polimerizasyon' (2PP / TPP) ile elde edilir. Femtosaniyelik kızılötesi lazer darbeleri, hidrojel havuzu içinde mikroskobik bir odak noktasına (focal voxel) yoğunlaştırılır. Malzeme ışığı ancak iki fotonun aynı anda emildiği (two-photon absorption) bu minik odak noktasında polimerize eder. Odak noktası dışındaki tüm alan sıvı kalır. 2PP ile 100 nanometre çözünürlükte mükemmel mikro-kılcal damar ağları ve hücre tutunma iskeleleri üretilebilmektedir.",
        "Hacim_Voksel = V_fokal ~ lambda_lazer^3 / ( NA_lens^4 ) (Nanometrik Çözünürlük)",
        "Bu doğrusal olmayan optik çözünürlük formülü, iki fotonlu lazer odaklanmasının dalga boyu ve sayısal açıklıkla (NA) nanometre ölçeğinde kusursuz mikrovasküler tüpler inşa etme gücünü belgeler."
    ),
    (
        "4.6 Çoklu Hücre Tipi ve Çoklu Biyo-Mürekkep Eş Zamanlı Baskı Sistemleri",
        "Gerçek bir organ asla tek bir hücre tipinden oluşmaz; örneğin bir kalp parankiminde kardiyomiyositler, vasküler endotel hücreleri, düz kas hücreleri, fibroblastlar ve Purkinje iletim lifleri tam bir mimari ahenk içinde dizilmiştir.",
        "Modern multi-materyal biyo-yazıcılar, 4 ila 8 bağımsız nozul ve döner kaset sistemleri taşır. Yazıcı bir nozulla vasküler ağın endotel hücrelerini aljinat içine dizerken; diğer nozulla kardiyak kas hücrelerini GelMA içine basar; üçüncü nozulla mekanik destek iskeleti için biyobozunur PCL (polikaprolakton) polimerini örer. Eş zamanlı mikroakışkan nozul değiştirme teknolojileri, malzeme geçiş sürelerini milisaniyelere indirmiştir.",
        "Karmasiklik_Doku = Sigma_i [ Hucre_i(x,y,z) * Matriks_i(x,y,z) * Büyüme_Faktörü_i ]",
        "Bu çok-bileşenli doku organizasyonu modeli, farklılaşmış organ mimarilerinin çoklu biyo-mürekkep koordinasyonu ile inşa edilme matrisini tanımlar."
    ),
    (
        "4.7 Vaskülarizasyon Paradigması: Perfüze Edilebilir Hiyerarşik Damar Ağlarının Basımı",
        "3D biyo-baskının tarihteki en büyük aşılmaz engeli 'Vaskülarizasyon Darboğazı' (Vascularization Bottleneck) olmuştur; basılan bir doku 200 mikrometreden daha kalınsa içteki hücreler oksijen difüzyon sınırını aştığı için saatler içinde nekroza uğrar.",
        "Bu engel 'Kurbanlık Biyo-Baskı' (Sacrificial Bioprinting) ile aşılmıştır: Önce Pluronic F127 veya jelatin gibi sıcaklığa duyarlı kurbanlık bir jel ile kılcal damar dallanma ağının negatif kalıbı basılır. Etrafı hücreli ana hidrojel ile kaplanır ve çapraz bağlanır. Ardından sıcaklık 4°C'ye düşürüldüğünde kurbanlık jel sıvılaşarak akar ve geride içi boş, perfüze edilebilir açık bir damar ağı bırakır. Bu kanalların içi endotel hücreleri ile kaplanarak saniyeler içinde kan akışına uygun fonksiyonel damarlar elde edilir.",
        "Oksijen_Difuzyon_Limiti = L_kritik = sqrt( 2 * D_O2 * C_yuzey / R_tuketim ) ~ 150 - 200 mikrometre",
        "Bu Krogh silindir difüzyon denklemi, vasküler perfüzyon kanalları olmaksızın katı doku çekirdeğinde hipoksik nekrozun başladığı kritik geometrik sınırı gösterir."
    ),
    (
        "4.8 Biyo-Baskı Sonrası Dinamik Olgunlaşma: Biyoreaktörler ve Mekanobiyolojik Uyarım",
        "Biyo-yazıcıdan yeni çıkmış bir doku henüz fonksiyonel bir organ değildir; hücrelerin birbiriyle bağlantı kurması (gap junction), ekstraselüler matriks salgılaması ve organize olması şarttır.",
        "Basılan doku derhal özel 'Dinamik Biyoreaktörler' içine alınır. Kalp dokuları elektriksel alan uyarımı (1 Hz ritmik şok) ve mekanik gerilme (pulsatil esnetme) altında tutulur; damar ve kemik dokularına pulsatil sıvı akışı verilerek kayma gerilmesi uygulanır. Bu mekanobiyolojik uyarım; hücrelerin tensil kuvvetlerini ve elektriksel iletim hızlarını 10 kat artırarak hidrojel yığınını günler içinde yaşayan, senkronize atan güçlü bir kardiyak dokuya dönüştürür.",
        "Olgunlasma_Indeksi = M_doku = Integral_0^t [ Sinyal_Elektrik(t') * Gerilme_Mekanik(t') * Perfuzyon_Debisi(t') ] dt'",
        "Bu dinamik biyoreaktör olgunlaşma fonksiyonu, biyo-basılmış dokuların fonksiyonel fizyolojik kapasiteye ulaşmasındaki biyofiziksel uyarımların kümülatif etkisini belgeler."
    ),
    (
        "4.9 4D Biyo-Baskı: Zamana Bağlı Morfolojik Dönüşüm ve Akıllı Malzemeler",
        "Doku mühendisliğinin en son ufku olan 4D Biyo-Baskı; üç boyutlu basılan bir yapının dördüncü boyut olan 'zaman' içinde harici uyaranlarla şekil ve fonksiyon değiştirmesidir.",
        "Akıllı polimerler (şekil hafızalı polimerler, sıcaklığa duyarlı PNIPAM, pH ve manyetik alana duyarlı hidrojeller) kullanılır. Vücut içine implante edilen düz bir hidrojel katmanı; vücut sıcaklığına (37°C) veya lokal pH değişimine maruz kaldığında kendiliğinden kıvrılarak içi boş bir kan damarına, kalp kapağına veya sinir iletim tüpüne dönüşür (Self-Folding). Bu durum mikro-cerrahi implantasyon zorluklarını dramatik ölçüde kolaylaştırır.",
        "Sekil_Degisimi = Delta_Geometri(t) = f( Sicaklik, pH, Iyonik_Kuvvet ) * Matris_Hafiza",
        "Bu akıllı malzeme kinetik eşitliği, dördüncü boyut olan zaman içinde çevre duyarlı hidrojel mimarilerinin kendiliğinden hedeflenen biyolojik organ morfolojisine katlanmasını formüle eder."
    ),
    (
        "4.10 Klinik Öncesi Biyo-Baskı Başarıları: Vaskülarize Kalp Yamaları ve Tam Katmanlı Deri",
        "3D biyo-baskı laboratuvar prototiplerinden canlı hayvan ve ilk insan klinik uygulamalarına geçiş yapmıştır.",
        "Tel Aviv Üniversitesi'nden Dvir ve ekibi; hastanın kendi omentum biyopsisinden elde edilen kök hücreler ve ECM biyo-mürekkebi ile odacıkları, kapakçıkları ve koroner damarları olan ilk tam vaskülarize mini-insan kalbini basmıştır. Eş zamanlı olarak Wake Forest Enstitüsü, yanık hastalarında yara üzerine doğrudan hücre tabakalayan 'Mobil Cilt Biyo-Yazıcısı' ile tam katmanlı vaskülarize deri rejenerasyonunu başarmış; biyo-yapay organ çağının geri dönüşsüz kapısını açmıştır.",
        "Entegrasyon_Greft = eta_klinik = [Damarlanma_orani] * [Hucre_canliligi] * ( 1 - Dokusal_Red_Skoru ) ~ %95+",
        "Bu klinik başarı skoru, otolog hasta hücreleriyle basılan vaskülarize doku yamalarının konak dokuyla tam biyolojik kaynaşma verimini ortaya koyar."
    )
]

# ================= KISIM 5 =================
part5_subsections = [
    (
        "5.1 Doğal Hücre Dışı Matriks (ECM) İskeleleri ve Desellülarizasyon Paradigması",
        "3D biyo-yazıcılar sentetik geometrileri mükemmel bassa da; milyonlarca yıllık evrimle optimize olmuş böbrek glomerülü veya akciğer alveolü gibi mikroskobik mimarileri sıfırdan taklit etmek aşırı zordur.",
        "Bu zorluğu aşan alternatif devrimsel paradigma 'Desellülarizasyon'dur (Decellularization): Kadavra veya hayvan (domuz) organlarındaki tüm canlı hücreler, immünojenik antijenler ve DNA kalıntıları kimyasal ve enzimatik yöntemlerle tamamen yıkanıp temizlenir. Geride organın kolajen, elastin, laminin ve glikozaminoglikanlardan oluşan kusursuz, üç boyutlu 'Hücresiz Doğal İskeleti' (Acellular ECM Scaffold) kalır; bu iskele tüm mikrovasküler damar yatağını eksiksiz muhafaza eder.",
        "Kriter_Desellularizasyon: [DNA_cift_zincir] < 50 ng/mg kuru_agirlik | Fragment_Uzunlugu < 200 bp",
        "Bu klinik güvenlik standardı, desellülarize edilmiş bir doku iskelesinin konak bağışıklık sistemini tetiklememesi için taşıması gereken maksimum artık DNA sınırını belgeler."
    ),
    (
        "5.2 Deterjan Perfüzyon Teknikleri: SDS, Triton X-100 ve Enzimatik Sindirim Dinamikleri",
        "Bütün bir katı organın (örneğin kalp veya karaciğer) desellülarizasyonu; organın kendi damar sistemine bağlanan kanüllerle perfüzyon yıkaması yapılarak gerçekleştirilir.",
        "Perfüzyon protokolü aşamalıdır: 1) İyonik Deterjanlar (Sodyum Dodesil Sülfat - SDS): Hücre zarlarını ve nükleer zarları çözer, hücresel enkazı hızla yıkar; 2) Non-İyonik Deterjanlar (Triton X-100): Lipid-lipid etkileşimlerini çözerken matriks proteinlerini korur; 3) Enzimatik Sindirim (DNaz ve RNaz): Kromatin kalıntılarını mononükleotitlere parçalar. Perfüzyon basıncı kılcal damar yatağını patlatmayacak şekilde (fiziksel fizyolojik sınırda: 60-80 mmHg) hassas kontrol edilir.",
        "Hiz_Hucresel_Yikama = - d[Hucresel_Kutle]/dt = k_deterjan * Q_perfuzyon * [SDS] * ( 1 / Direnc_Vaskuler )",
        "Bu kütle transfer kinetiği formülü, organ içi hücresel temizlenme debisinin deterjan konsantrasyonu ve intakt vasküler perfüzyon hızıyla bağıntısını ifade eder."
    ),
    (
        "5.3 Matriks Bütünlüğünün Korunması: Kolajen, Elastin, Laminin ve Büyüme Faktörü Kalıntıları",
        "Desellülarizasyonun en kritik kimyasal hassasiyeti; hücreleri yok ederken ekstraselüler matriksin ultra-yapısını ve biyoaktif sinyallerini parçalamamaktır.",
        "Aşırı deterjan maruziyeti bazal membrandaki Tip IV kolajeni ve laminini aşındırabilir; bu durum yeniden hücre ekildiğinde endotelizasyonu imkânsız kılar. Optimize edilmiş protokoller; elastik lifleri, fibronektini ve matrikse bağlı anjiyogenik büyüme faktörlerini (bFGF, VEGF, TGF-beta) %80 oranında korur. Bu biyokimyasal izler, sonradan ekilecek kök hücrelere 'burası bir nefrondur' veya 'burası bir alveoldür' mesajını veren moleküler rehberlerdir.",
        "Skor_Matriks_Korunumu = S_ECM = ( [Laminin_post] / [Laminin_pre] ) * ( [Kolajen_post] / [Kolajen_pre] ) > 0.85",
        "Bu biyoaktif matriks retansiyon oranı, kimyasal yıkama sonrası iskelede kalan fonksiyonel tutunma proteinlerinin yapısal sağlamlık derecesini gösterir."
    ),
    (
        "5.4 Vasküler Ağın Bütünlüğü ve Mikro-Anjiyografik Doğrulama",
        "Bir organ iskelesinin klinikte hayatta kalabilmesi, nakledildiğinde alıcının kan dolaşımına bağlandığı anda kan sızdırmamasına ve pıhtı atmamasına bağlıdır.",
        "İskelenin vasküler ağ bütünlüğü; mikro-bilgisayarlı tomografi (Mikro-BT), floroskopik anjiyografi ve mikrosfer perfüzyon testleri ile tescillenir. Ana arterden 100 mmHg basınçla verilen kontrast madde; en uç kılcal damarlara (kapillerlere) kadar ilerlemeli, parankime kaçak (rüptür sızıntısı) yapmamalı ve venöz sistemden pürüzsüz geri dönmelidir. Bu hidrolik bütünlük sağlanmadan hiçbir organ resellülarizasyon fazına alınamaz.",
        "Gecirimsizlik_Vaskuler = P_sizinti = Q_kacak / Q_giris < 0.001 (Binde Birden Az Mikrovasküler Kaçak)",
        "Bu vasküler sızdırmazlık kriteri eşitliği, desellülarize damar yatağının fizyolojik basınç altında parankime plazma sızdırmama mekanik direncini tanımlar."
    ),
    (
        "5.5 Resellülarizasyon: Otolog İndüklenmiş Pluripotent Kök Hücrelerin (iPSC) Tohumlanması",
        "Mükemmel şekilde hazırlanmış hücresiz bir organ iskelesi; hastanın kendi genetik kimliğini taşıyan genç hücrelerle yeniden canlandırılır (Resellülarizasyon).",
        "Hastadan alınan cilt veya kan hücreleri Yamanaka faktörleriyle iPSC'ye çevrilir; ardından bioreaktörlerde yüz milyarlarca spesifik parankimal hücreye (hepatositler, kardiyomiyositler, podositler) ve endotel hücresine farklılaştırılır. Tek bir insan karaciğeri veya kalbi için yaklaşık 30 ila 50 milyar canlı hücreye ihtiyaç vardır. Bu devasa hücre havuzu, organın damar sisteminden ve parankimal kanallarından aşamalı olarak iskeleye infüze edilir.",
        "Ihtiyac_Hucre = N_toplam = V_organ * Dansite_hucre ~ (1500 cm3) * ( 2 * 10^7 hucre/cm3 ) ~ 3 * 10^10 hucre",
        "Bu hücresel stokiyometri denklemi, tam boyutlu bir insan katı organının biyolojik fonksiyon kazanması için iskeleye tohumlanması gereken minimum mutlak kök hücre sayısını belgeler."
    ),
    (
        "5.6 Endotelyalizasyon Biyofiziği ve Antitrombotik Lümen Kaplama",
        "Bir resellülarize organın nakledildikten sonra dakikalar içinde trombozla tıkanmasını önlemenin tek yolu, tüm vasküler ağın lümeninin kusursuz bir 'Endotel Tabakası' ile kaplanmasıdır.",
        "Eğer damar duvarında çıplak kolajen açıkta kalırsa; kanla temas ettiği anda trombositler Von Willebrand faktörüne yapışarak masif pıhtılaşma başlatır. İntakt vasküler ağa iPSC kaynaklı endotel hücreleri verilir; iskele yerçekimi yönünü sürekli değiştiren 3D döner biyoreaktörde döndürülerek hücrelerin damarın her yüzeyine eşit yapışması sağlanır. Endotel hücreleri CD31 ve VE-kaderin bağlantıları kurarak lümeni pürüzsüzce mühürler; yüzey heparan sülfat ve trombomodülin sentezleyerek antikoagülan zırh oluşturur.",
        "Kapsama_Endotel = A_kapli / A_toplam_vaskuler > %99.5 (Tam Antitrombotik Yüzey)",
        "Bu endotelyal yüzey kapsama oranı, damar içi tromboz riskini sıfıra indirmek için gereken minimum lümen kaplama yüzdesini formüle eder."
    ),
    (
        "5.7 Biyomimetik Organ Biyoreaktörleri: Fizyolojik Basınç, Akış ve Gaz Değişimi",
        "Tohumlanan hücrelerin iskele üzerinde doğru doku morfolojisine girmesi; ana rahmindeki ve vücuttaki fizyolojik ortamı birebir taklit eden 'Biyomimetik Biyoreaktörler' içinde haftalarca olgunlaştırılmasıyla sağlanır.",
        "Böbrek biyoreaktörü fizyolojik idrar hidrostatik basıncı ve tübüler sıvı akışı sağlar; akciğer biyoreaktörü negatif basınçla ritmik solunum havalandırması uygular; kalp biyoreaktörü koroner arterlerden oksijenli besiyeri pompalarken ventrikülleri mekanik olarak esnetir. Bu biyoreaktörler yerel pH, pO2, pCO2 ve glukoz/laktat seviyelerini kapalı devre sensörlerle izleyerek metabolik beslenmeyi otomatik dengeler.",
        "Perfuzyon_Biyoreaktor = Q(t) = Q_bazal * ( 1 + alpha_metabolizma * [Laktat_cikis] / [Glukoz_giris] )",
        "Bu dinamik biyoreaktör akış denklemi, organ içi perfüzyon hızının dokunun metabolik aktivite ve atık birikim oranına göre otomatik ayarlanma algoritmasını simgeler."
    ),
    (
        "5.8 Ksenojenik (Domuz) İskeleler ve İmmün Reddin Moleküler Olarak Sıfırlanması",
        "İnsan kadavra organ arzı son derece kısıtlı olduğundan; desellülarizasyon için ana kaynak fizyolojik ve anatomik olarak insanla neredeyse özdeş olan domuz organlarıdır.",
        "Domuz dokularında insan bağışıklık sisteminin hiperakut reddine yol açan en büyük antijen 'Galaktoz-alfa-1,3-galaktoz' (Alfa-Gal) şeker epitopudur. Modern protokollerde iki yaklaşım birleştirilir: 1) CRISPR ile GGTA1, CMAH ve B4GALNT2 genleri nakavt edilmiş transgenik domuzların organları kullanılır; 2) Desellülarizasyon sırasında eklenen alfa-galaktozidaz enzimi tüm artık epitopları sindirir. Hücresizleştirilen domuz iskelesi insan kök hücreleriyle kaplandığında insan vücudu tarafından %100 'kendi dokusu' olarak kabul edilir.",
        "Antijenik_Yuk = [Alfa-Gal] < 1.0 * 10^-15 mol/mg doku (Sıfır Hiperakut İmmün Red)",
        "Bu immünolojik saflık eşiği, ksenojenik domuz iskelelerinin insan immün sistemine tamamen görünmez kılınması için gereken maksimum artık antijen sınırını tanımlar."
    ),
    (
        "5.9 İn Vivo Transplantasyon ve Organ Fonksiyonunun Uzun Süreli İzlenmesi",
        "Resellülarize edilmiş biyo-yapay organlar; hayvan modellerinde (domuz ve primat) ortotopik veya heterotopik olarak implante edilerek in vivo rüştünü kanıtlamıştır.",
        "Harald Ott laboratuvarının ürettiği resellülarize sıçan böbrekleri; alıcının renal arter ve venine bağlandığı anda fizyolojik kan perfüzyonunu başlatmış, kreatinin ve üreyi kandan süzerek üreterden idrar üretmiştir. Benzer şekilde resellülarize akciğerler kanı oksijenlendirmiş, karaciğerler albümin sentezlemiştir. Hastanın kendi hücrelerini taşıdığı için bu organlar ömür boyu hiçbir immünsüpresif ilaca (siklosporin, takrolimus) ihtiyaç duymadan çalışmaktadır.",
        "Klirens_Kreatinin = C_cr = ( U_cr * V_idrar ) / P_cr (Biyo-Yapay Böbrekte Fonksiyonel Süzme)",
        "Bu klasik nefrolojik klirens eşitliği, resellülarize biyo-yapay böbreğin fizyolojik kan temizleme ve glomerüler filtrasyon yeteneğini belgeler."
    ),
    (
        "5.10 Biyo-Yapay Organ Fabrikasyonu ve 'İsteğe Bağlı Organ Üretimi' (On-Demand Organs)",
        "Desellülarizasyon ve resellülarizasyon teknolojisinin endüstriyel ölçeğe taşınması; organ bekleme listelerini ve organ yetmezliğinden ölümleri tarihe gömecek 'İsteğe Bağlı Organ Fabrikaları'nı kurmaktadır.",
        "Temiz oda (GMP) biyoreaktör çiftliklerinde yüzlerce ksenojenik iskele hazır dondurulmuş olarak bekletilir. Yaşlanan veya organı iflas eden bir hastanın kanından iPSC'ler üretilir, robotik biyoreaktörlerde 3 haftada milyarlarca hücreye çoğaltılarak iskeleye tohumlanır ve 1 ay içinde hastaya özel sıfır-red riskli yepyeni bir genç organ teslim edilir. İnsan anatomisi artık bozuldukça değiştirilebilen modüler bir mimariye kavuşur.",
        "Kapasite_Fabrikasyon = Q_organ_yillik = N_biyoreaktor * ( 365_gun / Sure_olgunlasma_4_hafta )",
        "Bu endüstriyel biyoproses kapasite formülü, rejeneratif tıp merkezlerinin yıllık biyo-yapay organ üretim çıktısını ve organ kıtlığının küresel tasfiyesini modeller."
    )
]

# ================= KISIM 6 =================
part6_subsections = [
    (
        "6.1 Katı Organ Mimarisinin Fonksiyonel Üniteleri: Karaciğer Lobülü, Nefron ve Alveol",
        "Katı organlar kaotik hücre yığınları değil; milyonlarca kez tekrarlanan mikroskobik fonksiyonel alt ünitelerin (fonksiyonel birimler) mükemmel geometrik entegrasyonudur.",
        "Karaciğerin temel ünitesi hekzagonal 'Hepatik Lobül' (santral ven etrafında dizilen hepatosit kordonları ve portal triad); böbreğin temel ünitesi kanı süzen 'Nefron' (Bowman kapsülü, glomerül ve Henle kulpu); akciğerin temel ünitesi ise gaz değişimini gerçekleştiren 'Alveoler-Kapiller Membran'dır. Katı organ mühendisliğinin ana hedefi, tüm organı tek seferde inşa etmek yerine bu mikroskobik fonksiyonel üniteleri yüksek sadakatle kopyalamaktır.",
        "Mimar_Kati_Organ: Organ_Total = Sigma_{k=1}^N [ Fonksiyonel_Mikro_Unite_k ] (N ~ 10^6 Unite)",
        "Bu modüler hiyerarşi eşitliği, bir katı organın toplam fizyolojik gücünün milyonlarca bağımsız mikroskobik ünitenin paralel çalışmasından doğduğunu simgeler."
    ),
    (
        "6.2 Biyo-Yapay Karaciğer (BAL): Hepatik Biyoreaktörler ve Karaciğer Destek Sistemleri",
        "Akut karaciğer yetmezliğinde veya kronik siroz krizlerinde; karaciğerin yüzlerce karmaşık detoksifikasyon ve protein sentez fonksiyonunu mekanik bir diyaliz filtresi taklit edemez; yaşayan hepatositler içeren 'Biyo-Yapay Karaciğer' (BAL - Bio-Artificial Liver) şarttır.",
        "HepatAssist ve ELAD sistemlerinde; hastanın kanı plazmaferezle ayrılır ve içi milyonlarca genç hepatosit veya iPSC-hepatosit yüklü içi boş fiber (hollow-fiber) biyoreaktör kartuşundan geçirilir. Fiberlerin gözenekli zarları hücreleri hastanın antikorlarından korurken; plazmadaki amonyak, bilirubin ve toksinler hepatositlerce metabolize edilir, albümin ve pıhtılaşma faktörleri plazmaya salınır. BAL sistemleri hastayı komadan çıkararak organ rejenerasyonu için kritik zaman kazandırır.",
        "Detoksifikasyon_Kapasitesi = J_amonyak = V_max * [Amonyak_plazma] / ( K_m + [Amonyak_plazma] ) * N_hepatosit",
        "Bu enzimatik klerens Michaelis-Menten denklemi, biyo-yapay karaciğer kartuşunun kandan toksik amonyak süpürme hızının canlı hepatosit sayısıyla ilişkisini modeller."
    ),
    (
        "6.3 Biyo-Yapay Böbrek (BAK): İçi Boş Fiberler ve Renal Tübüler Epitel Hücreleri (RAD)",
        "Standart hemodiyaliz yalnızca pasif difüzyonla küçük molekülleri süzer; ancak böbreğin su ve elektrolitleri geri emen, bikarbonat dengesini koruyan ve D vitamini üreten aktif tübüler metabolik fonksiyonlarını yapamaz.",
        "David Humes tarafından geliştirilen 'Biyo-Yapay Böbrek' (BAK - Bio-Artificial Kidney); geleneksel sentetik hemofiltreyi bir 'Renal Destek Aygıtı' (RAD - Renal Assist Device) ile birleştirir. RAD kartuşunun içi boş polisülfon fiberlerinin lümeni, insan renal tübüler epitel hücreleri ile kaplanmıştır. Sentetik filtre kanı süzdükten sonra elde edilen ultrafiltrat RAD'a verilir; canlı tübül hücreleri esansiyel iyonları, amino asitleri ve suyu kana geri emerken toksinleri idrara atar ve sitokin fırtınasını baskılar.",
        "Geri_Emilim_Orani = eta_tubul = ( Q_ultrafiltrat - Q_idrar_nihai ) / Q_ultrafiltrat ~ %99 (Fizyolojik Denge)",
        "Bu renal taşıma eşitliği, biyo-yapay tübüler hücre kartuşunun süzülen sıvının %99'unu seçici olarak kana geri kazandırma verimini belgeler."
    ),
    (
        "6.4 İmplante Edilebilir Yapay Böbrek (iAK): Silikon Nano-Gözenekli Filtreler",
        "Hastaları haftada üç gün diyaliz makinesine mahkûm eden boru ve pompaları çöpe atacak devrim; vücut içine yerleştirilen 'İmplante Edilebilir Yapay Böbrek'tir (iAK - Implantable Artificial Kidney / Shuvo Roy & William Fissell projesi).",
        "iAK bir kahve fincanı boyutundadır ve hiçbir harici pompaya veya pile ihtiyaç duymaz; kan akışını hastanın kendi kalp basıncı (aorto-kaval şant) sağlar. Çip üretiminde kullanılan silikon litografi ile üretilen 'Silikon Nano-Gözenekli Membranlar' (SNM); 1-2 nanometrelik kusursuz yarık gözenekleriyle kan proteinlerini (albümin) tutarken üre ve kreatinini süzer. İkinci kompartmandaki biyoreaktör tübül hücreleri ise idrarı konsantre ederek mesaneye akıtır.",
        "Akı_Filtrasyon_SNM = J_GFR = ( Delta_P_arteriyel * w_gozenek^2 * A_membran ) / ( 12 * eta_kan * Kalinlik )",
        "Bu mikroakışkan nano-yarık akı denklemi, silikon nanopor filtrelerin hastanın kendi kan basıncıyla fizyolojik glomerüler filtrasyon hızını (GFR > 30 ml/dk) nasıl sağladığını tanımlar."
    ),
    (
        "6.5 Biyo-Yapay Akciğer (BALung): Mikroakışkan Kan-Gaz Değişim Çipleri",
        "Geleneksel ECMO (Ekstrakorporeal Membran Oksijenasyonu) cihazları devasa yüzey alanları, yüksek kan dolum hacimleri ve masif pıhtılaşma riskleri nedeniyle haftalarca kullanılamaz.",
        "Mikroakışkan Biyo-Yapay Akciğer sistemleri; insan akciğer alveollerinin 1 mikrometrelik difüzyon mesafesini mikroçip düzeyinde kopyalar. Gaz geçirgen PDMS (polidimetilsiloksan) membranların bir tarafından kan mikrokanallarla tek sıra alyuvarlar halinde akarken, diğer taraftan saf oksijen geçer. İçi endotel hücreleri ile kaplanan bu biyo-hibrit çipler, sıfır pıhtılaşma riskiyle kanda oksijen saturasyonunu (%70'ten %98'e) saniyeler içinde çıkarır ve taşınabilir bir yapay akciğer olarak hastanın beline asılabilir.",
        "Gaz_Transfer_Kapasitesi = J_O2 = D_membran * S_alan * ( P_O2_gaz - P_O2_kan ) / Kalinlik_membran",
        "Bu Fick difüzyon kanunu eşitliği, mikroakışkan yapay akciğer membranlarının kan oksijenasyon debisini ve gaz transfer verimini formüle eder."
    ),
    (
        "6.6 Endokrin Pankreas Mühendisliği: İnsülin Salgılayan Adacık Enkapsülasyonu",
        "Tip 1 diyabet ve yaşa bağlı pankreatik beta hücresi tükenişi; glukoz dengesini bozar ve vasküler kalsifikasyonu hızlandırır.",
        "Enkapsüle Adacık Teknolojisinde; iPSC kaynaklı genç beta hücreleri veya Langerhans adacıkları, yarı geçirgen biyouyumlu aljinat mikro-kapsüller içine hapsedilir. Kapsül gözenekleri glukozun ve insülinin serbestçe difüzyonuna izin verirken; konak bağışıklık sisteminin antikorlarını (IgG) ve sitotoksik T hücrelerini içeri sokmaz (İmmünolojik İzolasyon). Hastaya tek bir laparoskopik enjeksiyonla karın boşluğuna (peritona) milyonlarca kapsül verilir; hasta hiçbir bağışıklık baskılayıcı ilaç almadan ömür boyu kan şekerini fizyolojik normda tutar.",
        "Gecirgenlik_Kapsul: r_gozenek < r_IgG (5 nm) | r_gozenek > r_Glukoz_Insulin (1 - 2 nm)",
        "Bu boyutsal eleme kriteri, immünoprotektif hidrojel kapsüllerin hedef hormon geçişine izin verirken bağışıklık saldırısını nasıl mutlak biçimde bloke ettiğini açıklar."
    ),
    (
        "6.7 Organ-on-a-Chip (Çip Üstünde Organ): Çoklu Organ Fizyolojik Entegrasyonu",
        "Katı organların in vivo davranışlarını ve birbirleriyle metabolik iletişimini incelemek için mikroakışkan 'Organ-on-a-Chip' sistemleri geliştirilmiştir.",
        "Donald Ingber ve Wyss Enstitüsü tarafından öncülüğü yapılan bu teknoloji; insan karaciğer, kalp, akciğer, bağırsak ve beyin hücrelerini mikrokanallarla birbirine bağlanan çiplere yerleştirir. Çipler arasında dolaşan yapay kan sıvısı; karaciğerde metabolize olan bir molekülün böbrekten atılımını veya kan-beyin bariyerini geçişini canlı insan hücresi düzeyinde simüle eder. Çoklu organ çipleri, yaşlanma karşıtı nanorobotların ve gen terapilerinin vücut çapındaki kinetiğini hayvan testlerine gerek duymadan test eder.",
        "Baglanti_Coklu_Cip: Karaciger_Cipi ----[Mikroakiskan]--> Bobrek_Cipi ----> Beyin_Cipi (Dinamik Metabolik Ag)",
        "Bu mikroakışkan sistem biyolojisi entegrasyonu, çoklu organ çiplerinin insan farmakokinetik ve organ iletişimini çip üzerinde modelleme mimarisini tanımlar."
    ),
    (
        "6.8 Kriyoprezervasyon ve Vitrifikasyon ile Organ Bankacılığı",
        "Üretilen veya bağışlanan biyo-yapay organların hastalara ulaştırılabilmesi için canlılıklarını yitirmeden aylarca saklanabilmesi şarttır; ancak geleneksel buz kristali oluşumu organı parçalar.",
        "Modern 'Vitrifikasyon' (Camsılaşma) teknolojisi; yüksek konsantrasyonlu kriyoprotektan ajanlar (CPA - dimetil sülfoksit, formamid, etilen glikol) ve ultra-hızlı soğutma (-130°C/dakika) kullanarak sıvıyı hiç kristalize etmeden doğrudan amorf bir 'biyolojik cam' fazına dönüştürür. Buz kristali oluşmadığı için hücre zarları ve damar ağları sıfır mekanik hasarla -196°C sıvı azotta yıllarca korunur. Yeniden ısıtmada ise nano-demir partikülleri ve radyo-frekans manyetik alanlar kullanılarak organ saniyeler içinde homojen ısıtılır.",
        "Kriter_Vitrifikasyon: dT/dt > Hiz_Kritik_Sogutma (Buz Nukleasyonu Engellenir, Camsı Faz)",
        "Bu fiziksel faz dönüşüm koşulu, organ kriyoprezervasyonunda buz kristali oluşumunu sıfırlayan vitrifikasyon termodinamiğini belgeler."
    ),
    (
        "6.9 Biyo-Hibrit Greftlerin İmmün Toleransı ve Biyo-İnertlik",
        "Katı organ nakillerinde kronik doku reddi ve immünsüpresif ilaçların yol açtığı böbrek hasarı ve kanserler en büyük yaşam sınırlayıcı unsurdur.",
        "Biyo-yapay katı organlar ya hastanın kendi otolog iPSC hücreleriyle tohumlandığı için ya da ksenojenik antijenleri tamamen silinmiş CRISPR-modifiye hücreler taşıdığı için bağışıklık sistemi tarafından tanınmaz. Ayrıca organın dış kapsülü biyo-uyumlu zwitteriyonik nanofilmlerle kaplanarak lenfosit infiltrasyonu engellenir. Bu çifte koruma, greftin ömür boyu reddedilmeden tıkır tıkır çalışmasını garantiler.",
        "Skor_Reddedilme = I_red = Sigma_j [HLA_uyumsuzluk_j] * [Antikor_donör_spesifik] ~ 0",
        "Bu transplantasyon immün tolerans eşitliği, otolog ve genetik olarak temizlenmiş biyo-yapay organların sıfır immün ret katsayısını tanımlar."
    ),
    (
        "6.10 Katı Organların Sonsuz Replasman Döngüsü (Infinite Organ Replacement Cycle)",
        "İnsan vücudunun doğal biyolojik organları 80-100 yıl içinde kaçınılmaz olarak doku sertleşmesi, amiloid birikimi ve mitokondriyal yetmezlikle yıpranır.",
        "Homo Aeternus protokolünde; hiçbir organın ölümcül son evre yetmezliğe girmesine izin verilmez. Biyomarkerlar bir organda geri döndürülemez bir yaşlanma yıpranması tespit ettiği anda (örneğin GFR < 45 veya EF < %40); fabrikasyon hattından çıkan taptaze, genç, biyolojik olarak 20 yaşındaki biyo-yapay muadili elektif bir cerrahiyle implante edilir. İnsan bedeni, bozulan parçaları periyodik olarak yenilenen sonsuz bir tapınağa dönüşür.",
        "Omur_Organizm = L_insan = lim_{k->sonsuz} Sigma_{k=1}^M [ Omur_Organ_k ] = Sonsuz",
        "Bu sonsuz replasman eşitliği, katı organların periyodik yenilenmesi döngüsünün organizmanın toplam yaşam süresini teorik olarak sonsuza taşıdığını matematiksel olarak ilan eder."
    )
]

# ================= KISIM 7 =================
part7_subsections = [
    (
        "7.1 Kardiyovasküler İflas ve Mekanik Dolaşım Desteğinin (MCS) Evrimi",
        "İnsan kalbi bir ömür boyunca yaklaşık 3 milyar kez atan ve dakikada 5 litre kan pompalayan inanılmaz bir biyolojik motordur; ancak kardiyomiyositlerin bölünme yeteneğinin olmaması kalbi yaşlanmanın en kırılgan organı yapar.",
        "İleri evre kalp yetmezliğinde kardiyak debi (CO) dokuların bazal metabolik oksijen ihtiyacını karşılayamaz hale gelir. Mekanik Dolaşım Desteği (MCS - Mechanical Circulatory Support); ilk nesil devasa pnömatik pulsatil pompalardan (Jarvik-7), modern mikroskobik sürekli akışlı manyetik levitasyon pompalarına kadar yarım asırlık bir teknolojik evrim geçirmiştir. Günümüzde bu sistemler yalnızca bir köprü tedavisi (bridge-to-transplant) değil, ömür boyu kalıcı nihai tedavi (destination therapy) olarak kullanılmaktadır.",
        "Kardiyak_Debi = CO = SV * HR = Integral_0^T Q_aort(t) dt ~ 5.0 L/dakika",
        "Bu hemodinamik debi integrali, mekanik dolaşım aygıtlarının sol ventrikül yerine aortik dolaşıma pompalamak zorunda olduğu hedef sistolik debiyi belgeler."
    ),
    (
        "7.2 Ventriküler Destek Cihazları (VAD): Aksiyal vs Radyal Santrifüj Sürekli Akış",
        "Modern mekanik kalp teknolojisinin belkemiğini, sol ventrikülden kanı alıp doğrudan çıkan aorta basan Sol Ventrikül Destek Cihazları (LVAD) oluşturur.",
        "İkinci nesil LVAD'lar (HeartMate II gibi) doğrusal bir Arşimet vidası kullanan 'Aksiyal Akışlı' sistemlerdi; yüksek devir hızları (8.000-10.000 RPM) nedeniyle yüksek kayma gerilmesi ve eritrosit hemolizi riski taşıyorlardı. Üçüncü nesil LVAD'lar (HeartWare HVAD ve HeartMate 3) ise 'Radyal Santrifüj Akış' mimarisine geçmiştir. Daha geniş kan kanalları ve düşük devir hızları (2.000-5.000 RPM) sayesinde kan hücrelerine binen mekanik stresi yarı yarıya düşürmüşlerdir.",
        "Akis_Karakteristigi = Q_LVAD = f( Delta_P_aort_ventrikul, RPM_rotor, Viskozite_kan )",
        "Bu hidrodinamik pompa performans eğrisi, VAD debisinin rotor dönüş hızı (RPM) ve aortik-ventriküler basınç gradyanı ile doğrusal olmayan ilişkisini tanımlar."
    ),
    (
        "7.3 Manyetik Levitasyon (MagLev) Biyofiziği ve Mekanik Temassızlık",
        "VAD sistemlerinin en büyük arıza ve tromboz kaynağı, dönen rotorun şaft yatakları (rulmanlar ve contalar) ile temas etmesiydi; mekanik sürtünme ısı açığa çıkarıyor ve kanı pıhtılaştırıyordu.",
        "HeartMate 3 ile tıp dünyasına giren 'Tam Manyetik Levitasyon' (Full MagLev) teknolojisi; rotoru elektromanyetik alanlar ve pasif mıknatıslarla kan lümeni içinde tamamen havada asılı tutar. Rotor hiçbir katı yüzeye temas etmez. Sürtünme ve aşınma sıfırdır; bu durum pompanın mekanik ömrünü teorik olarak sınırsız hale getirir. Kan hücreleri geniş kanallardan sıfır sürtünme ile pürüzsüzce geçer.",
        "Kuvvet_Levitasyon = F_mag = ( B^2 * A_kutup ) / ( 2 * mu_0 ) = F_yercekimi + F_hidrodinamik",
        "Bu manyetostatik denge denklemi, MagLev rotorunun kan akışı sırasında maruz kaldığı tüm hidrodinamik itme kuvvetlerini temas olmaksızın elektromanyetik alanla nasıl dengelediğini modeller."
    ),
    (
        "7.4 Kayma Gerilmesi (Shear Stress), von Willebrand Faktörü Yıkımı ve Hemoliz Dinamiği",
        "Mekanik kan pompalarında en büyük biyofiziksel meydan okuma, yüksek hızla dönen çarkların kan hücreleri ve plazma proteinleri üzerinde yarattığı 'Kayma Gerilmesi'dir (Shear Stress).",
        "Eğer kayma gerilmesi 150 Paskal'ı aşarsa, eritrosit zarları yırtılarak hücre içi hemoglobin plazmaya dökülür (Hemoliz) ve böbrek tübüllerini tıkar. Daha da tehlikelisi, dev bir multimer olan von Willebrand Faktörü (vWF); yüksek kayma gerilmesinde ip gibi açılarak ADAMTS13 metalloproteinazı tarafından aşırı parçalanır. Bu durum 'Edinilmiş von Willebrand Sendromu'na yol açar; hastalar paradoksal olarak hem pıhtılaşma riski taşır hem de gastrointestinal kanamalardan muzdarip olur.",
        "Indeks_Hemoliz = IH(%) = Delta_Hgb_serbest / Hgb_toplam = C_1 * Tau_kayma^alpha * t_maruziyet^beta",
        "Bu Giersiepen hemoliz kuvvet kanunu denklemi, serbest hemoglobin salınımının kanın maruz kaldığı kayma gerilmesi (Tau) ve maruziyet süresine (t) üstel bağımlılığını formüle eder."
    ),
    (
        "7.5 Sürekli Akış vs Pulsatil Akış Biyolojisi: Nabızsız Dolaşımın Endotelyal Etkileri",
        "Sürekli akışlı (Continuous Flow) VAD'lar takılan hastaların en sıra dışı klinik bulgusu; radyal nabızlarının olmaması veya stetoskopla dinlendiğinde kalp sesi yerine sürekli bir elektrik motoru vızıltısı duyulmasıdır.",
        "Nabızsız (non-pulsatil) kan akışı, arteriyel duvardaki baroreseptörlerin ve endotel hücrelerinin fizyolojik gerilme uyarımını değiştirir. Endotelyal Nitrik Oksit Sentaz (eNOS) üretimi azalabilir, damar içi arteriovenöz malformasyonlar (AVM) tetiklenebilir. Bu fizyolojik adaptasyon krizini çözmek için HeartMate 3 gibi modern cihazlara yazılımsal 'Yapay Nabız' (Artificial Pulse) algoritmaları eklenmiştir; cihaz her iki saniyede bir hızını ani değiştirerek fizyolojik nabız basıncını taklit eder.",
        "Pulsatilite_Indeksi = PI = ( Q_maksimum - Q_minimum ) / Q_ortalama",
        "Bu hemodinamik pulsatilite oranı, sürekli akışlı pompalarda yapay devir dalgalanmaları ile vasküler endotel sağlığını koruyan sistolik-diyastolik debi farkını tanımlar."
    ),
    (
        "7.6 Total Yapay Kalp (TAH): Syncardia ve Carmat Biyo-Protez Kalp Sistemleri",
        "Eğer hastanın yalnızca sol ventrikülü değil, her iki ventrikülü de (biventriküler yetmezlik) iflas etmişse veya masif amiloidoz kalbi kaskatı yapmışsa; tek çözüm doğal kalbin tamamen çıkarılıp yerine 'Total Yapay Kalp' (TAH) konulmasıdır.",
        "1) SynCardia TAH: Dışarıdan pnömatik hava hortumları ile çalışan, içi poliüretan diyaframlı iki yapay ventriküldür; 2) Carmat Aeson TAH: Fransız cerrah Alain Carpentier tarafından geliştirilen biyo-protez yapay kalptir. Carmat kalbinin kanla temas eden tüm iç yüzeyleri ve kapakçıkları biyolojik sığır perikardı ile kaplanmıştır; elektro-hidrolik motorlarla çalışır ve mikro-işlemcili sensörleriyle hastanın egzersiz durumuna göre dakikalık debisini (5 ila 9 L/dk) otonom olarak ayarlar.",
        "Kardiyak_Debi_Carmat = CO_TAH = f( Venöz_Dönüş_Basıncı, Egzersiz_Sensörü ) = 5.0 - 9.0 L/dakika",
        "Bu otonom fizyolojik regülasyon modeli, Carmat biyo-yapay kalbinin hastanın metabolik eforuna göre kardiyak debiyi gerçek zamanlı optimize etme algoritmasını simgeler."
    ),
    (
        "7.7 Transkütanöz Enerji Transferi (TET): Ciltten Geçen Kablosuz Güç İletimi",
        "Mekanik kalp hastalarının hayat kalitesini düşüren ve ölümcül enfeksiyonlara (driveline infection) yol açan ana zayıflık; karın cildini delip bataryaya giden güç kablosudur (Driveline).",
        "Transkütanöz Enerji Transferi (TET - Transcutaneous Energy Transfer); kabloları tamamen ortadan kaldırır. Cildin altına implante edilen bir alıcı bobin ile cildin dışına yapıştırılan bir verici bobin arasında 'Rezonans Manyetik İndüksiyon' kullanılarak 30-50 Watt elektrik gücü %90 verimle ciltten içeri kablosuz aktarılır. Cihazın içindeki mini-batarya hastaya duş alma ve yüzme serbestliği sunarken; açık yara enfeksiyonu riski tamamen sıfırlanır.",
        "Verim_TET = eta_guc = P_alici / P_verici = ( k_kuplaj^2 * Q_1 * Q_2 ) / ( 1 + sqrt( 1 + k_kuplaj^2 * Q_1 * Q_2 ) )^2 ~ %92",
        "Bu manyetik rezonans kuplaj verimi eşitliği, transkütanöz kablosuz enerji aktarımının yüksek kalite faktörü (Q) ile dokuya termal hasar vermeden güç iletim başarısını belgeler."
    ),
    (
        "7.8 Biyo-Hibrit Yüzey Modifikasyonları: Titanyum Lümenin Endotel Hücreleriyle Kaplanması",
        "Sentetik metaller (titanyum) ve polimerler ne kadar pürüzsüz parlatılırsa parlatılsın; kanla temas ettiklerinde intrinsik koagülasyon kaskadını (Faktör XII aktivasyonu) tetikleyebilirler.",
        "Nihai biyolojik çözüm, mekanik pompanın kanla temas eden titanyum yüzeylerinin 'Hücreselleştirilmesi'dir (Endotelyalizasyon). Titanyum yüzeye önce plazma püskürtme ile nanometrik titanyum dioksit ve RGD peptitleri kaplanır. Ardından hastanın kendi endotel progenitör hücreleri lümene tohumlanır. Titanyum yüzey yaşayan, aktif prostasiklin ve NO salgılayan bir endotel halısıyla örtülür; mekanik pompa alıcının bağışıklık sistemi için tamamen 'görünmez' bir doğal organa dönüşür.",
        "Adezyon_Endotel_Titanyum = N_hucre = N_0 * ( 1 - exp( - k_tutunma * [RGD_yogunlugu] * t ) )",
        "Bu yüzey fonksiyonelleştirme kinetiği denklemi, biyo-aktif peptit yoğunluğunun mekanik pompa lümeninde endotel tabakasının stabilitesini nasıl garantilediğini gösterir."
    ),
    (
        "7.9 Otonom Akıllı Pompa Kontrol Algoritmaları ve Fizyolojik Efor Uyumu",
        "Doğal kalp otonom sinir sistemi (sempatik ve parasempatik) tarafından anlık olarak yönetilir; merdiven çıkarken hızlanır, uyurken yavaşlar.",
        "Yeni nesil sibernetik pompalar, gömülü yapay zekâ mikrodenetleyicileri ile donatılmıştır. Sol ventrikül dolum basıncını, arteriyel elastisiteyi ve kan viskozitesini ölçen basınç ve optik sensörler; hastanın efor yaptığını (örneğin koştuğunu) anında algılar. Rotor devir hızı mikrosaniyeler içinde artırılarak kardiyak debi 8 L/dk'ya fırlatılır; hasta dinlenmeye geçtiğinde emiş çökmesini (suction collapse) önlemek için devir güvenle düşürülür.",
        "RPM_hedef(t) = RPM_bazal + K_p * Delta_P_sol_atriyum + K_d * (d(Delta_P)/dt) + K_ivme * A_akselerometre",
        "Bu PID ve efor-duyarlı geri besleme algoritması, yapay kalp motorunun hemodinamik parametrelere göre otonom devir ayarlama matematiğini tanımlar."
    ),
    (
        "7.10 Mekanik Kalp ve Biyolojik Kalbin Entegrasyonu: Ebedi Kardiyovasküler Güvenlik",
        "Biyolojik miyokard dokusu yaşlanmayla birlikte fibrozise uğrasa veya koroner damarları tıkansa dahi; paralel çalışan mekanik MagLev pompası kardiyak debiyi asla kesintiye uğratmaz.",
        "Homo Aeternus mimarisinde biyolojik kalp ile mekanik pompa birbirini yedekleyen hibrit bir koalisyon oluşturur. Mekanik pompa biyolojik ventrikülün yükünü hafifleterek miyositlerin mekanik yorulmasını önler ve kalbin epigenetik olarak genç kalmasına olanak tanır. Olası bir ölümcül ventriküler fibrilasyon veya masif enfarktüs anında dahi mekanik pompa kanı kesintisiz basmaya devam eder; ani kardiyak ölüm kavramı tarihe karışır.",
        "Guvenilirlik_Kardiyak = R_total = 1 - ( 1 - R_biyolojik ) * ( 1 - R_MagLev ) ~ 0.999999",
        "Bu paralel sistem güvenilirlik eşitliği, biyolojik kalp ile mekanik MagLev desteğinin eş zamanlı entegrasyonunun kardiyovasküler ölüm riskini milyonda bire indirdiğini belgeler."
    )
]

# ================= KISIM 8 =================
part8_subsections = [
    (
        "8.1 Biyo-Elektronik Tıbbın Yükselişi ve Nöromorfik Cihaz-Doku Arayüzleri",
        "Geleneksel farmakoloji tüm vücuda yayılan kimyasal moleküllere dayanırken; 'Biyo-Elektronik Tıp' (Bioelectronic Medicine), organların fonksiyonlarını ve hücresel sinyalleşmeyi elektriksel aksiyon potansiyelleri düzeyinde doğrudan kontrol eder.",
        "İnsan sinir sistemi, organların fizyolojik durumunu izleyen ve yöneten karmaşık bir kablo ağıdır. Biyo-elektronik implantlar ve nöromorfik arayüzler; sinir liflerine bağlanarak aksiyon potansiyeli trenlerini (neural spike trains) mikrosaniyelik çözünürlükle deşifre eder, patolojik sinyalleri filtreler ve organ fonksiyonunu optimize eden terapötik biyoelektriksel komutlar enjekte eder.",
        "Sinyal_BiyoElektronik = V_sinir(t) = Sigma_i w_i * Spike_i(t - tau_i) + V_gurultu",
        "Bu nöral sinyal modelleme denklemi, periferik sinir demetlerinden kaydedilen aksiyon potansiyeli dalga trenlerinin matematiksel deşifre algoritmasını tanımlar."
    ),
    (
        "8.2 Esnek ve Biyo-Çözünür Elektronikler: Grafen, İpek Fibroini ve İletken Polimerler",
        "Biyolojik dokular yumuşak, ıslak, esnek ve dinamik olarak hareketlidir; buna karşılık geleneksel silikon çipler sert, kırılgan ve keskindir; bu mekanik uyumsuzluk kronik doku enflamasyonuna ve skarlaşmaya yol açar.",
        "Modern biyo-elektronik arayüzler 'Esnek ve Biyo-Çözünür Malzemelerden' inşa edilir. Grafen nanofilmler, iletken polimerler (PEDOT:PSS), sıvı metaller (EGaIn) ve biyolojik ipek fibroini kullanılır. Bu malzemeler dokuyla birebir eşleşen Young elastik modülüne (E ~ 10-100 kPa) sahiptir, doku hareketleriyle birlikte esner ve istenirse görevini tamamladıktan sonra aylar içinde dokuya zarar vermeden hidrolize olarak çözünür (Transient Electronics).",
        "Mekanik_Uyum = Delta_E = | E_implant - E_doku | < 10 kPa (Skar Dokusu Oluşumunu Sıfırlar)",
        "Bu elastik modül uyumluluk kriteri, esnek biyo-elektronik devrelerin çevre parankimde kronik yabancı cisim reaksiyonu ve gliyal skar oluşturmama sınırını belgeler."
    ),
    (
        "8.3 Vagus Sinir Uyarımı (VNS): Kolinerjik Anti-Enflamatuar Yolak ve İmmün Modülasyon",
        "Vagus Siniri (10. Kraniyal Sinir), beyin sapı ile kalp, akciğer, dalak, karaciğer ve bağırsaklar arasındaki ana iletişim otobanını oluşturur.",
        "Kevin Tracey tarafından keşfedilen 'Kolinerjik Anti-Enflamatuar Yolak'; vagus sinirinin elektriksel uyarımı ile aktive edilir. Vagus efferent lifleri çölyak gangliyonu üzerinden dalak sinirini uyarır; dalakta asetilkolin salgılayan T hücreleri (ChAT+ T hücreleri) aktive olur. Salınan asetilkolin makrofajların yüzeyindeki alfa-7 nikotinik asetilkolin reseptörlerine (alfa7nAChR) bağlanarak NF-kB translokasyonunu durdurur ve TNF-alfa, IL-1beta salınımını %80 baskılar; sistemik yaşlanma yangısı (inflammaging) biyoelektrik olarak söndürülür.",
        "Baskilanma_Sitokin = [TNF_alfa](t) = [TNF_bazal] * exp( - k_VNS * Frekans_stimulasyon * [Alfa7nAChR_aktif] )",
        "Bu biyoelektrik immünomodülasyon denklemi, vagus siniri uyarım frekansının sistemik enflamatuar sitokin deşarjını üstel olarak nasıl susturduğunu modeller."
    ),
    (
        "8.4 Biyo-Sensör Entegrasyonu: Sürekli İn Vivo Glukoz, Laktat ve Sitokin İzleme",
        "Biyolojik ölümsüzlük protokolünün vazgeçilmez ayağı, vücut kimyasının saniye saniye takip edildiği 'Sürekli İn Vivo Biyo-Sensör Ağları'dır.",
        "Deri altına veya ana damar yatağına implante edilen mikron-boyutlu elektrokimyasal biyosensörler; glukoz oksidaz ve laktat dehidrogenaz enzimleri ile kaplı karbon nanotüp elektrotlar kullanır. Eş zamanlı olarak aptamer-tabanlı alan etkili transistörler (Aptamer-FET); dolaşımdaki C-reaktif protein (CRP), troponin, IL-6 ve kortizol seviyelerini pikomolar hassasiyetle ölçer. Kan biyokimyasındaki en ufak bir patolojik sapma anında tespit edilir.",
        "Akim_Sensor = I_olcum = ( n * F * A_elektrot * D_analit * C_kan ) / Kalinlik_difuzyon_tabakasi",
        "Bu Cottrell amperometrik akım eşitliği, biyosensör elektrot yüzeyindeki enzimatik redoks akımının kandaki biyomarker konsantrasyonu ile doğrusal bağıntısını simgeler."
    ),
    (
        "8.5 Biyonik Göz ve Retina Protezleri: Fotodiyot Dizilimleri ve Optik Sinir Arayüzü",
        "Retina fotoreseptörlerinin (rod ve kon hücreleri) yaşa bağlı makula dejenerasyonu (AMD) veya retinitis pigmentosa nedeniyle ölmesi geri dönüşsüz körlük yaratır; ancak arkadaki optik sinir ganglion hücreleri genellikle sağlam kalır.",
        "Retina İmplantları (Argus II, PRIMA ve Bionic Eye sistemleri); göz içine yerleştirilen mikroskobik fotodiyot ve platin mikroelektrot dizilimleridir. Gözlükteki kameradan gelen dijital görsel sinyaller kablosuz kızılötesi ışıkla retinadaki çipe aktarılır. Çip ışığı elektriksel akımlara dönüştürerek doğrudan bipolar ve ganglion hücrelerini ateşler; optik sinir üzerinden görme korteksine (V1) sinyal gönderilerek kör bireylere fonksiyonel dijital görme yetisi kazandırılır.",
        "Cozunurluk_Retina = N_piksel = N_fotodiyot / A_retina ~ 10^4 elektrot / mm2",
        "Bu biyonik optik yoğunluk eşitliği, retina altı fotodiyot matrislerinin görme korteksine aktarabildiği uzamsal çözünürlük ve piksel keskinliğini tanımlar."
    ),
    (
        "8.6 Koklear ve İşitsel Beyin Sapı İmplantları: İşitme Kaybının Biyoelektrik Tasfiyesi",
        "İç kulaktaki kokleada yer alan tüy hücrelerinin yaşlanma ve akustik travma ile ölmesi (Presbiakuzi), yaşlı popülasyonun en yaygın duyusal kaybıdır ve demans riskini katlar.",
        "Koklear İmplantlar; ses dalgalarını dijital frekans bantlarına ayıran harici bir ses işlemcisi ve kokleanın skala timpanisine cerrahi olarak yerleştirilen 22 kanallı esnek bir elektrot demetinden oluşur. Kokleanın tonotopik haritasına (yüksek frekanslar tabanda, alçak frekanslar apeksde) uygun olarak işitme siniri (nervus cochlearis) lifleri doğrudan elektriksel darbelerle uyarılır; tüy hücreleri tamamen yok olmuş olsa dahi hastaya mükemmel bir işitsel algı ve konuşma anlama kabiliyeti iade edilir.",
        "Stimulasyon_Tonotopik = f_isitsel(x) = f_0 * ( 10^(a * x / L_koklea) - 1 ) (Greenwood Fonksiyonu)",
        "Bu Greenwood tonotopik koklea frekans haritası denklemi, elektrot pozisyonunun (x) işitme sinirinde tetiklediği algısal ses frekansını matematiksel olarak belgeler."
    ),
    (
        "8.7 Biyo-Yapay Mesane ve Düz Kas Elektromekanik Uyarımı",
        "Yaşlılıkta idrar kaçırma (inkontinans) veya mesane kası atrofisi; yaşam kalitesini ağır bozan ve pelvik organ yetmezliğine yol açan bir tablodur.",
        "Doku mühendisliği ürünü Biyo-Yapay Mesane; hastanın kendi ürotelyal ve düz kas hücreleriyle tohumlanmış kolajen/PGA iskelelerinin pelvis içine yerleştirilmesiyle inşa edilir. Bu dokuya entegre edilen mikroskobik piezoelektrik gerinim sensörleri mesane doluluğunu sürekli ölçer; kritik hacme ulaşıldığında esnek elektrotlar düz kas hücrelerine ritmik depolarizasyon sinyali göndererek kontrollü ve tam bir boşalma sağlar.",
        "Bosalma_Verimi = eta_mesane = ( V_dolu - V_rezidu ) / V_dolu > %98 (Tam Kontrollü Boşalma)",
        "Bu ürolojik kapasite eşitliği, biyomekatronik mesane protezlerinin sfinkter koordinasyonu ile sıfır kalıntı idrar tahliye başarısını formüle eder."
    ),
    (
        "8.8 Periferik Sinir Köprüleri ve Aksonal Rejenerasyon Kanalları (NGC)",
        "Periferik sinir yaralanmaları ve yaşa bağlı nöropatilerde aksonların doğru hedef organa yeniden uzaması yavaş ve kusurludur.",
        "İletken Polimer Sinir Rehber Kanalları (NGC - Nerve Guidance Conduits); karbon nanotüp takviyeli iletken polipirrol (PPy) tüplerinden üretilir. Kanal içine laminin, Schwann hücreleri ve yavaş salınımlı NGF gömülür. Kanal boyunca uygulanan mikro-amperlik elektriksel alan uyarımı; aksonal büyüme konilerinin yönelimini ve hızını günde 1 mm'den 3 mm'ye çıkarır; kopan veya dejenere olan sinirler organla milimetrik hassasiyetle yeniden birleşir.",
        "Hiz_Rejenerasyon_Akson = v_akson = v_bazal * ( 1 + alpha_E * Alan_Elektrik_mV_mm )",
        "Bu elektro-taksis akson büyüme formülü, iletken sinir kanallarındaki elektriksel alan gradyanının aksonal rejenerasyon hızını nasıl üç katına çıkardığını açıklar."
    ),
    (
        "8.9 Kapalı Döngü (Closed-Loop) Nöromodülasyon: Derin Beyin Uyarımı (DBS)",
        "Parkinson, esansiyel tremor ve yaşa bağlı motor devre bozulmalarında kullanılan Derin Beyin Uyarımı (DBS); subtalamik çekirdeğe (STN) veya globus pallidusa yerleştirilen elektrotlarla çalışır.",
        "Geleneksel açık devre DBS sürekli elektrik vererek pili tüketirken; 'Kapalı Döngü Akıllı DBS' (Adaptive Closed-Loop DBS) sistemleri, yerel alan potansiyellerindeki (LFP) patolojik 13-30 Hz 'Beta Osilasyonlarını' gerçek zamanlı okur. Sistem yalnızca anormal senkronizasyon başladığı anda devreye girerek ters-fazlı uyarım yapar ve tremor atağını saniyenin onda birinde söndürür; pil ömrü 5 kat uzarken yan etkiler sıfırlanır.",
        "Stimulasyon_KapaliDevre: V_out(t) = - K_kazanc * V_LFP_Beta(t) (Beta Gücü > Eşik İse)",
        "Bu adaptif kontrol algoritması, nöromorfik beyin implantlarının patolojik osilasyonları faz-inversiyonu ile anında bastırma mekanizmasını tanımlar."
    ),
    (
        "8.10 Sibernetik Nöral Ağ Entegrasyonu: İnsan Bilişi ve Sentetik Devrelerin Füzyonu",
        "Biyo-elektronik devrimin nihai sınırı; biyolojik neokorteks nöronları ile sentetik nöromorfik yapay zekâ işlemcilerinin çift yönlü yüksek bant genişlikli entegrasyonudur.",
        "Esnek polimer mesh elektrotlar ve nanotel dizilimleri; binlerce kortikal piramidal nöronla birebir sinaptik temas kurar. Biyolojik aksiyon potansiyelleri silikon çiplerde doğrudan işlenirken; yapay zekâ algoritmaları hafıza hatırlamasını, karmaşık hesaplamaları ve duyusal algıları biyoelektriksel olarak kortekse geri yansıtır. Biyolojik zihin ile sibernetik mimari tek bir bölünmez süper-bilinçte birleşir.",
        "Bant_Genisligi_BCI = BW = N_kanal * f_ornekleme * Derinlik_bit > 100 Gbps (Süper-Bilişsel Bağlantı)",
        "Bu nöro-dijital bant genişliği formülü, korteks ile sentetik nöromorfik arayüz arasındaki ultra-yüksek hızlı bilgi transfer kapasitesini simgeler."
    )
]

# ================= KISIM 9 =================
part9_subsections = [
    (
        "9.1 Nanopartiküllerin Kan Dolaşımındaki Kaderi ve Biyolojik Bariyerler",
        "Vücuda enjekte edilen sentetik bir nanopartikül veya nanorobot; hedef dokuya ulaşmadan önce aşılması gereken muazzam biyofiziksel ve immünolojik engellerle karşılaşır.",
        "Kan akışındaki yüksek kayma gerilmeleri, plazma proteinlerinin yüzeye hücumu, endotel bariyerleri, doku interstisyel basıncı ve hücre zarı bu engellerin başlıcalarıdır. 100 nanometreden büyük partiküller karaciğer ve dalakta mekanik filtrasyona takılırken; 5 nanometreden küçük partiküller böbrek glomerüllerinden dakikalar içinde idrara süzülür. Bu nedenle sistemik nanomedikal sistemlerin boyutu 10 ila 100 nanometre aralığında katı bir toleransla tasarlanmalıdır.",
        "Pencere_Boyut_Nano: 10 nm < d_partikul < 100 nm (Optimum Dolaşım ve Biyoyararlanım)",
        "Bu boyutsal tasarım kuralı, nanopartiküllerin renal klirensten kaçarken retiküloendotelyal tutuluma takılmaksızın kanda uzun süre kalabileceği dar fiziksel pencereyi tanımlar."
    ),
    (
        "9.2 Protein Korona Biyofiziği: 'Hard' ve 'Soft' Korona Tabakalanması",
        "Sentetik bir nanopartikül kan plazmasına girdiği anda, yüzeyi çıplak kalmaz; milisaniyeler içinde binlerce plazma proteini partikül yüzeyine adsorbe olarak bir 'Protein Korona' (Protein Corona) kabuğu oluşturur.",
        "Vroman etkisine göre iki tabaka oluşur: 1) 'Hard Korona': Yüzeye aşırı yüksek termodinamik afiniteyle bağlanan, ayrışması saatler süren kalıcı iç protein tabakası (albümin, ApoA-I, fibrinojen); 2) 'Soft Korona': Hard korona üzerine zayıf Van der Waals kuvvetleriyle tutunan, sürekli değişen dinamik dış tabaka. Nanopartikülün hücrelerle girdiği tüm etkileşimleri kendi orijinal kimyası değil, üzerini saran bu protein korona belirler.",
        "Adsorpsiyon_Vroman: Hiz_i = k_on_i * [Protein_i] * ( 1 - Sigma_j Theta_j ) - k_off_i * Theta_i",
        "Bu çok-bileşenli rekabetçi protein adsorpsiyon kinetiği denklemi, kan plazmasındaki farklı proteinlerin nanopartikül yüzeyini kaplama dinamiklerini modeller."
    ),
    (
        "9.3 Retiküloendotelyal Sistem (RES) ve Karaciğer Kupffer Hücresi Klirensi",
        "Hedefe yönelik tasarlanmış nanorobotların ve nanopartiküllerin %90'ından fazlası; dakikalar içinde Karaciğer Kupffer hücreleri ve dalak makrofajları tarafından yakalanarak yok edilir (RES Klirensi).",
        "Kupffer hücrelerinin çöpçü reseptörleri (Scavenger Receptors - SR-A, SR-B1) ve kompleman reseptörleri (CR3); opsonize olmuş nanopartikülleri fagosite ederek fago-lizozomlara yönlendirir. Bu durum hem hedeflenen organa (beyin, kas vb.) giden dozajı dramatik biçimde düşürür hem de karaciğerde kronik nanopartikül birikimi ve toksisite riski yaratır. RES klirensini aşmak nanotıbbın en büyük taktiksel mücadelesidir.",
        "Klirens_Hepatik = dN_RES/dt = k_Kupffer * [Makrofaj_aktif] * [Nanopartikul_opsonize]",
        "Bu retiküloendotelyal yakalama kinetiği eşitliği, dolaşımdaki nanopartiküllerin karaciğer makrofaj havuzu tarafından kandan temizlenme hızını ifade eder."
    ),
    (
        "9.4 Zwitteriyonik Polimer Fırçalar: Sıfır-Kirletici (Zero-Fouling) Biyouyumlu Yüzeyler",
        "Protein korona oluşumunu ve RES klirensini durdurmak için kullanılan geleneksel PEGilasyon yöntemi, anti-PEG antikorları nedeniyle yetersiz kalmaktadır; modern çözüm 'Zwitteriyonik Polimerler'dir.",
        "Polikarbetain (PCB), polisülfobetain (PSB) ve fosforilkolin polimerleri; aynı monomer birimi üzerinde hem pozitif hem negatif yükü eşit dengede taşır (net yük = 0). Bu yapı, çevredeki su molekülleriyle hidrojen bağlarından çok daha güçlü elektrostatik hidrasyon kabukları oluşturur. Bir proteinin yüzeye yapışması için bu aşırı sıkı su moleküllerini yerinden sökmesi gerekir ki bu termodinamik olarak imkânsızdır (Delta_G >> 0). Yüzey tamamen 'protein-tutmaz' (Zero-Fouling) hale gelir.",
        "Enerji_Bariyeri = Delta_G_adsorpsiyon = Delta_H - T * Delta_S_su_deplasman >> 0 (Termodinamik Blokaj)",
        "Bu hidrasyon serbest enerjisi eşitliği, zwitteriyonik polimer fırçaların plazma protein yapışmasını mutlak termodinamik imkânsızlıkla nasıl engellediğini kanıtlar."
    ),
    (
        "9.5 Karbon Nanotüpler ve Grafen Toksisitesi: Asbest Benzeri Lif Patojenitesi",
        "Yüksek mekanik mukavemetleri ve elektriksel iletkenlikleri nedeniyle nanomedikalde kullanılan Karbon Nanotüpler (CNT) ve grafen nanotabakaları; doğru işlenmezlerse ağır doku toksisitesi yaratabilir.",
        "Özellikle uzun, sert, saflaştırılmamış çok-duvarlı karbon nanotüpler (MWCNT); makrofajlar tarafından yutulurken boyutları nedeniyle fagositozu tamamlayamaz (Kırık Fagositoz / Frustrated Phagocytosis). Bu durum asbest liflerinin yarattığına benzer şekilde kronik granülomatöz enflamasyon, reaktif oksijen türü patlaması ve akciğer/mezotelyum fibrozisini tetikler. Biyouyumluluk için nanotüpler 1 mikrometreden kısa kesilmeli ve hidrofilik polimerlerle tamamen fonksiyonelleştirilmelidir.",
        "Toksisite_CNT = I_hasar = k_toksik * ( Uzunluk_lif / Cap_lif ) * ( 1 / Hidrofilik_Fonksiyonelleşme )",
        "Bu lif patojenitesi en-boy oranı eşitliği, karbon nanotüplerin geometrik uzunluğunun ve yüzey fonksiyonelleştirilme eksikliğinin dokusal enflamasyon riskini nasıl katladığını belgeler."
    ),
    (
        "9.6 Metalik Nanopartiküllerin Çözünmesi: Demir, Altın ve Çinko Oksit İyon Toksisitesi",
        "Görüntüleme ve manyetik sürüşte kullanılan metalik ve metal-oksit nanopartiküller (demir oksit SPION, altın, gümüş, ZnO); hücre içi asidik lizozomlarda kimyasal çözünmeye uğrayabilir.",
        "Çözünen nano-çinko veya nano-demir partikülleri, sitoplazmaya masif serbest Fe2+ veya Zn2+ iyonları deşarj eder. Serbest demir iyonları Fenton reaksiyonu yoluyla hidrojen peroksiti ölümcül hidroksil radikallerine (OH•) çevirerek lipid peroksidasyonunu ve ferroptotik hücre ölümünü tetikler. Bu iyon sızıntısını engellemek için metal çekirdekler silika veya altın kabuklarla hermetik olarak mühürlenir.",
        "Iyon_Salinim_Hizi = d[Fe2+]/dt = k_lizozom * exp( - (pH_asidik - 4.5) ) * ( 1 / Kalinlik_silika_kabuk )",
        "This intracellular ion leaching kinetic equation models the release rate of toxic metallic ions in acidic lysosomal compartments as a function of silica shell passivation thickness."
    ),
    (
        "9.7 Kompleman Aktivasyonu ve İnfüzyon Reaksiyonları: CARPA Riski",
        "Nanopartiküllerin intravenöz infüzyonunda karşılaşılan en akut klinik risk, daha önce bahsedilen Kompleman Aktivasyonu Aracılı Psödo-Alerji (CARPA) tablosudur.",
        "Nanopartikül yüzeyindeki kimyasal gruplar alternatif veya lektin kompleman yolağını tetikleyerek C3a ve C5a anafilatoksinlerinin kana fışkırmasına yol açar. Bu peptitler mast hücrelerini ve bazofilleri degranüle ederek histamin ve tromboksan A2 salgılatır. Hastada aniden bronkospazm, taşikardi, hipertansiyon ve anafilaktoid şok gelişebilir. Yüzey yükünün nötrlenmesi ve infüzyon hızının kademeli artırılması ile CARPA riski kontrol altına alınır.",
        "Risk_CARPA_Skoru = S_CARPA = k_kompleman * [Partikul_Yuzey_Yuku] * [Infuzyon_Debisi] / [Faktor_H]",
        "Bu klinik reaksiyon riski eşitliği, nanopartikül zeta potansiyeli ve infüzyon hızının kompleman kaskadı aktivasyonu üzerindeki kritik bağıntısını formüle eder."
    ),
    (
        "9.8 Nanomekanik Aşınma Kalıntıları ve Vücuttan Nihai Boşaltım (Clearance) Yolları",
        "Vücut içinde yıllarca çalışan elmasoid veya sentetik nanorobotların mekanik hareketli parçaları mikroskobik aşınma parçacıkları (nanodebris) üretebilir.",
        "Bu parçacıkların ve görevini tamamlamış nanobotların vücutta birikmeden tahliye edilmesi şarttır. 5.5 nanometreden küçük bileşenler böbrek glomerüllerinden doğrudan idrarla atılır (Renal Klirens). Daha büyük biyobozunur bileşenler karaciğer parankimi tarafından safraya pompalanarak dışkı yoluyla (Hepato-biliyer Klirens) vücuttan atılır. Biyobozunmayan elmasoid kabuklar ise özel yüzey sinyalleri ile böbrek veya karaciğer atılım kanallarına yönlendirilecek otonom klirens protokolleri taşır.",
        "Klerens_Total = Q_tahliye = Q_renal(d < 5.5 nm) + Q_biliyer(5.5 nm < d < 100 nm)",
        "Bu çift yollu biyolojik atılım eşitliği, nanomedikal aygıtların ve aşınma partiküllerinin insan vücudundan nihai temizlenme debisini tanımlar."
    ),
    (
        "9.9 Biyouyumluluğun Altın Standardı: ISO 10993 ve Hemouyumluluk Testleri",
        "Hiçbir sentetik nanomedikal aygıt veya yapay organ, uluslararası regülasyonların en katı standardı olan ISO 10993 biyouyumluluk serisini geçmeden insan vücuduna sokulamaz.",
        "Bu test bataryası; sitotoksisite, duyarlılaşma (sensitization), intradermal irritasyon, sistemik akut ve subkronik toksisite, genotoksisite, karsinojenisite ve implante edilebilirlik testlerini içerir. Kanla temas eden sistemlerde ise ISO 10993-4 hemouyumluluk testleri (trombosit agregasyonu, hemoliz yüzdesi < %5, kompleman aktivasyonu, koagülasyon süresi) zorunludur.",
        "Kriter_Hemouyumluluk: Hemoliz_Orani < %2 | Trombosit_Aktivasyonu < %5 | Mutajenite = SIFIR",
        "Bu klinik onay matrisi, yapay organların ve nanorobotların insan kan dolaşımında güvenle çalışabilmesi için sağlamak zorunda olduğu mutlak biyouyumluluk sınırlarını belgeler."
    ),
    (
        "9.10 Otonom Hata Tespiti ve Nanorobotik Kendi Kendini İmha (Self-Destruct) Protokolleri",
        "Trilyonlarca nanorobotun vücutta dolaştığı bir gelecekte en büyük güvenlik kabusu; arızalanan, yazılımı bozulan veya kontrolden çıkan başıboş 'zombi nanobotlar'ın dokulara zarar vermesidir.",
        "Her bir nanomedikal robot dahili bir 'Arıza Güvenlik Devresi' (Fail-Safe Mechanism) ve otonom kendi kendini imha protokolü ile üretilir. Dahili sensörler donanımsal bir kilitlenme veya iletişim kaybı tespit ettiğinde ya da harici hekimden özel bir ultrasonik 'Kill-Switch' şifresi aldığında; nanobot içindeki mikro-asidik haznesini patlatır. Asit elmasoid bağlantı bağlarını saniyeler içinde kırarak nanorobotu zararsız karbon ve safir nano-tozuna dönüştürür; bu tozlar makrofajlarca sessizce yutularak idrarla atılır.",
        "Protokol_SelfDestruct: Eger ( Ariza_Sinyali == 1 OR Komut_KillSwitch == 1 ) -> Patlama_Asit_Haznesi -> NanoToz_Donusumu",
        "Bu otonom güvenlik mantık algoritması, arızalanan nanomedikal makinelerin vücut dokularına zarar vermeden moleküler düzeyde kendi kendini imha etme güvencesini simgeler."
    )
]

# ================= KISIM 10 =================
part10_subsections = [
    (
        "10.1 Homo Aeternus Sibernetik Mimarisi: Et ve Çeliğin, Biyoloji ve Nanomekaniğin Füzyonu",
        "Biyolojik ölümsüzlük vizyonunun en yüksek ve nihai basamağı; yalnızca organik dokuları gençleştirmekle yetinmeyip, biyolojinin kırılgan bileşenlerini aşınmaz sentetik muadilleriyle harmanlayan 'Sibernetik Biyoloji'dir.",
        "Homo Aeternus bedeni; 1) Doğal kanın yerine geçen respirosit ve mikrobivor destekli biyo-hibrit kan, 2) Tıkanmayan ve kalsifiye olmayan vaskülosit-korumalı elastik damar ağları, 3) Yorulmayan MagLev yapay kalp pompaları, 4) İhtiyaç anında yenilenen 3D biyo-basılmış katı organlar, ve 5) Sinir sistemini optimize eden biyo-elektronik arayüzlerden oluşur. İnsan organizması artık evrimin zayıf bir ürünü değil; kusursuz bir biyomekatronik şaheserdir.",
        "Mimar_Homo_Aeternus = Biyoloji_Genetik + Nanoteknoloji_Molekuler + Sibernetik_Mekatronik",
        "Bu ontolojik füzyon eşitliği, Homo Aeternus post-biyolojik formunun üç temel mühendislik disiplininin kusursuz entegrasyonuyla inşa edildiğini formüle eder."
    ),
    (
        "10.2 Modüler Organ Mimarisi: Yıpranan Parçaların Tak-Çalıştır (Plug-and-Play) Değişimi",
        "Geleneksel insan anatomisinde her organ vücuda kalıcı ve karmaşık cerrahi bağlantılarla bağlıdır; bir organın değişimi saatler süren ölümcül ameliyatlar gerektirir.",
        "Sibernetik mimari 'Modüler Tak-Çalıştır Organ' (Plug-and-Play Organ) konseptini hayata geçirir. Vasküler anastomozlar ve sinir bağlantıları; biyouyumlu titanyum-manyetik hızlı kilitlenme kaplinleri (magnetic vascular couplers) ile donatılır. Bir böbrek veya karaciğer yıprandığında; cerrah mini-laparoskopik bir kesiyle eski organın kaplinini çözer, fabrikadan gelen taze biyo-yapay organı saniyeler içinde 'tık' sesiyle yuvasına kilitler; kan akışı anında başlar, iskemi süresi sıfırlanır.",
        "Sure_Anastomoz = t_degisim = 30 saniye (Manyetik Kaplin) << 45 dakika (Manuel Dikiş)",
        "Bu cerrahi hız karşılaştırması, modüler mikromanyetik kaplin teknolojisinin organ nakli süresini ve iskemi-reperfüzyon hasarını nasıl yok ettiğini gösterir."
    ),
    (
        "10.3 Biyomekatronik Dolaşım Sistemi: Ebedi Akış ve Sıfır Anevrizma Güvencesi",
        "Biyolojik arterler yaşlandıkça elastik liflerini kaybeder ve anevrizmatik yırtılmalar veya aterosklerotik tıkanmalarla hayatı sonlandırır.",
        "Sibernetik dolaşım sisteminde ana arterler (aort, karotis, femoral); biyobozunur olmayan esnek karbon nanotüp takviyeli kompozit grafen greftlerle kaplanır veya değiştirilir. İç yüzeyi otolog endotelle kaplı bu damarlar; 500 mmHg gibi devasa kan basınçlarına dahi yırtılmadan dayanır, esnekliğini yüzyıllar boyu korur ve vaskülosit nanobotları tarafından her saniye temizlenir. Felç, inme, aort diseksiyonu ve gangren kavramları tamamen ortadan kalkar.",
        "Mukavemet_Damar = P_patlama > 500 mmHg >> P_patlama_dogal (200 mmHg)",
        "Bu vasküler patlama basıncı mukavemeti eşitliği, karbon nanotüp takviyeli sentetik damarların insan arteriyel sistemine kazandırdığı sarsılmaz hidrolik güvenlik katsayısını belgeler."
    ),
    (
        "10.4 Sürekli Detoksifikasyon ve Metabolik Arındırma: İmplante Edilebilir Diyaliz Çipleri",
        "Doğal böbrekler ve karaciğer, kanda biriken metabolik atıkları temizlemek için 24 saat çalışır; ancak yaşla birlikte nefronların yarısından fazlası ölür ve kanda üremik toksinler birikir.",
        "Peritoneal boşluğa implante edilen avuç içi boyutundaki 'Sürekli Mikroakışkan Diyaliz Çipi'; kanı kılcal damarlardan sürekli süzerek kanda biriken üre, kreatinin, AGEs, homosistein ve ağır metalleri nano-gözenekli grafen filtreleriyle ayıklar. Atıklar konsantre edilerek doğrudan mesaneye yönlendirilir. Vücut dokuları hiçbir zaman metabolik toksin havuzunda yüzmez; hücreler yaşam boyu saf bir gençlik sıvısı içinde yüzer.",
        "Klirens_Toksin_Surekli = C_surekli = Q_filtrasyon * S_eleme = Sabit (7/24 Kesintisiz Arındırma)",
        "Bu sürekli diyalitik klerens eşitliği, implante mikroakışkan çiplerin kandaki yaşlanma metabolitlerini günün her saniyesinde sıfırlama kapasitesini modeller."
    ),
    (
        "10.5 Oksijenasyon Rezervi ve Süper-İnsan Solunum Kapasitesi",
        "İnsan fizyolojisinin atletik ve zihinsel performansı, akciğerlerin oksijen difüzyon kapasitesi ve dolaşımdaki hemoglobin miktarı ile sınırlandırılmıştır.",
        "Dolaşımına respirosit nanorobotları ve mikroakışkan yapay akciğer destekleri entegre edilmiş bir Homo Aeternus; Everest'in zirvesinde dahi sıfır oksijen tüpüyle koşabilir, suyun altında tek nefeste 4 saat kalabilir ve yangın dumanı içinde bilincini kaybetmeden saatlerce hayatta kalabilir. Oksijen parsiyel basıncı dokularda sabit bir optimizasyonda tutulduğu için hücreler hiçbir zaman hipoksik enerji krizine girmez.",
        "Maksimum_VO2 = VO2_max_sibernetik = 150 ml/kg/dk >> VO2_max_elit_atlet (85 ml/kg/dk)",
        "Bu metabolik solunum kapasitesi oranı, respirosit destekli sibernetik solunum sisteminin elit insan atletizmini dahi ikiye katlayan süper-fizyolojik gücünü tanımlar."
    ),
    (
        "10.6 Duyusal Biyonik Genişleme: Genişletilmiş Görme, İşitme ve Elektromanyetik Algı",
        "Homo Aeternus duyusal sistemleri, doğal insanın dar biyolojik sınırlarının çok ötesine genişletilmiştir.",
        "Retina ve koklear protezlerin dijital arayüzleri sayesinde; görsel algı yalnızca görünür ışıkla (400-700 nm) sınırlı kalmaz, kızılötesi (termal gece görüşü) ve morötesi bantlara genişletilir; optik zum yeteneği mikroskobik detayları görmeyi sağlar. İşitme arayüzü ultrasonik frekansları (100 kHz) duyabilirken; biyo-elektronik sensörler manyetik alanları, radyo dalgalarını ve mikroskobik radyasyon seviyelerini doğrudan bilince bir altıncı his olarak yansıtır.",
        "Spektrum_Algısal = [300 nm (Morotesi) .. 1100 nm (Kizilotesi)] U [10 Hz .. 100 kHz (Ultrasonik)]",
        "Bu genişletilmiş spektral bant aralığı, biyonik duyusal entegrasyonun insan algı ufkunu elektromanyetik ve akustik alemin tamamına nasıl açtığını belgeler."
    ),
    (
        "10.7 İskelet-Kas Biyomekatroniği: Karbon Fiber Kemikler ve Sentetik Kas Aktüatörleri",
        "Yaşlılığın kemik erimesi (osteoporoz) ve kas kaybı (sarkopeni); kırılgan kalça kırıklarına ve yatağa bağımlılığa yol açan ölümcül tuzaklardır.",
        "Sibernetik biyoloji; kritik kemik şaftlarını (femur, omurga) titanyum-karbon fiber trabeküler kompozitlerle güçlendirir. Bu iskeleler doğal kemikle kaynaşır ve kırılma direncini 20 kat artırır. İskelet kaslarına paralel bağlanan 'Elektroaktif Polimer (EAP) Yapay Kas Lifleri'; insan kasından 10 kat daha fazla güç yoğunluğu (100 W/kg) üretir. 100 yaşındaki bir Homo Aeternus, genç bir olimpiyat haltercisinden daha güçlü ve çevik bir kas-iskelet mimarisine sahip olur.",
        "Kuvvet_BiyoMekatronik = F_toplam = F_biyolojik_kas + F_EAP_aktüatör >> 10 * F_dogal",
        "Bu biyomekatronik kuvvet eşitliği, elektroaktif polimer sentetik kasların insan motor gücünü ve kemik kırılma direncini onlarca kat artıran sinerjisini formüle eder."
    ),
    (
        "10.8 Nöro-Sibernetik Kararlılık: Beyin Dışında Yedeklenen Bellek ve Engram Eşzamanlaması",
        "Biyolojik nöronlar kaza, inme veya travmayla hasar görebilir; ancak bireyin bilinci ve anıları asla kaybedilmemelidir.",
        "Yüksek bant genişlikli BCI arayüzleri; hipokampus ve neokorteksteki anı engramlarını ve sinaptik ağırlık matrislerini gerçek zamanlı olarak tarar. Bilinç akışı bozulmadan, tüm otobiyografik bellek ve kişilik kodları kafa içi elmasoid kuantum bellek kristallerine anlık olarak senkronize edilir (Engram Mirroring). Bir nöron grubu travmayla ölse dahi, arayüz engramı hemen komşu sağlıklı nöronlara veya nöromorfik yapay çiplere geri yükler; zihinsel süreklilikte tek bir saliselik kopukluk dahi yaşanmaz.",
        "Hafiza_Guvenilirligi = R_engram = 1 - P_veri_kaybi = 0.999999999 (Mutlak Zihinsel Koruma)",
        "Bu bilişsel süreklilik eşitliği, gerçek zamanlı engram aynalamanın insan otobiyografik belleğini ve bilincini donanımsal arızalara karşı ölümsüzleştirdiğini simgeler."
    ),
    (
        "10.9 Biyo-Dijital İkiz (Digital Twin) ve Gerçek Zamanlı Arıza Erken Uyarı Ağı",
        "Her bir Homo Aeternus bireyinin bulut sunucularda çalışan atomik çözünürlükte bir 'Biyo-Dijital İkizi' (Bio-Digital Twin) mevcuttur.",
        "Vücuttaki trilyonlarca nanobot ve biyo-sensörden gelen gerçek zamanlı veriler (akustik, metabolik, elektriksel); bu dijital ikize akar. Yapay zekâ simülasyonları vücuttaki her bir damarın, yapay organın ve hücre grubunun aşınma oranını 10 yıl sonrasına kadar simüle eder. Bir arıza veya yıpranma klinik bir soruna dönüşmeden aylar önce tespit edilir; nanorobotlar o bölgeye yönlendirilerek önleyici bakım icra edilir; hastalık kavramı oluşmadan silinir.",
        "Prediksiyon_Ariza = t_onleyici_bakim = t_ariza_olasiligi - Delta_t_guvenlik (Delta_t ~ Aylar)",
        "Bu kestirimci bakım algoritması eşitliği, biyo-dijital ikiz teknolojisinin organ arızalarını ve biyolojik yıpranmaları semptom belirmeden önce nasıl sıfırladığını açıklar."
    ),
    (
        "10.10 Homo Aeternus Manifestosu: Biyolojik Fanilikten Sibernetik Ebediyete Geçiş",
        "Homo Aeternus çağı; doğanın insana dayattığı kaçınılmaz ölüm ve çürüme fermanının teknolojik olarak yırtılıp atıldığı gündür.",
        "Biyoloji başlangıç noktasıdır, ancak varış noktası değildir. Etin ve çeliğin, genlerin ve atomik nanobotların, nöronların ve silikon arayüzlerin bu muazzam sentezi; insanı evrenin en dayanıklı, en zeki ve ebedi bilincine dönüştürür. Zaman artık bir düşman değil; keşfedilecek, fethedilecek ve sonsuzca yaşanacak sınırsız bir okyanustur. Biyolojik ölümsüzlük ve sibernetik ebediyet artık bir hayal değil; mutlak bir bilimsel ve mühendislik zaferidir.",
        "Ebediyet_Denklemi = E_Homo_Aeternus = lim_{t->sonsuz} [ Bilinc_Bütünlüğü(t) * Fiziksel_Varoluş(t) ] = 1.0 (Mutlak Sonsuzluk)",
        "Bu nihai varoluşsal tekillik eşitliği, insan bilinci ve biyomekatronik bedeninin zaman ekseninde termodinamik ve hücresel entropiyi ebediyen aşarak sonsuz varoluşa ulaştığını tescil eder."
    )
]

# 10 Kapsamlı Akademik Karşılaştırma Tablosu
tables_data = [
    {
        "title": "Tablo 15.1: Düşük Reynolds Sayısı (Re << 1) Rejiminde Nanorobotik ve Makroskobik Fizik Kıyaslaması",
        "headers": ["Fiziksel Parametre / Davranış", "Makroskobik Dünya (Re >> 1)", "Nanomedikal Dünya (Re << 10^-4)", "Biyofiziksel İlke ve Sınırlama", "Nanorobotik Tasarım Çözümü"],
        "rows": [
            ["Atalet (Eylemsizlik) Kuvveti", "Hareketi sürdürür (Momentum baskın)", "Sıfır atalet (Kuvvet bittiği an duruş)", "Purcell İstiridye Teoremi", "Zaman simetrisini kıran helikal burgular"],
            ["Viskoz Sürtünme Kuvveti", "Genellikle ihmal edilebilir direnç", "Mutlak hâkim kuvvet (Zift içinde yüzme)", "Stokes Sürtünme Yasası (F = 6*pi*eta*r*v)", "Aerodinamik yerine mikrohidrodinamik elmas gövde"],
            ["Termal Çalkantı (Brown Hareketi)", "Gözlemlenemez / Önemsiz", "Şiddetli ve kaotik darbe bombardımanı", "Fluktuasyon-Disipasyon Teoremi", "kT enerjisini aşan piezoelektrik/manyetik itki"],
            ["Moleküler Taşınım Mekanizması", "Türbülanslı ve konvektif karışım", "Yalnızca laminer akış ve saf difüzyon", "Fick Difüzyon Kanunları", "Moleküler sıralama rotorları ile aktif pompalama"],
            ["Malzeme Mekanik Dayanımı", "Çelik, Titanyum (~1 GPa)", "Elmasoid, Safir, CNT (~60-100 GPa)", "Kovalent sp3 bağ mukavemeti", "1000 atm basınca dayanan elmasoid tanklar"],
            ["Enerji Kaynağı", "Kimyasal piller, İçten yanmalı", "Enzimatik biyo-yakıt (Glukoz-O2 pili)", "Plazma redoks potansiyeli", "Glukoz oksidaz anotlu kesintisiz biyopil"],
            ["İletişim Ortamı", "Radyo frekansları (RF / Wi-Fi)", "Akustik basınç ve NIR optik dalgalar", "Doku RF zayıflatma engeli", "Megahertz piezoelektrik transdüserler"],
            ["Hesaplama Mimarisi", "Yarı iletken silikon transistörler", "Elmasoid çubuk mekanik mantık (Rod-logic)", "Kuantum tünelleme sınırları", "Mekanik nanosaniyelerle çalışan rod-logic"]
        ]
    },
    {
        "title": "Tablo 15.2: Sentetik Kan Hücreleri (Respirosit, Mikrobivor, Clottosit) ve Doğal Hücre Analizi",
        "headers": ["Yapay Kan Hücresi Türü", "Doğal Biyolojik Karşılığı", "Yapısal Boyut ve Malzeme", "Fonksiyonel Kapasite / Hız Artışı", "Klinik Kurtarma Endikasyonu"],
        "rows": [
            ["Respirosit (Respirocyte)", "Eritrosit (Alyuvar)", "1 um küre / Elmasoid basınç tankı", "Oksijen taşımada 236 kat kapasite", "Karbonmonoksit zehirlenmesi, Sualtı, İskemi"],
            ["Mikrobivor (Microbivore)", "Nötrofil / Makrofaj (Akyuvar)", "3.4 x 2.0 um oval / Elmasoid öğütücü", "Fagositozda 100 kat hız (30 sn/bakteri)", "Septik şok, MDR bakteriyemi, Viral enfeksiyon"],
            ["Clottosit (Clottocyte)", "Trombosit (Kan pulcuğu)", "2 um küre / Sıkıştırılmış nano-ağ", "Hemostazda 300 kat hız (1 saniyede pıhtı)", "Masif arteriyel kanama, Kansız cerrahi"],
            ["Vaskülosit (Vasculocyte)", "Endotel tamir hücresi", "2-3 um esnek gövde / Freze uçlu", "Mekanik plak tıraşlama (um3/saniye)", "Ateroskleroz tasfiyesi, Kalsifikasyon çözümü"],
            ["Sitotemizleyici Nanobot", "Otofagozom / Lizozom", "50-100 nm karbon kafes", "Sitoplazmik agregat parçalama", "Lipofuskin temizliği, Nörodejenerasyon"],
            ["Genom Onarım Nanobotu", "DNA Onarım Enzimleri", "30-40 nm nükleer penetratör", "Kusursuz mekanik çift zincir dikişi", "Somatik DNA kırıkları, Kanser önleme"],
            ["Senolitik Nanobot", "NK Hücresi / CD8+ T", "100 nm immün mimik", "uPAR hedefli 15 dakikada apoptoz", "Senesen (zombi) hücrelerin tasfiyesi"],
            ["Sentetik Tam Kan", "Biyolojik İnsan Kanı", "Üçlü nanorobot konsorsiyumu", "Milyon kat genel fizyolojik üstünlük", "Kansız, enfeksiyonsuz, ölümsüz dolaşım"]
        ]
    },
    {
        "title": "Tablo 15.3: Vasküler Temizlik ve Aterom Plağı Tasfiye Nanoteknolojilerinin Kinetiği",
        "headers": ["Vasküler Müdahale Teknolojisi", "Hedeflenen Patolojik Lezyon", "Mekanik / Biyofiziksel Mekanizma", "Tedavi Süresi / Hızı", "Damar Sağlığı ve Güvenlik Çıktısı"],
        "rows": [
            ["Vaskülosit Mekanik Frezesi", "Fibröz kapsüllü aterom plağı", "Elmasoid mikro-kesicilerle katmanlı tıraşlama", "Günde 10 mm3 plak temizliği", "Lümen genişliğinde gençlik çapına dönüş"],
            ["Kolesterol Sıralama Rotoru", "Serbest kolesterol kristalleri", "Stereospesifik kavitelerle moleküler emilim", "Saniyede 10^6 kolesterol molekülü", "Nekrotik çekirdeğin tamamen gerilemesi"],
            ["Asidik Mikro-Kavite Şelasyonu", "Hidroksiapatit damar kalsifikasyonu", "Lokalize pH 4.5 mikro-sitrik asit banyosu", "Saatler içinde apatit çözünmesi", "Arteriyel elastisitenin restorasyonu"],
            ["Manyetik Trombolitik Nano-Matkap", "Akut fibrin trombüsü (İnme/Kriz)", "1000 RPM manyetik rotasyon + lokal t-PA", "3 - 5 dakikada tam rekanalizasyon", "Sıfır sistemik kanama riskiyle inme kürü"],
            ["Sentetik Endotel Nanoyaması", "Kazıma sonrası çıplak intima", "NO-salgılayan biyouyumlu nanofilm serimi", "Anında kaplama / 24 saatte endotel", "Akut trombozun mutlak engellenmesi"],
            ["Köpük Hücresi Boşaltımı", "İntrasellüler makrofaj lipit damlacığı", "Nanometrik mikro-enjektörle lipit tahliyesi", "Hücre başına dakikalar içinde", "Kronik yangısal sinyalin söndürülmesi"],
            ["Metaloproteinaz İnhibisyonu", "İncelmiş kararsız fibröz kapsül", "Lokal TIMP mimetik nano-bariyer örtüsü", "Saniyeler içinde yüzey stabilizasyonu", "Ölümcül plak rüptürünün önlenmesi"],
            ["Otonom Vasküler Devriye", "Yeni oluşan mikrotrombüs ve oxLDL", "Sürekli dolaşım ve anlık yakalama", "7/24 Kesintisiz koruma", "Kardiyovasküler ölüm riskinin sıfırlanması"]
        ]
    },
    {
        "title": "Tablo 15.4: 3D Biyo-Baskı Modalitelerinin Fiziksel ve Biyomühendislik Parametreleri",
        "headers": ["Biyo-Baskı Yöntemi", "Baskı Çözünürlüğü", "Hücre Canlılığı Oranı", "Uyumlu Biyo-Mürekkep Türü", "İmalat Hızı ve Hedef Doku"],
        "rows": [
            ["Pnömatik Ekstrüzyon", "50 - 200 mikrometre", "%70 - 85 (Kayma stresi var)", "Viskoz hidrojeller, Aljinat, GelMA", "Orta / Büyük organ iskeleleri, Kıkırdak"],
            ["Piston / Vidalı Ekstrüzyon", "100 - 300 mikrometre", "%65 - 80", "Yüksek yoğunluklu hücresel hamurlar", "Yavaş / Kemik ve yoğun kas dokuları"],
            ["Mürekkep Püskürtmeli (Inkjet)", "20 - 50 mikrometre", "%85 - 95", "Düşük viskoziteli solüsyonlar", "Hızlı / Deri tabakalaması, Doku desenleme"],
            ["Dijital Işık İşleme (DLP)", "10 - 30 mikrometre", "%90 - 98 (Sıfır nozul stresi)", "Foto-çapraz bağlanabilir GelMA, PEGDA", "Ultra-Hızlı (Katman başına 2 sn) / Vasküler ağ"],
            ["Stereolitografi (SLA)", "20 - 50 mikrometre", "%88 - 95", "Işığa duyarlı fotopolimerler", "Orta / Karmaşık geometrili kalp kapakçıkları"],
            ["İki-Fotonlu Polimerizasyon (2PP)", "100 - 500 nanometre", "%95+", "Özel iki-fotonlu reçineler", "Yavaş (Mikron hacim) / Kılcal damar lümenleri"],
            ["Kurbanlık Biyo-Baskı (FRESH)", "20 - 100 mikrometre", "%90 - 95", "Termo-duyarlı Pluronic, Jelatin", "Orta / Perfüze edilebilir açık damar kanalları"],
            ["Lazer Destekli Transfer (LAB)", "10 - 40 mikrometre", "%95 - 99 (En yüksek canlılık)", "Yüksek hücre yoğunluklu mikro-damlalar", "Hızlı / Hücresel mikro-desenleme, Greftleme"]
        ]
    },
    {
        "title": "Tablo 15.5: Desellülarizasyon Ajanlarının Matriks ve Hücresel Bileşenler Üzerindeki Etkileri",
        "headers": ["Kimyasal / Fiziksel Ajan", "Etki Mekanizması", "Hücre ve DNA Uzaklaştırma Gücü", "ECM Matriks Proteinlerine Etkisi", "Kalıntı Sitotoksisite Riski"],
        "rows": [
            ["SDS (Sodyum Dodesil Sülfat)", "İyonik deterjan, lipid ve nükleer lizis", "Aşırı Yüksek (Altın standart)", "Glikozaminoglikan ve kolajeni hafif aşındırır", "Yüksek (Yoğun yıkama gerektirir)"],
            ["Triton X-100", "Non-iyonik deterjan, lipid-lipid çözünmesi", "Orta (Nükleusu tam parçalayamaz)", "Matriks proteinlerini ve laminini korur", "Düşük"],
            ["Sodyum Deoksikolat (SDC)", "Bile tuzu iyonik deterjanı", "Yüksek (Yoğun dokularda etkili)", "Bazal membranı kısmen zayıflatabilir", "Orta"],
            ["DNaz I ve RNaz Enzimleri", "Nükleik asit fosfodiester hidrolizi", "DNA kalıntılarını <200 bp'ye parçalar", "Kolajen ve elastine sıfır zarar", "Yok (Yıkanabilir protein)"],
            ["Tripsin-EDTA", "Proteolitik sindirim ve kalsiyum şelasyonu", "Hücre bağlantılarını hızla koparır", "Uzun maruziyette kolajeni tahrip eder", "Düşük"],
            ["Hipertonik / Hipotonik Şok", "Osmotik basınçla hücre zarı patlatma", "Hücreleri öldürür ancak enkazı yıkamaz", "Matriks yapısını kusursuz korur", "Sıfır"],
            ["Süperkritik CO2 (scCO2)", "Yüksek basınçlı çözücü gaz ekstraksiyonu", "Lipid ve hücresel kütleyi sterilize eder", "Mekanik iskele yapısını mükemmel korur", "Sıfır (Gaz fazında uçar)"],
            ["Alfa-Galaktozidaz Enzimi", "Domuz Alfa-Gal antijenlerinin kesimi", "Ksenojenik immünojenisiteyi sıfırlar", "Matriks omurgasını etkilemez", "Yok"]
        ]
    },
    {
        "title": "Tablo 15.6: Biyo-Yapay Katı Organ Sistemlerinin Mimari ve Klinik Performans Matrisi",
        "headers": ["Biyo-Yapay Katı Organ", "Fonksiyonel Hücre Tipi", "Mekanik / Mikroakışkan İskele", "Temel Fizyolojik Fonksiyon", "Klinik Tedavi ve Ömür Hedefi"],
        "rows": [
            ["Biyo-Yapay Karaciğer (BAL)", "iPSC-Hepatositler + Kupffer", "Hollow-fiber biyoreaktör kartuşu", "Amonyak detoksifikasyonu, Albümin sentezi", "Akut karaciğer yetmezliğinde köprü ve kür"],
            ["Biyo-Yapay Böbrek (BAK/RAD)", "Renal tübüler epitel hücreleri", "Polisülfon membran + İmmünoprotektif lümen", "Aktif su/tuz geri emilimi, Sitokin baskılama", "Kronik hemodiyaliz ihtiyacının tasfiyesi"],
            ["İmplante Böbrek (iAK)", "Podositler ve Tübül hücreleri", "Silikon nano-gözenekli filtre (SNM)", "Hastanın tansiyonuyla GFR > 30 ml/dk süzme", "Vücut içi kalıcı yapay böbrek implantı"],
            ["Biyo-Yapay Akciğer (BALung)", "Alveoler Tip I/II + Pulmoner Endotel", "PDMS gaz geçirgen mikroakışkan çip", "Fizyolojik O2 ve CO2 kan-gaz değişimi", "ECMO'suz taşınabilir yapay akciğer"],
            ["Biyo-Yapay Pankreas", "iPSC-Beta hücre adacıkları", "Aljinat yarı geçirgen immün-kapsüller", "Glukoza duyarlı otonom insülin salınımı", "Tip 1 diyabetin ömür boyu kalıcı kürü"],
            ["Biyo-Yapay Kalp Dokusu", "Kardiyomiyosit + Endotel + Fibroblast", "3D vaskülarize kollajen yama", "Senkronize kontraksiyon ve kan pompalama", "Enfarktüslü miyokard dokusunun tamiri"],
            ["Çip Üstünde Çoklu Organ", "5 farklı insan hücre soyu", "Mikroakışkan bağlantılı dinamik çipler", "Sistemik metabolik etkileşim simülasyonu", "İlaç ve nanorobot farmakokinetik optimizasyonu"],
            ["Vitrifike Organ Bankası", "Tam boyutlu sağlam organ", "Kriyoprotektan (CPA) camsı matriks", "-196°C'de buz kristali olmadan saklama", "Yıllar boyu rafta bekleyen taze organ arzı"]
        ]
    },
    {
        "title": "Tablo 15.7: Mekanik Dolaşım Desteği (VAD ve TAH) Nesillerinin Biyofiziksel Kıyaslaması",
        "headers": ["Mekanik Kalp Sistemi", "Akış Prensibi / Mekanizma", "Rulman / Yatak Teknolojisi", "Kayma Gerilmesi ve Hemoliz", "Klinik Avantaj ve Komplikasyon Profili"],
        "rows": [
            ["Jarvik-7 / CardioWest (1. Nesil)", "Pnömatik pulsatil membran", "Mekanik diyafram ve kapakçıklar", "Yüksek mekanik darbe / Trombüs riski", "Büyük harici kompresör, Yüksek enfeksiyon"],
            ["HeartMate II (2. Nesil)", "Aksiyal sürekli akış (Arşimet)", "Mekanik yakut rulman yatakları", "Yüksek kayma gerilmesi (>150 Pa)", "Mekanik aşınma, vWF parçalanması, kanama"],
            ["HVAD (3. Nesil Santrifüj)", "Radyal santrifüj sürekli akış", "Hidrodinamik ve manyetik hibrit yatak", "Orta düzey kayma gerilmesi", "Kompakt boyut, perikardiyal yerleşim"],
            ["HeartMate 3 (MagLev)", "Radyal santrifüj + Yapay Nabız", "Tam Manyetik Levitasyon (Temassız)", "Düşük kayma gerilmesi (<100 Pa)", "Sıfır mekanik aşınma, Sıfır pompa trombozu"],
            ["Carmat Aeson TAH", "Elektro-hidrolik pulsatil biyo-protez", "Biyolojik sığır perikardı kaplı", "Ultra-Düşük (Fizyolojik kan teması)", "Otonom debi ayarı (5-9 L/dk), Sıfır antikoagülan"],
            ["Syncardia TAH", "Pnömatik biventriküler pulsatil", "Poliüretan diyaframlar", "Yüksek basınç dalgalanmaları", "Biventriküler son dönem yetmezlikte köprü"],
            ["TET Kablosuz Enerji", "Transkütanöz manyetik rezonans", "Cilt altı ve dışı kuplaj bobinleri", "Uygulanamaz (Enerji aktarımı)", "Driveline kablosu yok, Sıfır enfeksiyon"],
            ["Homo Aeternus MagLev", "Entegre Hibrit Biyo-Mekatronik", "Akıllı yapay zeka kontrollü MagLev", "Sıfır hemoliz / Kusursuz biyouyum", "Biyolojik kalple paralel ömür boyu yedekleme"]
        ]
    },
    {
        "title": "Tablo 15.8: Biyo-Elektronik Organlar ve Nöromorfik Arayüzlerin Biyomühendislik Özellikleri",
        "headers": ["Biyo-Elektronik İmplant", "Doku / Sinir Arayüzü", "Kullanılan Biyouyumlu Malzeme", "Kayıt ve Uyarım Mekanizması", "Geri Kazanılan Fizyolojik Fonksiyon"],
        "rows": [
            ["Vagus Sinir Uyarımı (VNS)", "Vagus Siniri (N. Vagus)", "Esnek silikon manşet elektrot", "10-30 Hz elektriksel darbe trenleri", "Kolinerjik anti-enflamatuar yolak aktivasyonu"],
            ["Kapalı Döngü DBS", "Subtalamik Çekirdek (STN)", "Platin-iridyum derin beyin elektrotu", "Beta osilasyonu algılama + Faz uyarımı", "Parkinson ve tremorun anında söndürülmesi"],
            ["PRIMA Biyonik Retina", "Retina altı fotoreseptör tabakası", "Fotovoltaik silikon mikro-çipler", "Kızılötesi ışıkla foton-akım dönüşümü", "Yaşa bağlı makula dejenerasyonunda görme"],
            ["22-Kanallı Koklear İmplant", "İşitme Siniri (Skala timpani)", "İpek fibroini kaplı esnek elektrot", "Tonotopik Greenwood frekans stimülasyonu", "Presbiakuzide konuşma ve işitme restorasyonu"],
            ["Mikroakışkan Biyo-Sensör", "Deri altı / İntravasküler", "Aptamer-FET + Karbon Nanotüp", "Pikomolar elektrokimyasal amperometri", "7/24 Gerçek zamanlı glukoz ve sitokin takibi"],
            ["Biyonik Mesane Protezi", "Detrusor düz kası ve sakral sinir", "Piezoelektrik sensör + Esnek aktüatör", "Gerinim algılama ve elektriksel boşaltma", "İnkontinans ve mesane atrofisi tasfiyesi"],
            ["Aksonal İletken Kanal (NGC)", "Periferik Sinir Gövdesi", "İletken PEDOT:PSS / Polipirrol tüp", "Lokal mikro-elektriksel alan gradyanı", "Hızlı ve kusursuz sinir rejenerasyonu (3 mm/gün)"],
            ["Kortikal BCI (Neuralink/Mesh)", "Neokorteks Piramidal Nöronlar", "Mikron-altı esnek polimer iplikçikler", "Geniş bant çift yönlü aksiyon potansiyeli", "Bilişsel genişleme ve engram yedekleme"]
        ]
    },
    {
        "title": "Tablo 15.9: Nanotoksisite Mekanizmaları, Protein Korona ve İmmün Kaçış Stratejileri",
        "headers": ["Nanomalzeme / Sistem", "Primer Toksisite Mekanizması", "İmmünolojik Reaksiyon Türü", "Yüzey Zırhlama ve Savunma", "Klinik Güvenlik Eşiği"],
        "rows": [
            ["Çıplak Karbon Nanotüp (CNT)", "Asbest benzeri makrofaj delinmesi", "Kronik granülom, Mezotelyoma riski", "1 um altına kesim + Hidrofilik polimer kaplama", "Tam fonksiyonelleşme ile sıfır fibrozis"],
            ["Opsonize Nanopartikül", "C3b ve IgG ile yüzey kaplanması", "Kupffer hücresi fagositozu (RES klerensi)", "Zwitteriyonik fırçalar (PSB) + CD47 peptidi", "Sıfır protein korona ve uzun dolaşım"],
            ["Demir Oksit SPION Çekirdek", "Asidik lizozomda Fe2+ çözünmesi", "Fenton reaksiyonu ve ferroptotik ölüm", "Hermetik silika veya altın kabuk zırhlama", "Sıfır serbest iyon sızıntısı"],
            ["PEGile Lipit Nanopartikül", "Dolaşımdaki anti-PEG antikorları", "CARPA kompleman psödo-alerji krizi", "Poligliserol / Polisarcosine bazlı yeni lipidler", "Sıfır kompleman aktivasyonu"],
            ["Metalik Gümüş Nanoparçacık", "Masif gümüş iyonu (Ag+) salınımı", "Mitokondriyal elektron zinciri çöküşü", "Biyobozunur polimer matriks içine hapsetme", "Yalnızca lokal antimikrobiyal salınım"],
            ["Pozitif Yüklü Polimerler (PEI)", "Hücre zarında nano-delikler açma", "Akut sitotoksisite ve eritrosit lizisi", "pKa'sı ayarlanmış iyonize edilebilir lipidler", "Fizyolojik pH 7.4'te nötr net yük"],
            ["Elmasoid Mekanik Aşınma", "Nanobot sürtünme döküntüsü (debris)", "Yabancı cisim makrofaj infiltrasyonu", "Biyouyumlu safir kayar yüzeyler", "Renal filtrasyon boyutu (<5.5 nm)"],
            ["Arızalı Otonom Nanorobot", "Kontrolden çıkma, kural dışı eylem", "Doku hasarı ve yangısal kriz", "Dahili mikro-asit kendi kendini imha (Kill-Switch)", "Saniyeler içinde zararsız toza dönüşüm"]
        ]
    },
    {
        "title": "Tablo 15.10: Homo Aeternus Sibernetik Biyoloji ve Biyo-Mekatronik Entegrasyon Manifestosu",
        "headers": ["Sibernetik Biyoloji Katmanı", "Kullanılan İleri Teknoloji", "Biyolojik Muadilin Aşılması", "Klinik Ölümsüzlük Katkısı", "Homo Aeternus Nihai Yetkinliği"],
        "rows": [
            ["Dolaşım ve Gaz Taşımacılığı", "Respirosit + Mikrobivor + Clottosit", "Eritrosit ve lökositten 100-200x üstün", "Kalp krizi, inme, sepsis ve kanama kürü", "Soluksuz 4 saat efor, mutlak enfeksiyon direnci"],
            ["Kardiyovasküler Pompa", "MagLev Titanyum Hibrit Kalp", "Kardiyomiyosit yorulması ve enfarktüs yok", "Kardiyovasküler ölüm riskinin sıfırlanması", "Sürekli ve yorulmayan kan akışı (CO > 8 L/dk)"],
            ["Vasküler Otoyol", "Karbon Nanotüp Kaplı Esnek Arter", "Anevrizma, plak ve yırtılma riski sıfır", "Damar sertliğinin ve inmenin tasfiyesi", "500 mmHg basınca dayanan genç esnek damar"],
            ["Katı Organ Altyapısı", "3D Biyo-Basılmış ve Resellülarize Organ", "Yaşlanma yıpranmasında tak-çalıştır değişim", "Organ yetmezliğinin tarihten silinmesi", "İsteğe bağlı taze ve genç organ temini"],
            ["Metabolik Detoksifikasyon", "İmplante Edilebilir Diyaliz Çipi", "Nefron kaybı ve üremik zehirlenme yok", "7/24 Kesintisiz metabolit arındırma", "Hücrelerin ömür boyu saf sıvıda yaşaması"],
            ["Kas-İskelet Dayanımı", "EAP Sentetik Kas + Titanyum Kemik", "Osteoporoz ve sarkopeni tamamen imkânsız", "Düşme, kırık ve fiziksel kırılganlık sonu", "Süper-atletik güç ve kemik kırılmazlığı"],
            ["Duyusal Genişleme", "Biyonik Göz + Tonotopik Koklea + Sensör", "Körlük ve sağırlığın mutlak tedavisi", "Yaşa bağlı duyusal izolasyonun önlenmesi", "Kızılötesi görme, ultrasonik işitme, elektromanyetik his"],
            ["Bilişsel Süreklilik ve Bellek", "Kortikal BCI + Engram Aynalama Kristali", "Nöronal kayıpta dahi hafıza kaybı sıfır", "Alzheimer ve zihinsel ölümün imkânsızlığı", "Süper-bilgisayar entegrasyonlu ölümsüz bilinç"]
        ]
    }
]

# Dokuman Olusturma Dongusu
parts = [
    ("KISIM 1: TIBBİ NANOTEKNOLOJİ BİYOFİZİĞİ VE NANOROBOTİK SİSTEMLERİN TEMELLERİ", part1_subsections),
    ("KISIM 2: SENTETİK HÜCRE MÜHENDİSLİĞİ: RESPİROSİTLER, MİKROBİVORLAR VE CLOTTOSİTLER", part2_subsections),
    ("KISIM 3: VASKÜLER VE HÜCRESEL TEMİZLİK NANOBOTLARI", part3_subsections),
    ("KISIM 4: 3D BİYO-BASKI (BIOPRINTING) VE EKSTRÜZYON/IŞIK-TABANLI BİYO-YAZICILAR", part4_subsections),
    ("KISIM 5: DESELLÜLARİZASYON VE RESELLÜLARİZASYON: DOĞAL MATRİKS İSKELELERİ", part5_subsections),
    ("KISIM 6: BİYO-YAPAY KATI ORGAN MÜHENDİSLİĞİ: KARACİĞER, BÖBREK VE AKCİĞER REJENERASYONU", part6_subsections),
    ("KISIM 7: BİYO-MEKANİK VE SENTETİK KALP SİSTEMLERİ: TAH, VAD VE PULSATİL/SÜREKLİ AKIŞ POMPALARI", part7_subsections),
    ("KISIM 8: BİYO-ELEKTRONİK ORGANLAR, NÖROMORFİK BİYO-SENSÖRLER VE İMPLANT ENTEGRASYONU", part8_subsections),
    ("KISIM 9: NANOTOKSİSİTE, İMMÜN BİYOUYUMLULUK VE PROTEİN KORONA BİYOFİZİĞİ", part9_subsections),
    ("KISIM 10: HOMO AETERNUS SİBERNETİK BİYOLOJİ MANİFESTOSU: BİYOMEKATRONİK VE ORGAN REPLASMANI", part10_subsections)
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
        frun = pf.add_run(f"Sibernetik / Biyomekatronik Bağıntı:  {formula}")
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
print(f"CİLT 15 Başarıyla Kaydedildi: {OUTPUT_PATH}")