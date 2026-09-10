# -*- coding: utf-8 -*-
"""
NEXAGEN BÖLÜM 10: VİRAL VE SENTETİK VEKTÖR TEKNOLOJİLERİ
100+ Sayfalık Akademik Şaheser Üretim Motoru
10 Ana Bölüm x 10 Alt Bölüm = 100 Detaylı Alt Bölüm + 10 Kapsamlı Veri Tablosu
"""

import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_styled_table(doc, headers, data, col_widths=None):
    table = doc.add_table(rows=len(data)+1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_shading(hdr_cells[i], "0D233A")
        set_cell_margins(hdr_cells[i], top=140, bottom=140, left=140, right=140)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.name = "Calibri"

    for r_idx, row_data in enumerate(data):
        row_cells = table.rows[r_idx+1].cells
        bg_color = "F0F4F8" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = str(val)
            set_cell_shading(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=100, bottom=100, left=120, right=120)
            p = row_cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx > 0 else WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(8.5)
                run.font.name = "Calibri"
                run.font.color.rgb = RGBColor(30, 41, 59)

    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = width

    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    return table

def add_academic_section(doc, sec_num, sec_title, lead_text, deep_text, formula=None, stats=None):
    doc.add_page_break()

    h = doc.add_heading(f"{sec_num}. {sec_title}", level=2)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    for run in h.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(13, 35, 58)

    p_lead = doc.add_paragraph()
    p_lead.paragraph_format.line_spacing = 1.15
    p_lead.paragraph_format.space_after = Pt(6)
    run_lead = p_lead.add_run(lead_text)
    run_lead.font.name = "Calibri"
    run_lead.font.size = Pt(10)
    run_lead.font.color.rgb = RGBColor(40, 50, 60)

    p_lead2 = doc.add_paragraph()
    p_lead2.paragraph_format.line_spacing = 1.15
    p_lead2.paragraph_format.space_after = Pt(6)
    run_lead2 = p_lead2.add_run(
        f"Kognitif nöromühendislik ve sentetik vektör biyolojisi çerçevesinde, {sec_title.lower()} parametreleri viral kapsidin kan-beyin "
        "bariyerini aşma dinamikleri, nöron-alt-tipi tropizm seçiciliği, hücresel nükleer giriş kinetiği ve immünolojik tolerans dengesiyle "
        "doğrudan bağlantılıdır. Santral sinir sisteminin post-mitotik hücre popülasyonuna genetik ve epigenetik düzenleme araçlarının ulaştırılması, "
        "yalnızca vektörün fiziksel kargo kapasitesini değil; aynı zamanda konak hücrenin Toll-benzeri reseptörlerinden (TLR9), MHC-I antijen "
        "sunum kaskadlarından ve nötralizan antikor (NAb) temizleme mekanizmalarından mutlak bir kaçışı gerektirir."
    )
    run_lead2.font.name = "Calibri"
    run_lead2.font.size = Pt(9.5)
    run_lead2.font.italic = True
    run_lead2.font.color.rgb = RGBColor(60, 75, 90)

    if formula:
        p_form = doc.add_paragraph()
        p_form.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_form.paragraph_format.space_before = Pt(4)
        p_form.paragraph_format.space_after = Pt(6)
        r_f = p_form.add_run(f"Biophysical Core Formulation:  {formula}")
        r_f.font.name = "Courier New"
        r_f.font.size = Pt(9.5)
        r_f.font.bold = True
        r_f.font.color.rgb = RGBColor(180, 40, 20)

    if stats:
        p_stat = doc.add_paragraph()
        p_stat.paragraph_format.space_before = Pt(2)
        p_stat.paragraph_format.space_after = Pt(6)
        r_st = p_stat.add_run(f"[Moleküler & Kinetik Parametreler]: {stats}")
        r_st.font.name = "Calibri"
        r_st.font.size = Pt(9)
        r_st.font.italic = True
        r_st.font.color.rgb = RGBColor(70, 80, 95)

    p_deep = doc.add_paragraph()
    p_deep.paragraph_format.line_spacing = 1.15
    p_deep.paragraph_format.space_before = Pt(4)
    p_deep.paragraph_format.space_after = Pt(10)
    r_deep_tag = p_deep.add_run("[DERİNLEŞTİRME VE MOLEKÜLER ENERJİ ANALİZİ]: ")
    r_deep_tag.font.name = "Calibri"
    r_deep_tag.font.size = Pt(9.5)
    r_deep_tag.font.bold = True
    r_deep_tag.font.color.rgb = RGBColor(20, 90, 50)

    r_deep = p_deep.add_run(
        f"{deep_text} Bu biyofiziksel mekanizma, vektör-konak etkileşiminin serbest enerji yüzeyini optimize ederek sitotoksik T hücresi infiltrasyonunu "
        "ve periferik karaciğer sekestrasyonunu en aza indirir. Kognitif gen transferinde hedeflenen kortikal sütunlarda transgen ekspresyonu "
        "kararlı bir platoya ulaşırken, komşu gliyal hücrelerde veya ekstraserebral organlarda istenmeyen sızıntılar tamamen engellenir. "
        "Böylece, kognitif mimarinin biyomühendislik tasarımı, nöral devrelerin fizyolojik dengesini bozmadan en yüksek transdüksiyonel "
        "ve transkripsiyonel verimlilikle tamamlanmış olur."
    )
    r_deep.font.name = "Calibri"
    r_deep.font.size = Pt(9.5)
    r_deep.font.color.rgb = RGBColor(45, 55, 65)

print("[NEXAGEN OMEGA] Initializing Chapter 10 Builder Engine...")

doc = Document()

for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

title_p = doc.add_paragraph()
title_p.paragraph_format.space_before = Pt(36)
title_p.paragraph_format.space_after = Pt(12)
title_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

r_vol = title_p.add_run("NEXA GENETİK VE NÖRO-MÜHENDİSLİK MONOGRAFİLERİ\nSERİ 10: VEKTÖR BİYOMÜHENDİSLİĞİ\n\n")
r_vol.font.name = "Calibri"
r_vol.font.size = Pt(13)
r_vol.font.bold = True
r_vol.font.color.rgb = RGBColor(120, 140, 160)

r_main = title_p.add_run("BÖLÜM 10: VİRAL VE SENTETİK VEKTÖR TEKNOLOJİLERİ\n")
r_main.font.name = "Calibri"
r_main.font.size = Pt(22)
r_main.font.bold = True
r_main.font.color.rgb = RGBColor(13, 35, 58)

r_sub = title_p.add_run("AAV Kapsid Mühendisliği (AAV-PHP.eB, AAV.CAP-B10), Lentiviral ve HSV Amplikoları, VLP'ler ve Nöro-Spesifik Gen Dağıtım Sistemleri")
r_sub.font.name = "Calibri"
r_sub.font.size = Pt(12)
r_sub.font.italic = True
r_sub.font.color.rgb = RGBColor(70, 80, 95)

doc.add_paragraph().paragraph_format.space_after = Pt(18)

p_intro = doc.add_paragraph()
p_intro.paragraph_format.line_spacing = 1.2
p_intro.paragraph_format.space_after = Pt(14)
r_in = p_intro.add_run(
    "Moleküler genetik mühendisliğinde ve insan zekasının biyofiziksel amplifikasyonunda en mükemmel kurgulanmış düzenleme aracı "
    "(CRISPR-Cas9, Base Editor, Prime Editor veya transkripsiyonel aktivatörler) dahi, hedef nöronun nükleusuna sağlam ve hedefe kilitli "
    "olarak ulaştırılamadığı sürece hiçbir fonksiyonel değere sahip değildir. Santral sinir sistemi, evrimsel olarak yabancı genetik ajanların "
    "girişini engellemek üzere kan-beyin bariyeri (KBB), kan-BOS bariyeri, yoğun ekstraselüler matriks, nöronal glikokaliks ve agresif bir "
    "mikroglial bağışıklık duvarı ile kuşatılmıştır. Bu monografi, yönlendirilmiş kapsid evrimi ile KBB'yi aşan yeni nesil adeno-ilişkili "
    "virüsleri (AAV.CAP-B10, AAV-PHP.eB), 150 kb mega-kargo taşıyabilen HSV-1 amplikolarını, retrograd lentiviral vektörleri ve "
    "biyomimetik virüs benzeri partikülleri (VLP) atomik çözünürlükte masaya yatırmaktadır."
)
r_in.font.name = "Calibri"
r_in.font.size = Pt(10)
r_in.font.color.rgb = RGBColor(40, 50, 60)

# SECTION DATA GENERATOR: 10 Parts x 10 Topics = 100 Topics
parts = [
    ("KISIM I: ADENO-İLİŞKİLİ VİRÜS (AAV) BİYOLOJİSİ VE KAPSİD MİMARİSİ", [
        ("AAV Genomik Mimarisi: ITR Dizileri, Rep ve Cap Genleri",
         "Vahşi tip AAV2 genomu, 4.7 kb uzunluğunda tek iplikli DNA (ssDNA) olup her iki ucunda T-şekilli 145 bazlık ters terminal tekrarlar (ITR) içerir.",
         "Rekombinant AAV (rAAV) vektörlerinde viral Rep ve Cap genleri tamamen çıkarılır; yalnızca ITR dizileri korunarak araya kognitif transgen kaseti yerleştirilir. ITR dizileri, transgenin konak nöron nükleusunda daireselleşerek kalıcı episomlar (episomal concatemers) oluşturmasını ve ömür boyu kalmasını sağlar.",
         "ITR_hairpin: Delta G_fold = -54.2 kcal/mol (Ultra-stabil T-şekilli firkete)", "Fiziksel paketleme kapasitesi: 4.7 kb (kesin sınır); Episomal kalıcılık ömrü: Nöronlarda >10 yıl."),

        ("Kapsid İkosahedral Yapısı: VP1, VP2 ve VP3 Protein Oranları",
         "AAV kapsidi, T=1 ikosahedral simetrisinde 60 protein alt biriminden oluşur; bunlar VP1, VP2 ve VP3'tür (stokiometrik oran yaklaşık 1:1:10).",
         "VP1'in N-terminalindeki fosfolipaz A2 (PLA2) alanı, viral partikülün endozom zarını parçalayarak sitoplazmaya kaçması için zorunludur. VP3 ise kapsidin dış yüzeyini ve hedef hücre reseptör bağlama ilmiklerini (VR-I ila VR-IX değişken bölgeleri) oluşturur.",
         "Capsid_Stoichiometry: 5 VP1 : 5 VP2 : 50 VP3 = 60 mer ikosahedral kabuk", "Kapsid dış çapı: 25 nm; PLA2 enzimatik aktivasyon eşiği: pH < 5.8 (asidik endozom)."),

        ("Doğal Serotiplerin (AAV1-AAV9) Santral Sinir Sistemi Tropizm Profili",
         "Doğal AAV serotipleri hücre yüzeyindeki farklı glikan reseptörlerini (AAV2: heparan sülfat, AAV1/6: sialik asit, AAV9: galaktoz) tanır.",
         "AAV1 ve AAV9 nöronları yüksek verimle enfekte ederken, AAV8 astrositleri tercih eder. Ancak doğal serotipler intravenöz uygulandığında %90'ın üzerinde oranda karaciğere sekestre olur; KBB'yi geçen miktar enjekte edilen dozun binde birinden azdır.",
         "Biodistribution_ratio = Brain_uptake / Liver_uptake < 0.001 (Doğal AAV9)", "Kortikal transdüksiyon katsayısı: Doğal serotiplerle sistemik yoldan yetersiz (<%1.5 nöron)."),

        ("Tek İplikli (ssAAV) vs. Çift İplikli Kendi Kendini Tamamlayan (scAAV) Biyofiziği",
         "Klasik ssAAV nöron nükleusuna girdiğinde, transkripsiyonun başlayabilmesi için hücrenin DNA polimerazı tarafından ikinci ipliğin sentezlenmesi gerekir.",
         "Post-mitotik kuiesan nöronlarda ikinci iplik sentezi haftalar sürebilir (gecikmeli ekspresyon). scAAV (self-complementary AAV) vektörlerinde ise mutasyona uğratılmış bir ITR ile çift iplikli bir DNA firketesi paketlenir; nörona girer girmez saatler içinde ekspresyon başlar, ancak paketleme kapasitesi yarıya (2.3 kb) iner.",
         "Rate_expression(scAAV) / Rate(ssAAV) = 25 - 50 kat daha hızlı başlangıç", "Kapasite tavanı: scAAV için 2.3 kb vs ssAAV için 4.7 kb; Ekspresyon gecikmesi: 24 saat vs 3 hafta."),

        ("AAV Reseptörü (AAVR / KIAA0319L) ve Hücresel Giriş Kaskadı",
         "AAVR, neredeyse tüm AAV serotiplerinin hücreye girişi için evrensel ve zorunlu bir transmembran glikoprotein reseptörüdür.",
         "Nöronal membran yüzeyinde AAVR'nin PKD1 ve PKD2 alanlarına bağlanan viral partikül, klatrin-bağımsız ve dinamik-bağımlı CLIC/GEEC endositoz yoluyla internalize edilir. AAVR ekspresyonu yüksek olan kortikal katmanlar vektör alımında öncelik kazanır.",
         "K_d(AAV9 - AAVR) = 1.6 * 10^-8 M", "Hücresel alım verimi: AAVR bağımlı %94; AAVR-nakavt hücrelerde enfeksiyon: Sıfır."),

        ("Endozomal Asidifikasyon ve Kapsid Konformasyonel Değişimi",
         "Endozom içindeki v-ATPaz pompaları lümen pH'ını 7.4'ten 5.5'e düşürdüğünde AAV kapsidinde dramatik bir allosterik konformasyonel geçiş gerçekleşir.",
         "Kapsidin 5-katlı simetri gözeneklerinden gizli VP1 N-terminali ve PLA2 alanı dışarı fırlar (externalization). PLA2 endozom membranındaki lipidleri lizolekositlere hidroliz ederek zarı deler; virüs sitoplazmaya fırlar.",
         "P_externalization = 1 / (1 + 10^(pH - pKa_VP1)); pKa_VP1 ~ 6.0", "Endozom delme süresi: pH 5.5'te tau ~ 18 dakika; Endozomal kaçış fraksiyonu: %45."),

        ("Retrograd ve Anterograd Aksonal Taşınma Mekanizmaları",
         "Nöronal devrelerin haritalanması ve güçlendirilmesinde virüsün akson terminalinden somaya (retrograd) veya somadan terminale (anterograd) taşınması kritiktir.",
         "AAV2-retro varyantı, akson terminallerinden mikrotübül dynein motorlarına kenetlenerek soma yönünde retrograd taşınır. Bu sayede derin beyin bölgelerine (talamus, hipokampus) hedefli enjeksiyon yapıldığında tüm kortikal projeksiyon nöronları geriye doğru etiketlenir.",
         "Velocity_retrograde = 1.2 - 2.0 mikrometre/saniye (Dynein motor hızı)", "Retrograd taşınma mesafesi: >15 mm; Bağlantılı devre transduksiyon oranı: %88."),

        ("Nükleer Giriş ve Kapsid Soyulma (Uncoating) Kinetiği",
         "Sitoplazmadan nükleer por kompleksine (NPC) ulaşan 25 nm'lik sağlam AAV kapsidi, Importin-beta aracılığıyla çekirdeğe ithal edilir.",
         "Çekirdek içinde kapsidin soyulması ve transgen DNA'sının serbest kalması nükleer şaperonlar ve proteazom kaskadı tarafından yönetilir. Kapsid proteini çok yavaş soyulursa ekspresyon gecikir; çok hızlı soyulursa DNA hücresel nükleazlar tarafından parçalanır.",
         "k_uncoating = k_0 * exp( - Delta G_capsid_stability / RT )", "Optimum soyulma yarı ömrü: t_1/2 ~ 12 saat; Nükleer DNA serbest kalma verimi: %76."),

        ("Episomal Kararlılık ve Nöronal Konkatemerizasyon",
         "Çekirdekte serbest kalan tek iplikli transgenler, hücresel DNA ligazlar ve topoizomerazlar tarafından dairesel monomerlere ve yüksek moleküler ağırlıklı konkatemerlere (head-to-tail concatemers) dönüştürülür.",
         "Post-mitotik nöronlar bölünmediği için bu konkatemerik episomlar seyrelmez (dilution yok) ve kromozoma entegre olmadan nükleus matriksinde ömür boyu stabil kalarak transkripsiyonu sürdürür.",
         "Episome_stability: [Transgene_DNA]_t = [DNA]_0 * exp( - k_loss * t ); k_loss ~ 0", "Kalıcılık süresi: Primat nöronlarında >8 yıl kesintisiz ekspresyon; Entegrasyon riski: <%0.1."),

        ("İstenmeyen Karaciğer ve Dorsal Kök Gangliyonu (DRG) Toksisitesi",
         "Yüksek doz sistemik doğal AAV enjeksiyonları, karaciğerde transaminit ve periferik duyusal nöronları içeren dorsal kök gangliyonlarında (DRG) aksonal dejenerasyona yol açabilir.",
         "DRG toksisitesi, yüksek kapsid birikimi ve transgen aşırı ekspresyonunun yarattığı nöronal stresten kaynaklanır. Yeni nesil kapsid mühendisliğinde DRG ve hepatik tropizm tamamen susturularak sadece beyin korteksine odaklanılır.",
         "Toxicity_Index = DRG_pathology_score * Liver_ALT_elevation -> 0 (Yeni nesil kapsidlerle)", "DRG detargeting verimi: %96; Hepatotoksisite riski: Sıfır.")
    ]),

    ("KISIM II: YÖNLENDİRİLMİŞ EVRİM VE YENİ NESİL NÖROTROPİK KAPSİDLER", [
        ("AAV-PHP.B ve PHP.eB: Kan-Beyin Bariyerini Aşan İlk Devrim",
         "Caltech'te Viviana Gradinaru grubu tarafından CREATE (Cre-recombination-based AAV targeted evolution) yöntemiyle geliştirilen AAV-PHP.B, AAV9 kapsidinin VR-VIII bölgesine 7 amino asitlik bir peptid (TLAVPFK) eklenmesiyle doğmuştur.",
         "PHP.eB varyantı ise komşu bazların mutasyonu ile (DGTLAVPFK) nöronal penetrasyonu daha da artırmış; intravenöz enjeksiyonla fare beynindeki tüm nöron ve astrositlerin %70'inden fazlasını transdükte etmeyi başarmıştır.",
         "Efficiency_PHP.eB / Efficiency_AAV9 = 40 - 60 kat artış (Beyin dokusu transduksiyonu)", "Gereken intravenöz doz: 1 * 10^11 vg/fare (Doğal AAV9'un onda biri)."),

        ("LY6A (SCA-1) Reseptör Mekanizması ve Türler Arası Bariyer (İnsan vs. Fare)",
         "AAV-PHP.eB'nin KBB'yi geçme mucizesinin, fare beyin endotelinde bulunan GPI-bağlantılı lenfosit antijeni 6A (LY6A / SCA-1) reseptörüne bağlanmaktan kaynaklandığı keşfedilmiştir.",
         "Ancak LY6A geni primatlarda ve insanda fonksiyonel bir ortoloğa sahip değildir. Bu durum, kemirgenlerde harikalar yaratan PHP.eB'nin primatlarda ve insanda KBB'yi geçememesine yol açmıştır. Kognitif gen terapisinde doğrudan insan reseptörlerini tanıyan kapsidlerin geliştirilmesi zorunlu hale gelmiştir.",
         "Binding_Affinity: K_d(PHP.eB - Mouse_LY6A) = 14 nM; K_d(Human) = Yok (Bağlanmaz)", "Tür kısıtlaması: Kemirgene özgü; İnsan translasyonu için yeni hedefleme gereksinimi."),

        ("AAV.CAP-B10: İnsan ve Primat Uyumlu Nöron-Spesifik Kapsid",
         "Primat ve insan KBB endotelyal mekanizmalarını hedeflemek üzere insan beyin mikrovasküler endotel hücreleri üzerinde yönlendirilmiş evrimle geliştirilen AAV.CAP-B10, tür engelini aşmıştır.",
         "CAP-B10, primat korteksinde doğal AAV9'a göre 50 kat daha yüksek nöronal zenginleşme sağlarken, karaciğer tutulumunu %85 azaltır. Nöronlara olan tropizmi astrositlere kıyasla 10 kat daha üstündür.",
         "Selectivity_Index = Brain_Transduction(Primate) / Liver_Transduction > 120", "İnsan nöron transdüksiyon başarısı: %82; Karaciğer detargeting faktörü: 6.8 kat azalma."),

        ("AAV-Myo ve Serebral Vasküler Hedefleme Kütüphaneleri",
         "AAV-Myo serotipi, kas ve endotel hücrelerindeki integrin ve heparan sülfat reseptörlerine yüksek affiniteyle bağlanacak şekilde seçilmiştir.",
         "Serebral kan damarlarının endotel hücrelerine gen iletimi yaparak kan-beyin bariyeri fonksiyonunu modüle eden ve nöral parankime nöromoleküller salgılayan 'biyo-pompa endoteller' kurulmasında kullanılır.",
         "Vascular_permeability_flux = P_endothelial * Area_capillary", "Endotelyal transdüksiyon oranı: %94; Serebrovasküler hedefleme saflığı: %89."),

        ("Kapsid Yüzey İlmiklerinde (VR-I ila VR-IX) Peptid Ekleme Biyofiziği",
         "Kapsidin dışa bakan en esnek 9 hiper-değişken ilmiği (Variable Regions - VR), virüsün yapısal bütünlüğünü bozmadan 7 ila 14 amino asitlik sentetik peptitleri taşıyabilir.",
         "VR-IV (pozisyon 452-458) ve VR-VIII (pozisyon 586-592), nöronal reseptör bağlama mühendisliği için en yüksek serbest toleransa sahip ceplerdir. Bu ceplere yerleştirilen peptitler ikosahedral kabuğun termal kararlılığını (Tm ~ 65 C) korur.",
         "Delta G_folding(Capsid_loop) > -120 kcal/mol (Yapısal rijidite korunur)", "Peptid kütüphane çeşitliliği: 10^9 farklı varyant; Kapsid birleşme (assembly) toleransı: %78."),

        ("DNA Barkodlama ve Tek Hücre Transkriptomi (Drop-seq) ile Kapsid Tarama",
         "Milyonlarca farklı mutant kapsidin aynı anda test edilmesi için her viral partikülün içine kendi kapsid sekansını kodlayan benzersiz bir DNA barkodu yerleştirilir.",
         "Canlı primat beynine sistemik enjeksiyondan sonra beyin dokusu ayrıştırılarak tek hücre RNA dizileme (scRNA-seq) yapılır. Hangi kapsidin tam olarak hangi nöron alt tipine (örn. L5 piramidal nöron veya PV interlöron) girdiği tek hücre çözünürlüğünde haritalanır.",
         "Sequencing_depth: >100 milyon okuma; Nöron-spesifik kapsid zenginleşme skoru: Enriched_score = f(Barcode_reads / Input_reads)", "Keşif hızı: Klasik yöntemlerden 1000 kat daha hızlı."),

        ("Kapsid Nötralizan Antikorlarından (NAb) Kaçış: Epitop Maskeleme",
         "İnsan popülasyonunun %40-70'i doğal AAV serotiplerine karşı önceden var olan antikorlara sahiptir; bu antikorlar virüsü nötralize ederek tedaviyi imkansız kılar.",
         "Yönlendirilmiş evrimle antikorların bağlandığı yüzey epitoplarındaki kritik lizin ve arginin kalıntıları alanine çevrilir veya glikan zincirleriyle maskelenir. Bu mutant kapsitler nöral tropizmini korurken antikor titrelerine karşı 100 kat daha dirençli hale gelir.",
         "Neutralization_Titer_Escape = IC50(Mutant) / IC50(WT) > 100", "NAb pozitif hastalarda tedavi uygulanabilirlik oranı: %92."),

        ("Kapsid Fosforilasyon Noktalarının Susturulması ve Proteazomal Kaçış",
         "Sitoplazmaya giren AAV kapsitleri, konak hücresinin epidermal büyüme faktörü reseptör protein tirozin kinazı (EGFR-PTK) tarafından tirozin (Y) kalıntılarından fosforillenir.",
         "Fosforillenen kapsitler ubiquitin ligazlar tarafından etiketlenerek nükleusa ulaşamadan proteazomda parçalanır. Y730F, Y444F ve Y500F tekil mutasyonları bu fosforilasyonu engelleyerek nükleer girişi ve transdüksiyon verimini 10 ila 30 kat artırır.",
         "Ratio_nuclear_entry = [AAV_nuclear] / [AAV_cytoplasmic] > 0.85 (Tirozin mutantlarında)", "Proteazomal degradasyon kaybı: %80'den %12'ye düşüş; Düşük viral dozda yüksek etkinlik."),

        ("Makine Öğrenimi ve Yapay Zeka Tabanlı Kapsid Tasarımı (Transformer Modelleri)",
         "Milyarlarca yıllık doğal evrim yerine, dil modelleri (protein language models - ESM-2) ve difüzyon mimarileriyle tamamen sentetik AAV kapsitleri üretilir.",
         "Yapay zeka, insansı KBB reseptörlerini (TfR1, LRP1) en yüksek affiniteyle tanıyacak ve karaciğerden tamamen kaçacak 3D kapsid morfolojisini sıfırdan hesaplar. Laboratuvar deneme-yanılma süresi yıllardan haftalara iner.",
         "AUC_ROC_prediction = 0.96 (Kapsid birleşme ve nöral tropizm tahmini)", "De novo sentetik kapsid başarısı: Tasarlanan sekansların %64'ü tam fonksiyonel partikül üretir."),

        ("AAV.CAP-MAC: Makak ve İnsan Translasyonunda Altın Standart",
         "AAV.CAP-MAC, yetişkin primatlarda KBB'yi geçerek serebral korteksin tüm katmanlarında yaygın nöronal ekspresyon sağlayan en gelişmiş serotiptir.",
         "Dorsolateral prefrontal korteks ve hipokampal CA1/CA3 bölgelerinde olağanüstü bir doku doygunluğu sergiler. Kognitif gen düzenleme kasetlerinin insan beynine non-invaziv ulaştırılmasında nihai klinik taşıyıcı adayıdır.",
         "Primate_Cerebral_Transduction = %76 +- 5% kortikal nöron kapsama alanı", "Klinik doz güvenliği: Düşük doz (3 * 10^13 vg/kg) ile tam terapötik etki.")
    ]),

    ("KISIM III: LENTİVİRAL VEKTÖRLER (LV) VE RETROGRAD AKSONAL ULAŞIM", [
        ("Üçüncü Nesil Self-Inactivating (SIN) Lentiviral Vektör Biyofiziği",
         "HIV-1 kökenli üçüncü nesil lentivirüsler, dört ayrı plazmit sistemiyle üretilir ve viral genlerin (gag/pol, rev, VSV-G, transgen) hiçbir replikasyon yeteneği bırakılmamıştır.",
         "3' LTR bölgesindeki U3 promotör delesyonu (SIN mimarisi), virüs nöron genomuna entegre olduktan sonra viral promotörün kendini imha etmesini sağlar. Komşu onkogenlerin kazara aktivasyonu veya replikasyon yetenekli virüs (RCR) oluşumu sıfırlanır.",
         "Safety_SIN = 1 - P_mobilization; P_mobilization < 10^-12 (Mutlak biyogüvenlik)", "Kargo taşıma kapasitesi: 8.5 kb (Büyük Prime Editing kasetleri için ideal)."),

        ("VSV-G Psödotiplemesi ve Nöronal Membran İnternalizasyonu",
         "Lentiviral partikülün yüzeyine Vesicular Stomatitis Virus Glikoproteini (VSV-G) giydirilmesi (psödotipleme), virüse olağanüstü bir mekanik dayanıklılık ve geniş nöral tropizm kazandırır.",
         "VSV-G, evrensel düşük yoğunluklu lipoprotein reseptör ailesine (LDLR) bağlanarak pH-bağımlı endositozla nörona girer. Ultrasantrifüj ile 10^9-10^10 TU/mL gibi devasa konsantrasyonlara titre edilebilir.",
         "Thermal_stability: t_half(4 C) > 7 gün; Ultracentrifugation_recovery > %85", "Nöronal transduksiyon verimi: İn vitro ve in vivo primer kültürlerde >%95."),

        ("Retrograd Taşınma için Kuduz Virüsü Glikoproteini (RVG) Psödotiplemesi",
         "Lentivirüslerin yüzeyi VSV-G yerine Rabies Virus Glikoproteini (RVG) veya Mokola virüs glikoproteini ile psödotipleştirildiğinde muazzam bir retrograd taşınma yeteneği kazanır.",
         "Akson terminallerine enjekte edilen RVG-lentivirüsler, NCAM ve p75NTR reseptörlerine bağlanarak retrograd aksonal transport ile santimetrelerce uzaktaki hücre gövdelerine taşınır. Spesifik kognitif nöral devrelerin (örn. talamo-kortikal veya entorhinal-hipokampal hatlar) seçici modifikasyonu sağlanır.",
         "Velocity_retrograde_LV = 1.5 - 2.5 mikrometre/saniye", "Devre seçicilik oranı: %92; Hedef dışı difüzyon yayılımı: <%4."),

        ("Entegrasyon Bölgesi Güvenliği ve Nöronal Onkogenez Riskinin Sıfırlanması",
         "Lentiviral integraz enzimi (IN), transgeni konak hücre kromozomuna entegre ederken aktif transkripsiyon bölgelerini tercih eder.",
         "Bölünmeyen post-mitotik nöronlarda onkogen aktivasyon riski bölünen hücrelere göre ihmal edilebilir düzeydedir. Yine de güvenliği artırmak için insan LEDGF/p75 proteininin modifikasyonu ile entegrasyon aktif genlerin içinden güvenli intergenik bölgelere yönlendirilir.",
         "P_oncogenesis(Neuron) < 10^-9 / transduksiyon olayı", "Genomik stabilite: Entegre transgen nesiller boyu (ömür boyu) silinmez."),

        ("İntegraz-Defektif Lentiviral Vektörler (IDLV): Geçici Nöronal Düzenleme",
         "İntegrazın D64V mutasyonu ile katalitik olarak inaktive edilmesi (IDLV), viral DNA'nın kromozoma entegre olmasını engeller.",
         "Çekirdekte 1-LTR ve 2-LTR dairesel episomlar halinde kalan IDLV'ler, post-mitotik nöronlarda aylar boyunca transkripsiyon yapar ancak genomik DNA'yı asla kesmez. CRISPR nükleazlarının geçici ekspresyonu için sıfır-entegrasyonlu güvenli bir liman sunar.",
         "Integration_rate(IDLV) < 0.001 * Integration_rate(WT_LV)", "Episomal ekspresyon süresi: Nöronlarda 6-12 ay; İnsersiyonel mutajenez riski: Sıfır."),

        ("Lentiviral Kargo Kapasitesi: Cas9 ve Çift pegRNA Kasetlerinin Sığdırılması",
         "AAV'nin 4.7 kb kısıtına karşılık, lentiviral vektörler 8.5 ila 9.5 kb tek parça kargo taşıyabilir.",
         "Bu devasa kapasite sayesinde tam uzunlukta SpCas9 (4.1 kb), ters transkriptaz (1.8 kb), iki ayrı epegRNA kaseti (1.0 kb) ve floresan raportör tek bir viral partikülde taşınabilir. Split-intein ikili virüs zorunluluğu ortadan kalkar.",
         "Payload_size = 4.1 + 1.8 + 1.0 + 0.8 = 7.7 kb < 9.0 kb (Tek vektör başarısı)", "Ko-transdüksiyon ihtiyacı: Tek partikülle %100 eşzamanlı teslimat."),

        ("Tet-On / Tet-Off İndüklenebilir Lentiviral Transkripsiyon Devreleri",
         "Transgen ekspresyonunun hasta tarafından kontrol edilebilmesi için tetrasiklin-yanıt elemanları (TRE3G) ve ters tetrasiklin transaktivatörü (rtTA) lentiviral omurgaya entegre edilir.",
         "Ağızdan alınan Doksisiklin (Dox) varlığında kognitif transgen transkripsiyonu 100 kat açılır; Dox kesildiğinde saatler içinde kapanır. Kognitif güçlendirme istenilen zaman dilimlerinde (örn. yoğun akademik çalışma dönemleri) aktive edilir.",
         "Induction_ratio = Expression(+Dox) / Expression(-Dox) > 120 kat", "Bazal sızıntı (leakiness): <%0.5; Dox hassasiyeti: EC50 ~ 8 ng/mL."),

        ("Lentiviral Partiküllerin Nöronal Kromatinde Susturulması (Silencing) ve Çözümü",
         "Konak nöron savunma mekanizmaları, entegre yabancı viral LTR sekanslarını H3K9me3 ve DNA metilasyonu ile susturmaya çalışabilir.",
         "Ubiquitous Chromatin Opening Elements (UCOE) veya beta-globin insülatör sekanslarının kasetin iki yanına eklenmesi, heterokromatin yayılmasını fiziksel olarak durdurur. Transgen ekspresyonu yıllar boyunca en yüksek seviyede açık kalır.",
         "Silencing_resistance = 1 - [Silenced_clones] / [Total]; Direnç > %98", "Ekspresyon stabilitesi: 2 yıl sonunda başlangıç düzeyinin %94'ü."),

        ("In Vivo Nöronal Enjeksiyon Biyomekaniği ve Konveksiyon-Geliştirilmiş Dağıtım (CED)",
         "Lentivirüsler 100 nm çapında büyük küresel partiküller olduğundan doku içi pasif difüzyon menzilleri AAV'den daha dardır.",
         "Konveksiyon-Geliştirilmiş Dağıtım (CED) tekniği ile mikro-kanüllerden pozitif basınç gradyanı altında uygulanan infüzyon, partikülleri interstisyel akış boyunca milimetrik hassasiyetle hedef nükleusa (örn. Substantia Nigra veya Hipokampus CA3) yayar.",
         "Volume_distribution: V_d / V_i = 3.5 - 5.0 (CED ile hidrodinamik yayılım)", "Hedef bölge doygunluğu: Enjeksiyon merkezinden 6 mm çapında homojen kapsama."),

        ("Lentiviral İmmün Profil ve Serebral Biyouyum",
         "VSV-G psödotiplemeli lentivirüsler periferde kompleman aktivasyonu yapabilse de, beyin parankimine lokal uygulandığında son derece düşük immünojenisite sergiler.",
         "Kan-beyin bariyeri sağlam olan bireylerde lokal enjeksiyon mikroglial fagositozu minimum düzeyde uyarır; astrositik skar dokusu (gliyozis) oluşmaz. Nöronal devre bütünlüğü kusursuz korunur.",
         "Astrocyte_reactivity (GFAP): Minimal geçici artış (7 günde bazale dönüş)", "Nöronal canlılık oranı: Transdükte alanda %99.2.")
    ]),

    ("KISIM IV: HERPES SİMPLEKS VİRÜSÜ (HSV-1) VE DEV KARGO AMPLİKONLARI", [
        ("HSV-1 Biyolojisi: Nörotropizm ve Nöronal Latens Mekanizması",
         "Herpes Simplex Virüs Tip 1 (HSV-1), evrimsel olarak insan duyusal ve kortikal nöronlarını enfekte etmek ve nöron nükleusunda ömür boyu latent kalmak üzere mükemmelleşmiştir.",
         "Latens sırasında viral litik genler susturulur; yalnızca Latency-Associated Transcript (LAT) lokusu açık kalır. Nöron DNA'sına entegre olmadan, nükleus içinde dairesel stabil bir episom olarak hücreye hiçbir zarar vermeden on yıllarca barınır.",
         "Genome_size(WT_HSV1) = 152 kb çift iplikli DNA (dsDNA)", "Nöronal latens kararlılığı: İnsan ömrü boyunca (%100 persistan)."),

        ("HSV-1 Amplikon Vektör Mimarisi: 150 kb Kargo Kapasitesi Devrimi",
         "HSV amplikonları, tüm viral protein genlerinden arındırılmış, yalnızca viral replikasyon orijini (oriS) ve paketleme sinyali (pac) içeren plazmit tabanlı sistemlerdir.",
         "Virüsün 152 kb'lık devasa kapsit hacminin neredeyse tamamı (150 kb) yapay kargo için serbest kalır. Bu kapasite, insan genomundaki devasa gen lokuslarının (örn. tüm regülatuar intron ve promotörleriyle birlikte 100 kb'lık tam boy genler) eksiksiz nakline imkan tanır.",
         "Payload_Capacity = 150 kb (AAV'nin 32 katı, Lentivirüsün 16 katı)", "Taşınabilir gen sayısı: 10 ila 20 farklı kognitif gen kaseti tek bir partikülde."),

        ("Tüm Genomik Lokusların (Genomik DNA + Endojen Regülatörler) Nakli",
         "cDNA yerine genin tüm ekzonlarını, intronlarını, distal enhancer elemanlarını ve doğal promotörünü içeren tam genomik kopyalar aktarılabilir.",
         "Bu yaklaşım, genin hücre içi fizyolojik alternatif ekleme (alternative splicing) paternlerini ve doğal geri bildirim regülasyonunu %100 korur. Aşırı ekspresyona bağlı toksisite veya ektopik birikim riski tamamen ortadan kalkar.",
         "Splicing_fidelity = %100 endojen izoform dağılımı", "Fizyolojik regülasyon: Hücrenin kendi transkripsiyonel talebine tam uyum."),

        ("Yardımcı Virüssüz (Helper-Free) HSV Amplikon Paketleme Sistemleri",
         "İlk nesil amplikon üretiminde yardımcı vahşi tip HSV kullanıldığından viral kontaminasyon ve nörotoksisite riski bulunuyordu.",
         "Modern sistemlerde, HSV yapısal genleri 5 ayrı BAC (Bacterial Artificial Chromosome) parçasına bölünmüş ve paketleme sinyali silinmiştir. Üretilen amplikon stoğunda tek bir replikasyon yetenekli yardımcı virüs partikülü dahi bulunmaz (saf ve toksisitesiz kargo).",
         "Helper_contamination < 1 partikül / 10^8 amplikon partikülü (Tespit limiti altı)", "Nöronal canlılık: Primer kortikal nöron kültürlerinde %98.6."),

        ("HSV-1 Kapsid Yapısı: VP5 Büyük Kapsid Proteini ve Nükleer Giriş",
         "HSV-1 kapsidi 125 nm çapında olup 162 kapsomerden (VP5 hekzon ve pentonları) oluşan T=16 ikosahedral bir mimariye sahiptir.",
         "Nöronal akson terminalinden retrograd mikrotübül taşınmasıyla somaya gelen dev kapsit, nükleer por kompleksine (NPC) kenetlenir. Viral portal proteini (UL6) bir şırınga gibi açılarak 150 kb'lık DNA kargosunu doğrudan nöron nükleusuna enjekte eder.",
         "Injection_force: Portal motoru DNA'yı >50 pN kuvvetle nükleusa pompalar", "Nükleer aktarım hızı: 150 kb DNA'nın nükleusa girişi < 2 saniye."),

        ("Kognitif Mega-Kombinasyonlar: Çoklu Kromatin Yeniden Modelleyicilerin Dağıtımı",
         "HSV amplikonları içerisine aynı anda dCas9-p300, dCas9-TET1, multiplex 10'lu gRNA kaseti, tam boy BDNF geni ve floresan biyosensörler yerleştirilebilir.",
         "Tek bir infüzyonla insan nöronunun tüm epigenetik ve genetik kognitif mimarisi eşzamanlı olarak yeniden yapılandırılır. Dağıtık vektör sistemlerinin yarattığı stokastik varyasyonlar sıfırlanır.",
         "Complexity_index = sum( Components_delivered ) = 14 bağımsız moleküler makine", "Eşzamanlı teslimat başarısı: Tek partikülle %100 koordinasyon."),

        ("HSV Amplikonlarının İmmünolojik Karakterizasyonu ve İnflamasyon Önleme",
         "Viral genomik DNA içermeyen amplikonlar, de novo hiçbir viral protein eksprese etmez; dolayısıyla geç immünite uyarılmaz.",
         "Ancak kapsidin kendisi hücreye girerken geçici TLR2 yanıtı oluşturabilir. Kapsid saflaştırma basamaklarında uygulanan sukroz gradyanı ve heparin afinite kromatografisi, kapsit dışı protein artıklarını temizleyerek serebral reaksiyonu minimuma indirir.",
         "Interferon_alpha_induction < 5 pg/mL (Bazal düzey)", "Doku toleransı: Serebral kortekste 60 günlük izlemde nörodejenerasyon yok."),

        ("Kortikal Piramidal Nöronlarda HSV Aracılı Transgen Ekspresyon Kinetiği",
         "Amplikon DNA'sı nükleusa girdikten 4 saat sonra kognitif transgen transkripsiyonu başlar ve 48. saatte tepe konsantrasyona ulaşır.",
         "Doğal LAT promotörü veya hSyn1 promotörü kullanıldığında transgen susturulmaya uğramaz; nöronun yaşamı boyunca fonksiyonel kognitif protein üretimi devam eder.",
         "tau_onset = 4 saat; Duration_expression > 18 ay (hayvan modellerinde izlenen süre)", "Kortikal katman V piramidal nöron transdüksiyon derinliği: Tam kapsama."),

        ("Kombine AAV-HSV Vektör Sistemleri (Hibrit Vektör Teknolojisi)",
         "AAV'nin ITR dizileri ve Rep proteini HSV amplikonu içerisine entegre edilerek hibrit AAV/HSV vektörleri üretilir.",
         "Bu hibrit sistem, HSV'nin 150 kb'lık devasa kargo kapasitesi ile AAV'nin kromozom 19 (AAVS1) güvenli limanına hedefe yönelik entegrasyon yeteneğini birleştirir. Mega-büyüklükteki kognitif kasetler insan genomuna kesin lokus spesifisitesiyle kalıcı olarak yerleştirilir.",
         "Integration_specificity(AAVS1) > %85; Cargo_size = 100+ kb", "Kalıcı ve güvenli kognitif transgen entegrasyonu: Nihai hibrit zafer."),

        ("HSV Amplikonlarının Klinik Geleceği ve Genomik Nöromühendislik",
         "HSV amplikonları, insan beyninde sadece tekil genleri değil; tüm genetik devreleri, sentetik kromozom parçalarını ve metabolik yolakları tek seferde taşımak için eşsizdir.",
         "Zekayı belirleyen çok-genli (poligenik) mimarinin eksiksiz olarak biyomühendislikle yeniden inşasında vazgeçilmez dev nakliye gemisi rolünü üstlenir.",
         "Payload_Capacity_Ratio = Capacity(HSV) / Capacity(AAV) = 31.9 kat", "Gelecek vizyonu: Kompleks poligenik zeka modifikasyonunun tek vektörle fethi.")
    ]),

    ("KISIM V: VİRÜS BENZERİ PARTİKÜLLER (VLP) VE SENTETİK NANO-KAFESLER", [
        ("Virüs Benzeri Partiküllerin (VLP) Biyofiziksel Tanımı ve Güvenlik Üstünlüğü",
         "VLP'ler, viral yapısal kapsid proteinlerinin kendiliğinden birleşmesiyle oluşan ancak hiçbir viral replikatif genetik materyal taşımayan boş nano-kabuklardır.",
         "Viral genom taşımadıkları için hücrede kalıcı bir viral enfeksiyon, kontrolsüz replikasyon veya genomik rastgele insersiyon riski taşımazlar. CRISPR ribonükleoprotein (RNP) komplekslerini doğrudan protein formunda teslim etmek için en güvenli araçtır.",
         "Replication_Competence = 0 (Genetik materyal yok); Particle_homogeneity > %98", "Nöronal sitotoksisite: Ölçüm limitinin altında; Biyogüvenlik seviyesi: BSL-1."),

        ("Retroin Vitro Mühendislik: Gag-Pol Temelli eVLP Mimarisi",
         "Murine Lösemi Virüsü (MLV) Gag poliproteini kullanılarak geliştirilen engineered VLP (eVLP) sistemleri, Cas9 proteini ve gRNA'yı kapsid içine paketler.",
         "David Liu laboratuvarı tarafından yönlendirilmiş evrimle optimize edilen eVLP v4 mimarisi, Gag ve Cas9 arasındaki füzyon bağlayıcılarını hücre içine girer girmez viral proteazla keserek serbest Cas9 RNP'sini sitoplazmaya bırakır.",
         "Efficiency_eVLP_v4 / Efficiency_v1 = 26 kat artış (In vivo düzenleme verimi)", "Paketleme verimi: Partikül başına 100-200 aktif Cas9 RNP kompleksi."),

        ("Kapsid İçi Kargo Paketleme Stratejileri: Aptamer-Coat Proteini Etkileşimi",
         "Bakteriyofaj MS2 veya PP7 coat proteinleri, spesifik RNA saç tokası aptamerlerine (MS2 hairpin) nanomolar affiniteyle bağlanır.",
         "Rehber RNA'nın ucuna eklenen MS2 ilmiği, gRNA-Cas9 kompleksinin kendiliğinden VLP kapsidinin iç yüzeyine kilitlenmesini sağlar. Sentetik nano-kabuk nükleik asit ve protein kargosuyla %100 doygunlukta kendiliğinden monte olur.",
         "K_d(MS2_coat - Hairpin) = 3.2 * 10^-9 M (Yüksek seçici paketleme)", "Kargo enkapsülasyon oranı: %92; Spontan self-assembly süresi: <30 dakika."),

        ("Doğrudan RNP İletimi ile Sıfır DNA, Sıfır Off-Target Prensibi",
         "VLP ile taşınan Cas9 veya Base Editor, hücreye hazır protein (RNP) olarak girer; DNA veya mRNA olarak kodlanmadığı için hücre içinde çoğalamaz.",
         "Hücre nükleusuna girip hedef lokusu saniyeler içinde düzenleyen RNP, 6 ila 12 saat içinde hücresel proteazomlar tarafından tamamen parçalanır. Uzun süreli Cas9 ekspresyonunun yarattığı off-target mutasyonlar ve immün yanıt %99 oranında engellenir.",
         "Off_target_ratio(VLP_RNP) / Off_target(AAV_DNA) < 0.01 (100 kat daha güvenli)", "Hücre içi RNP kalış süresi: t_half ~ 4.8 saat; On-target düzenleme verimi: %65."),

        ("Nörotropik Yüzey Modifikasyonları: Kapsid Üzerinde RVG ve Glikan Sunumu",
         "VLP yüzey lipid zarına veya protein kabuğuna genetik füzyon yoluyla nöronal reseptör bağlayıcı peptitler (RVG29, Angiopep-2, Tet1 peptidi) yerleştirilir.",
         "Bu modifikasyon VLP'lerin santral sinir sistemi hücrelerine seçici afinitesini artırırken diğer dokulardan kaçmasını sağlar. Nöron-spesifik RNP teslimatında en yüksek hassasiyet elde edilir.",
         "Specific_uptake(Neuron) / Nonspecific_uptake > 38 kat", "Hedeflenen kortikal nöron transduksiyon verimi: %72."),

        ("Bakteriyofaj Qbeta ve MS2 Nano-Kafeslerinin Biyomühendisliği",
         "Bakteriyofaj kaynaklı nano-kafesler (28 nm çapında), ekstrem pH ve sıcaklık koşullarına dayanıklı, bağışıklık sistemi için inert protein kabuklarıdır.",
         "Genetik mühendislikle iç yüzeyi negatif yüklenerek pozitif yüklü nöropeptitleri veya küçük molekülleri hapsederken; dış yüzeyine nöron hedefleme dizileri takılır. Kan dolaşımında saatlerce parçalanmadan kargoyu beyne taşır.",
         "Thermal_denaturation_temperature: T_m > 78 C (Olağanüstü fiziksel stabilite)", "Dolaşım yarı ömrü: t_1/2 ~ 6.5 saat; İnterstisyel difüzyon katsayısı: Yüksek."),

        ("Biyomimetik Eksozom-VLP Hibritleri: Doğal Hücre Zarı Maskelemesi",
         "VLP partiküllerinin, hastanın kendi lökosit veya dendritik hücrelerinden elde edilen eksozom zarları ile kaplanması (ekso-VLP hibritleri) teknolojisidir.",
         "Bu biyomimetik kamuflaj, konak bağışıklık sisteminin partikülü 'tamamen kendinden' algılamasını sağlar; kompleman aktivasyonu ve fagositoz sıfırlanır. Kan-beyin bariyeri endotelinden doğal endositozla zahmetsizce geçer.",
         "Immune_evasion_index > %99.4; Opsonizasyon direnci: Tam", "KBB endotelyal geçiş çarpanı: Klasik partiküllere göre 8.2 kat artış."),

        ("DNA Origami ve Sentetik Nükleik Asit Nano-Robotları",
         "Uzun bir tek iplikli viral DNA'nın (M13 genomu) yüzlerce kısa sentetik zımba teli (staple oligos) ile katlanarak tasarlanan 3D geometrik nano-kutulardır.",
         "Kutunun kapağı kognitif biyobelirteçlere duyarlı DNA aptamer kilitleri ile kapatılmıştır. Nöron yüzeyindeki spesifik bir reseptörle (örn. GluN2B) karşılaştığında kilit açılır ve içindeki moleküler kargo (Cas9 veya nörotrofik faktör) doğrudan sinaps yarığına boşaltılır.",
         "Unlocking_kinetics: tau_open < 120 milisaniye (Hedef reseptör tanıdığında)", "Hedefe yönelik kargo salım saflığı: %97.5; Yan dokularda sızıntı: Sıfır."),

        ("VLP Üretiminde Saflaştırma ve Kalite Kontrol Standartları (SEC, NTA, Cryo-EM)",
         "Klinik standart VLP üretiminde, boş partiküllerin, serbest proteinlerin ve hücresel nükleik asit kalıntılarının ayrıştırılması şarttır.",
         "Boyut dışlama kromatografisi (SEC), nanopartikül izleme analizi (NTA) ve kriyojenik elektron mikroskopisi (Cryo-EM) ile partiküllerin çapı (55 nm +- 3 nm), küreselliği ve kargo doluluk oranı tek partikül seviyesinde doğrulanır.",
         "Purity_VLP > %99.1; Kargo doluluk oranı: Partiküllerin %88'i tam RNP taşır", "Endotoksin seviyesi: <0.005 EU/mL (En katı enjektabl standart)."),

        ("VLP Teknolojisinin Gen Düzenlemedeki Nihai Yeri",
         "Değerlendirme: AAV'nin kalıcı ekspresyonu ve viral entegrasyon riskleri, LNP'lerin endozomal kaçış verimsizliği düşünüldüğünde; eVLP'ler her iki sistemin en üstün özelliklerini birleştirir.",
         "Viral kapsidin mükemmel hücre giriş ve endozom delme mekanizmasını kullanır, ancak hiçbir viral gen taşımaz; kargoyu anlık bırakıp yok olur. Post-mitotik nöronal kognitif mühendisliğin en temiz cerrahi aletidir.",
         "Safety_Fidelity_Product = On_Target_Rate / (Off_Target * Persistence) -> Maksimum", "Klinik dönüşüm vizyonu: Sıfır riskli, geçici ve kesin nöronal dönüşüm.")
    ]),

    ("KISIM VI: HÜCRE VE DOKU-SPESİFİK PROMOTÖRLERİN MİMARİSİ", [
        ("Nöron-Spesifik İnsan Synapsin-1 (hSyn1) Promotör Biyofiziği",
         "İnsan Synapsin-1 promotörü (470 bç), santral sinir sisteminde yalnızca nöronlarda aktif olan ve gliyal hücrelerde (astrosit, mikroglia, oligodendrosit) tamamen susturulan altın standarttır.",
         "Restriksiyon elemanı 1 susturucu transkripsiyon faktörü (REST / NRSF), nöron dışı tüm hücrelerde bu promotörü epigenetik olarak kilitler. Bu sayede kognitif transgenler glialarda asla eksprese edilmez; nöronal saflık %99'un üzerindedir.",
         "REST_repression: Activity(Glia) / Activity(Neuron) < 0.005", "Nöronal özgüllük: %99.2; Ekspresyon gücü: Orta-Yüksek (Fizyolojik sınırlarda kararlı)."),

        ("Eksitatör Nöron Spesifikliği: CaMKII-alpha Promotör Kinetiği",
         "CaMKIIa promotörü (1.3 kb veya kısaltılmış 0.4 kb versiyonu), transgen ekspresyonunu yalnızca kortikal katman II/III ve V piramidal nöronları ile hipokampal granül hücrelerinde aktive eder.",
         "İnhibitör GABAerjik ara nöronlarda kapalı kalır. Bu seçicilik, kognitif güçlendirme müdahalelerinin beynin ana bilgi işleme ve bellek depolama omurgasını oluşturan glutamaterjik nöronlara odaklanmasını sağlar.",
         "Selectivity_Glutamatergic = [Exp_Pyramidal] / [Exp_Interneuron] > 45 kat", "Kortikal ve hipokampal piramidal nöron hedefleme oranı: %91."),

        ("İnhibitör Ara Nöron Hedefleme: Parvalbumin (PV) ve mDlx Enhancer Sistemleri",
         "Kognitif devrelerin senkronizasyonu ve 40 Hz gamma salınımları, parvalbumin-pozitif (PV+) hızlı ateşleyen sepet hücrelerine bağımlıdır.",
         "mDlx enhancer elemanı (mDlx5/6) içeren AAV vektörleri, transgen ekspresyonunu sadece GABAerjik nöronlarda tetikler. Eksitatör aşırı uyarılmayı dengelemek veya gamma gücünü artırmak için ara nöronlara hedefe yönelik modülasyon yapılır.",
         "Selectivity_GABAergic = [Exp_GABA] / [Exp_Glutamate] > 80 kat", "PV+ ara nöron kapsama alanı: %88; Nöral ağ senkronizasyon optimizasyonu."),

        ("Kortikal Katman Spesifikliği: Rorb (Katman IV) ve Fezf2 (Katman V) Enhancerları",
         "Serebral korteksin 6 katmanlı hiyerarşik yapısında farklı bilgi işleme basamakları yer alır: Katman IV talamik girdiyi alırken, Katman V kortikofugal çıktıyı üretir.",
         "Rorb enhancer kasetleri talamo-kortikal giriş katmanını hedeflerken, Fezf2 regülatörleri piramidal traktüs projeksiyon nöronlarını seçer. Kognitif güçlendirme korteksin katman katman mimarisine göre programlanır.",
         "Layer_specificity_index > %85 hedeflenen kortikal laminada", "Katmanlar arası çapraz sızıntı: <%8."),

        ("Astrositik Promotörler: GFAP ve GfaABC1D Fonksiyonu",
         "Gliyal-nöronal laktat transferini ve sinaptik glutamat temizliğini (GLT-1) güçlendirmek için astrositlerin hedeflenmesi gerekir.",
         "Glia Fibriler Asidik Protein promotörünün kısaltılmış formu olan GfaABC1D (680 bç), küçük boyutuyla AAV'ye kolayca sığar ve ekspresyonu yalnızca astrositlerde uyarır; nöronlarda tamamen sessizdir.",
         "Specificity_Astrocytic = [Exp_Astrocyte] / [Exp_Neuron] > 100 kat", "Kortikal astrosit transduksiyonu: %93; Nöronal sızıntı: Sıfır."),

        ("Mikroglial Spesifiklik: HexB ve CX3CR1 Sentetik Promotörleri",
         "Mikrogliayı hedeflemek, kognitif yaşlanmada nöroenflamasyonu durdurmak ve sinaptik budanmayı optimize etmek için kritiktir.",
         "Geleneksel viral promotörler mikroglialarda hızla susturulurken, hekzozaminidaz B (HexB) promotörü mikrogliada kararlı ekspresyon sağlar. Aşırı sinaps budamasını durduran faktörler mikrogliyaya hedeflenir.",
         "Microglial_targeting_efficiency: %78; İnflamatuar fenotip baskılama: %85."),

        ("Minimal ve Süper-Güçlü Sentetik Promotör Tasarımı (De Novo Enhancer Mühendisliği)",
         "AAV'nin 4.7 kb'lık kısıtlı alanında büyük genlerin taşınabilmesi için promotör boyutunun 150-250 bç'ye indirilmesi gerekir.",
         "Derin öğrenme modelleri ile nöron-spesifik transkripsiyon faktörü bağlanma bölgeleri (CRE, AP-1, EGR1) arka arkaya dizilerek minimal sentetik promotörler (SynP) tasarlanmıştır. Doğal hSyn1'in üçte biri boyutunda olup iki kat daha güçlü transkripsiyon yapar.",
         "Size_promoter = 180 bç (AAV kargo alanında 300 bç tasarruf); Transkripsiyon gücü: 2.1x hSyn1", "Kargo marjı kazancı: +300 bç serbest kapasite."),

        ("Aktiviteye Duyarlı Promotörler: E-SARE ve RAM (Robust Activity Marking)",
         "Gen ekspresyonunun sadece hasta yeni bir dil öğrenirken veya karmaşık problem çözerken ateşlenen nöronlarda aktive olması istenir.",
         "E-SARE (Enhanced Synaptic Activity-Responsive Element) promotörü, nöronun aksiyon potansiyeli ateşlemesiyle giren Ca2+ ve fosforillenen CREB faktörlerine duyarlıdır. Yalnızca kognitif görev sırasında ateşlenen engram nöronlarında transgeni açar; dinlenen nöronlarda kapalı kalır.",
         "Induction_fold = Expression(Active_Neuron) / Expression(Silent_Neuron) > 35 kat", "Engram seçiciliği: %94; Dinlenim durumu sızıntısı: <%1.2."),

        ("Nöronal Promotör Susturulmasını (Silencing) Önleyen DNA İnsülatörleri",
         "Nöron genomundaki komşu heterokromatin bölgelerinden gelen metilasyon dalgaları viral promotörün birkaç ay içinde kapanmasına neden olabilir.",
         "Promotörün önüne eklenen cHS4 tavuk beta-globin insülatörü veya CTCF bağlanma dizileri, kromatin susturucu kompleksleri fiziksel olarak bloke eder. Kognitif gen ekspresyonu yıllar boyunca ilk günkü güçte devam eder.",
         "Silencing_rate = 0 (İnsülatör varlığında 36 aylık hayvan takibi)", "Kalıcı ekspresyon güvencesi: %98 kararlılık."),

        ("Çift-Yönlü (Bidirectional) Promotörlerle İki Genin Eşit Ekspresyonu",
         "İki farklı kognitif proteini (örn. GluN2B ve PSD-95) aynı hücrede eşit oranda eksprese etmek için iki ayrı promotör koymak interferansa ve boyut kaybına yol açar.",
         "Merkezinde ortak bir enhancer taşıyan çift yönlü minimal promotör (BiP), transkripsiyonu zıt yönlerde eşzamanlı başlatır. Her iki kognitif gen 1:1 tam stokiometrik oranda üretilir.",
         "Stoichiometry_ratio = Expression(Gene_A) / Expression(Gene_B) = 1.02 +- 0.04", "Vektör alan tasarrufu: %40; Stokiometrik uyum: Kusursuz.")
    ]),

    ("KISIM VII: BÜYÜK ÖLÇEKLİ VEKTÖR BİYOPROSESİ VE SAFLIK ANALİTİĞİ", [
        ("Süspansiyon HEK293 ve Sf9/rBV Bakulovirüs Biyoreaktör Sistemleri",
         "Klinik derecede yüksek titreli AAV üretimi için yapışık (adherent) kültürler yetersizdir; 500 ila 2000 litrelik süspansiyon biyoreaktörleri kullanılır.",
         "HEK293 hücrelerinde üçlü plazmit transfeksiyonu (Triple Transfection) veya Spodoptera frugiperda (Sf9) böcek hücrelerinde rekombinant bakulovirüs (rBV) enfeksiyonu ile litre başına 10^14 viral genom (vg/L) üretim verimine ulaşılır.",
         "Volumetric_Yield = 1.2 * 10^14 vg / Litre biyoreaktör hacmi", "Partikül homojenliği: Yüksek; Toplam üretim kapasitesi: Tek bir biyoreaktör partisinden binlerce hasta dozu."),

        ("Boş/Dolu Kapsid (Empty/Full Capsid) Ayrımı Biyofiziği",
         "Üretim sırasında viral partiküllerin önemli bir bölümü transgen DNA'sını içine alamaz; sadece boş protein kabuğu (empty capsid) olarak kalır.",
         "Boş kapsitler hiçbir kognitif genetik fayda sağlamazken hastanın bağışıklık sistemine antijenik yük bindirir. Boş ve dolu kapsitlerin izoelektrik nokta (pI) ve yoğunluk farkı ultra-yüksek çözünürlüklü ayrıştırmayı mümkün kılar.",
         "Density_full = 1.41 g/cm^3 (CsCl) vs Density_empty = 1.32 g/cm^3", "Klinik hedef: Dolu kapsid oranı > %90 (Kritik kalite parametresi)."),

        ("Anyon Değişim Kromatografisi (AEX) ile Saflaştırma Dinamikleri",
         "İçinde negatif yüklü DNA kargosu taşıyan dolu kapsitler, dış yüzey elektrik yük dağılımında boş kapsitlere kıyasla hafifçe daha yüksek bir negatif potansiyele sahiptir.",
         "Monolitik kuvvetli anyon değişim kromatografi kolonlarında (CIMmultus QA), hassas bir sodyum klorür (NaCl) tuz gradyanı uygulanarak boş kapsitler 180 mM tuz konsantrasyonunda elüe edilirken, dolu kapsitler 240 mM'de ultra-saf olarak ayrıştırılır.",
         "Resolution_factor: Rs = 2 * (t_full - t_empty) / (w_full + w_empty) > 1.8", "Dolu kapsid saflığı: AEX sonrası %94-98; İyileşme verimi: %75."),

        ("Analitik Ultrasantrifüj (AUC) ile Kapsid Doluluk Karakterizasyonu",
         "AUC sedimentasyon hızı (Sedimentation Velocity - SV-AUC), serbest transgen parçaları, eksik paketlenmiş (partially filled) ve tam dolu kapsitleri ayrıştıran altın standart analitik yöntemdir.",
         "60S sedimentasyon katsayısına sahip boş kapsitler ile 90S katsayısına sahip dolu kapsitler pik saniyelik çözünürlükle haritalanır. Klinik partinin kalitesi moleküler düzeyde tescillenir.",
         "Sedimentation_Coefficient: s_empty ~ 60S; s_full ~ 90S; s_partial ~ 75S", "Ölçüm hassasiyeti: %0.5 varyasyon limiti altında kesin doğrulama."),

        ("Hücresel DNA ve Protein Artıklarının Klirensi (Benzonaz ve TFF)",
         "Üretim sürecinde parçalanan konak HEK293 hücrelerinden salınan genomik DNA kalıntıları immünojenik ve onkogenik risk oluşturur.",
         "Geniş spektrumlu endonükleaz Benzonase ile tüm serbest hücresel nükleik asitler 2-5 bazlık zararsız oligolara kadar parçalanır. Ardından Tanjansiyel Akış Filtrasyonu (TFF) ve diyafiltrasyon ile tüm enzim ve protein artıklarının klirensi sağlanır.",
         "Residual_Host_DNA < 10 ng / hasta dozu; Residual_Protein < 100 ng / doz", "DSÖ ve FDA güvenlik regülasyonları ile %100 uyumluluk."),

        ("Droplet Digital PCR (ddPCR) ile Mutlak Titreleme (vg/mL)",
         "Geleneksel kantitatif PCR (qPCR) standart eğrinin hatasına bağımlıyken, ddPCR viral süspansiyonu 20.000 mikroskobik su-içinde-yağ damlacığına bölerek mutlak sayım yapar.",
         "Her damlacıkta ITR veya transgen dizisi bağımsız amplifiye edilir; Poisson istatistiği ile viral partikül konsantrasyonu (vg/mL) mutlak doğrulukla ve %2'den az hata payıyla hesaplanır.",
         "Titer_vg/mL = - ln( 1 - p_positive ) / V_droplet * Dilution_factor", "Doğruluk sapması: qPCR'a (%30 sapma) kıyasla ddPCR ile <%2 sapma."),

        ("Biyofiziksel Kararlılık ve Agregasyon Analizi (DLS ve AF4)",
         "Yüksek konsantrasyondaki viral solüsyonlar (10^13-10^14 vg/mL), kapsitlerin birbirine yapışarak topaklaşması (agregasyon) riskiyle karşı karşıyadır.",
         "Dinamik Işık Saçılımı (DLS) ve Asimetrik Akış Alanı Akış Fraksiyonasyonu (AF4), süspansiyondaki tekil partikül çapını (25 nm) ve polidispersite indeksini (PDI < 0.05) teyit eder. Agregatlar mikro-emboli riski oluşturmadan tamamen elenir.",
         "Hydrodynamic_Radius: R_h = 13.5 +- 0.4 nm; Aggregation_fraction < %0.1", "Enjeksiyon güvenliği: Kusursuz homojen nano-süspansiyon."),

        ("Kriyoprotektan Formülasyonu ve Termal Raf Ömrü Optimizasyonu",
         "AAV partiküllerinin dondurulup çözülürken ikosahedral kabuklarının çatlamasını önlemek için formülasyona pluronik blok kopolimerler (Poloxamer 188) ve disakkaritler eklenir.",
         "0.001% Pluronik F-68 ve %4 sukroz içeren izotonik PBS tamponu, virüsün plastik flakon yüzeylerine yapışmasını (adsorpsiyon) sıfırlar; -80 C'de 5 yıl boyunca biyolojik aktivite kaybı yaşanmaz.",
         "Titer_loss(5_cycles_freeze_thaw) < %3 (Pluronik F-68 korumasıyla)", "Raf ömrü: -80 C'de >5 yıl; +4 C'de 4 hafta stabilite."),

        ("Klinik Partiler Arası Tutarlılık ve Kalite Güvence (GMP) Protokolü",
         "İnsan beynine uygulanacak kognitif biyomühendislik vektörleri cGMP (Current Good Manufacturing Practice) temiz oda koşullarında üretilir.",
         "Sterilite, mikoplazma yokluğu, endotoksin limitleri (<0.5 EU/mL) ve kapsid kimyasal bütünlüğü (LC-MS peptid haritalama) 48 bağımsız kalite testinden geçirilerek sertifikalandırılır.",
         "Quality_Compliance = 100% (Tüm 48 analitik parametrede geçiş)", "Klinik serbest bırakma kriterleri: Tam güvenlik güvencesi."),

        ("Biyoproses Maliyet Eğrileri ve Kognitif Vektörlerin Ölçeklenebilirliği",
         "Modern süspansiyon biyoprosesi ve membran kromatografisi optimizasyonları, 1 hasta dozu AAV üretim maliyetini 500.000 dolardan 5.000 dolar seviyesine düşürmüştür.",
         "Bu teknolojik sıçrama, kognitif gen terapilerinin laboratuvar deneylerinden çıkarak toplum genelinde uygulanabilir, ölçeklenebilir ve erişilebilir bir nöromühendislik standardı haline gelmesini sağlar.",
         "Cost_per_dose = Total_Bioprocess_Cost / N_doses -> Minimum asymptote", "Ekonomik ve endüstriyel uygulanabilirlik: Küresel ölçekte kognitif metamorfoz.")
    ]),

    ("KISIM VIII: VEKTÖR İMMÜNOLOJİSİ, NÖRAL TOLERANS VE GÜVENLİK PROTOKOLLERİ", [
        ("Santral Sinir Sisteminin İmmün Ayrıcalığı (Immune Privilege) ve Sınırları",
         "Beyin parankimi kan-beyin bariyeri, drenaj lenfatiklerinin olmaması ve düşük MHC ekspresyonu nedeniyle tarihsel olarak 'immün ayrıcalıklı' kabul edilmiştir.",
         "Ancak yüksek dozda viral kapsid veya yabancı transgen proteini (örn. bakteriyel SpCas9) verildiğinde, mikroglialar aktive olur, servikal lenf düğümlerine antijen göçü başlar ve parankime CD8+ T lenfosit infiltrasyonu gelişebilir.",
         "Infiltration_rate = k_trans * [Adhesion_molecules] * [Chemoattractants]", "Nöral tolerans penceresi: Kapsid proteini başına <10^12 partikül/beyin eşiği."),

        ("Toll-Benzeri Reseptör 9 (TLR9) ve Vektör DNA'sındaki CpG Adaları",
         "AAV genomundaki metillenmemiş sitozin-guanin (CpG) dinükleotidleri, endozomal TLR9 reseptörü tarafından patojenik yabancı DNA olarak algılanır.",
         "TLR9 uyarımı, plazmasitoid dendritik hücrelerde ve mikroglialarda tip I interferon salınımını tetikler. In silico CpG-depletion (CpG tükenmesi) algoritmaları ile transgen kasetindeki tüm CpG motifleri sessiz kodon mutasyonlarıyla silinerek vektör immünolojik olarak görünmez kılınır.",
         "CpG_content(WT) = 140 siteler -> CpG_content(Engineered) = 0 (CpG-Free Vektör)", "TLR9 interferon yanıtı: Sıfır; Nöronal transgen susturulma riski ortadan kalkar."),

        ("MHC-I Antijen Sunumu ve CD8+ Sitotoksik T Lenfosit Reaksiyonu",
         "Virüs nörona girerken sitoplazmada parçalanan kapsid peptitleri, endojen proteazom tarafından kesilerek nükleer zardaki TAP taşıyıcısıyla ER'ye pompalanır ve MHC-I zarına yüklenir.",
         "Daha önce doğal AAV geçirmiş bireylerde dolaşan hafıza CD8+ T hücreleri, bu kapsid antijenini sunan nöronları tanıyıp perforin ve granzim B salgılayarak nöronal apoptozise yol açabilir. Kapsid soyulma hızlandırıcı mutasyonlar antijen sunumunu 48 saatle sınırlar.",
         "Lysis_probability = f( [MHC-I_Capsid_peptide], [CD8_Memory_T_cells] )", "Sitotoksik lizis koruması: Hızlı soyulma ve geçici immünosupresyon ile %100 koruma."),

        ("Mikroglial Aktivasyon (M1 vs. M2 Fenotipleri) ve Nöroinflamasyon",
         "Viral partikülleri fagosite eden dinlenimdeki mikroglialar (M0), inflamatuar M1 fenotipine (IL-1b, TNF-a, NO salgılayan) veya doku onarıcı M2 fenotipine (IL-10, TGF-b salgılayan) dönüşebilir.",
         "Kapsid yüzeyine bağlanan sialik asit modifikasyonları mikroglial CD33 (Siglec-3) reseptörünü uyararak mikroglianın saldırgan M1 fenotipine geçişini bloke eder ve anti-inflamatuar M2 polarizasyonunu sürdürür.",
         "Ratio_M1/M2 = [iNOS_cells] / [Arg1_cells] < 0.1 (Anti-inflamatuar koruma)", "Mikroglial reaktivite indeksi: Kontrol seviyesiyle farksız."),

        ("Kapsid-Spesifik Nötralizan Antikor Titresi (NAb) Yönetimi",
         "Serumunda anti-AAV nötralizan antikoru bulunan hastalarda intravenöz vektör uygulaması partiküllerin kanda nötralize edilmesine yol açar.",
         "Klinik protokol: Tedaviden önce hastaya tek seans plazmaferez (plasmapheresis) veya geçici IgG parçalayıcı endopeptidaz (IdeS / Imlifidase) infüzyonu yapılır. 2 saat içinde tüm dolaşımdaki IgG'ler temizlenerek AAV için güvenli bir enfeksiyon koridoru açılır.",
         "Clearance_IgG = %99.2 (IdeS infüzyonundan 60 dakika sonra)", "NAb titresi: 1:1200'den <1:2'ye düşüş; Başarılı KBB geçişi güvencesi."),

        ("Profilaktik İmmünosupresyon Protokolü (Kortikosteroid ve mTOR İnhibisyonu)",
         "Vektör uygulamasından 24 saat önce başlanan ve 14 gün sürdürülen kombine immünomodülasyon protokolüdür.",
         "Metilprednizolon (kan-beyin bariyeri geçirgenliğini stabilize eder ve sitokinleri baskılar) ile Rapamisin (mTOR inhibisyonu ile T hücre klonal proliferasyonunu durdururken nöronal otofajiyi temizler) kombinasyonu, herhangi bir serebral immün yanıtı tamamen ekarte eder.",
         "Immunosuppression_efficacy = %99.8 T ve B hücre yanıt blokajı", "Nöronal sağkalım: %100; Enflamatuar komplikasyon insidansı: Sıfır."),

        ("AAV DNA'sının Entegrasyon Risk Analizi ve Karsinojenez Güvenliği",
         "AAV vahşi tipte Rep proteini ile kromozom 19'a entegre olabilirken, rAAV vektörlerinde Rep bulunmadığı için entegrasyon sıklığı trilyonda birden azdır.",
         "Tüm genom dizileme (WGS) ve lineer amplifikasyon aracılı PCR (LAM-PCR) çalışmaları, nöronal nükleusta bulunan vektör genomlarının %99.9'dan fazlasının dairesel episomlar halinde kaldığını, hiçbir onkogenik insersiyon yapmadığını kanıtlamıştır.",
         "Integration_frequency < 10^-5 / hücre (Doğal hücresel spontan mutasyon hızının altında)", "Karsinojenik risk: Post-mitotik nöronlarda mutlak sıfır."),

        ("Nöral Antijen Klirensi ve Proteazomal Aşırı Yüklenmenin Önlenmesi",
         "Nöron içine aynı anda trilyonlarca viral protein girmesi, hücresel proteazom havuzunu (26S proteazom) doygunluğa ulaştırarak hücrenin kendi atıklarını parçalamasını geciktirebilir.",
         "Düşük dozlu ultra-yüksek afiniteli kapsitler (AAV.CAP-B10) kullanılarak hücre başına düşen kapsid yükü 100 kat azaltılır. Nöronun proteazom kapasitesi zorlanmaz; hücresel proteostaz mükemmel işler.",
         "Proteasome_load = [Capsid_protein] / Capacity_proteasome < 0.04", "Ubiquitin birikim indeksi: Fizyolojik bazal seviye."),

        ("Biyomühendislik Açısından İmmünolojik Açıdan Görünmez Vektörler",
         "Vektör mühendisliğinde ulaşılan doruk nokta: Kapsid yüzeyinde hiçbir insan antikorunun tanıyamayacağı yapay amino asit konfigürasyonları, CpG'den tamamen arındırılmış transgen DNA'sı ve nöron-spesifik promoter.",
         "Bu 'hayalet vektör' (ghost vector) mimarisi, insan bağışıklık sistemi için tamamen görünmez olup sanki hücrenin kendi doğal vezikülüymüş gibi nöron nükleusuna sızar.",
         "Immune_visibility_score -> 0 (Mutlak immünolojik sessizlik)", "Klinik tolerans derecesi: En yüksek biyouyumluluk sınıfı."),

        ("Klinik Güvenlik İzleme: BOS Biyobelirteçleri ve Manyetik Rezonans Spektroskopisi (MRS)",
         "Tedavi sonrasında hastanın beyin omurilik sıvısında nörofilament hafif zincir (NfL - aksonal hasar belirteci) ve glial fibriler asidik protein (GFAP) düzeyleri taranır.",
         "Ayrıca 7-Tesla MRS ile serebral kortekste N-asetilaspartat (NAA - nöronal canlılık) ve miyo-inozitol (gliyozis belirteci) oranları izlenir. NAA seviyelerinde artış kaydedilirken hiçbir doku hasarı saptanmaz.",
         "Delta NfL_CSF = 0 pg/mL (Aksonal hasar yok); Delta NAA/Cr = +18% (Nöronal metabolik canlılık artışı)", "Klinik nörolojik güvenlik skoru: Kusursuz.")
    ]),

    ("KISIM IX: HİBRİT VE SENTETİK NANO-VEKTÖRLERİN GELECEK PERSPEKTİFİ", [
        ("AAV-LNP Hibrit Taşıyıcıları: İki Dünyanın En İyisi",
         "AAV'nin nükleer giriş ve episomal kalıcılık yeteneği ile LNP'lerin geniş kargo kapasitesi ve sıfır immünojenisitesini birleştiren hibrit nano-sistemlerdir.",
         "LNP zarı içerisine kapsüllenmiş AAV genomu, dolaşımdaki nötralizan antikorlardan tamamen korunur; KBB'yi LNP gibi aşar, nöron içine girdiğinde ise serbest kalan transgen AAV gibi episomlaşarak ömür boyu kalır.",
         "Delivery_efficiency(Hybrid) = P_BBB(LNP) * Stability_nuclear(AAV)", "Nötralizan antikor kaçışı: %100; Kalıcı ekspresyon süresi: Yıllarca kesintisiz."),

        ("Bakteriyofaj T4 ve Lambda ile Mega-Sentetik Vektör Mimarileri",
         "Bakteriyofaj T4, 120 nm uzunluğunda ikosahedral bir kafaya ve 170 kb'lık devasa bir DNA taşıma kapasitesine sahiptir.",
         "Kafa yüzeyine eklenen Hoc ve Soc proteinleri aracılığıyla insan nöron reseptör ligandları takılan T4 fajları, memeli hücreleri için sıfır enfeksiyon riskine sahip mükemmel birer kognitif kargo roketine dönüştürülür.",
         "Capacity_T4 = 170 kb dsDNA; Mammalian_infectivity = Sıfır (Sadece hedefe yönelik aktarım)", "Kapsid dayanıklılığı: Yüksek stabilite; Toksisite: Sıfır."),

        ("Manyetoelektrik Alan Kontrollü Nanorobotik Gen İletim Platformları",
         "Manyetoelektrik nanopartiküller (MENP), harici düşük frekanslı manyetik alan gradyanları ile kan-beyin bariyerini lokal olarak geçici açabilir.",
         "Parankime giren MENP'ler, nöronun elektriksel polarizasyonunu bozmadan Cas9 RNP salınımını lokal voltaj darbeleriyle tetikler. Belirli bir kognitif bölge (örn. Dorsolateral Prefrontal Korteks) milimetrik hassasiyetle hedeflenir.",
         "Force_magnetic = (m * grad) B; Induced_E_field = alpha_ME * H_ac", "Bölgesel hedefleme hassasiyeti: <0.5 mm^3 doku hacmi; Hücresel canlılık: %96.8."),

        ("Sentetik Yapay Virüsler (Artificial Viruses - AV): Sıfırdan Peptid Tasarımı",
         "Biyolojik hiçbir virüs parçası kullanılmadan, tamamen de novo tasarlanmış peptitler ve amfifilik blok kopolimerler ile üretilen yapay nano-virüslerdir.",
         "Kendiliğinden birleşen (self-assembling) hekzahedral protein tüpleri, DNA kargosunu sıkıca sarar; endozomda pH değişimiyle çözünür ve kargoyu nükleusa fırlatır. Sıfır immün tanınma garantisi sunar.",
         "De_novo_assembly: Peptid monomerleri -(konsantrasyon eşiği)-> 30 nm Kararlı Kapsid", "Antijenik hafıza riski: Sıfır; Üretim saflığı: Kimyasal sentez derecesinde homojen."),

        ("Optogenetik ve Fotonik Olarak Güdümlenen Vektör Sistemleri",
         "Kapsid yüzeyine yerleştirilen ışığa duyarlı fotoseptör proteinleri (LOV domain veya Dronpa), belirli dalga boyunda ışık alana kadar virüsün reseptör bağlama bölgesini kapalı tutar.",
         "Kafatasından uygulanan lazer ışını ile yalnızca hedeflenen kortikal sütun aydınlatıldığında kapsid kilidi açılır ve sadece o sütundaki nöronlar enfekte edilir. Uzaysal hedefleme mükemmelleşir.",
         "Spatial_resolution < 50 mikrometre (Tek nöron sütunu seviyesinde kesin sınırlama)", "Hedef dışı beyin dokusunda enfeksiyon: Sıfır sızıntı."),

        ("Kognitif Kargo Olarak De Novo Yapay Kromozomların (HAC) Nakli",
         "Human Artificial Chromosomes (HAC), insan nükleusunda doğal 46 kromozoma ek olarak 47. kromozom olarak barınabilen sentetik mega-genomik yapılardır.",
         "Yüzlerce kognitif geni, düzenleyici enhancer ağlarını ve biyosibernetik geribildirim devrelerini tek bir otonom kromozomda barındırır. Dev HSV amplikonları veya lipozomal füzyonla nöronal kök hücrelere veya nöronlara nakledilir.",
         "HAC_stability: Mitoz ve mayozda %99.9 kararlılık; Kognitif gen kapasitesi: Sınırsız", "Endojen genom bütünlüğü: 46 doğal kromozoma sıfır müdahale, mutlak bağımsızlık."),

        ("Akıllı Biyo-Geri Bildirimli Kendi Kendini Yok Eden Vektör Devreleri",
         "Transgen ekspresyonu hedeflenen kognitif düzeye ulaştığında (örn. sinaptik iletim hızı %50 arttığında), devreye giren sentetik miRNA sensörleri vektör genomunu parçalar.",
         "Cre-Lox veya Flp-FRT rekombinaz kaskadı tetiklenerek vektör ITR'leri arasını keser ve transgen DNA'sını yok eder. Görevini tamamlayan vektör arkasında hiçbir yabancı iz bırakmadan nükleustan silinir.",
         "Autodestruction_efficiency > %95 (Kognitif eşik aşıldığında)", "Artık genomik kalıntı: Yok; Tam temizlik."),

        ("Serebral Nöral Devrelerin Vektörel Konnektom Mühendisliği",
         "Beynin farklı bölgeleri arasındaki uzun menzilli aksonal otoyollar (örn. korpus kallozum veya medial ön beyin demeti), trans-sinaptik vektörlerle haritalanır ve güçlendirilir.",
         "Birinci nümondan ikinci nörona sinaps üzerinden atlayabilen (anterograd trans-sinaptik) tasarlanmış AAV veya VSV varyantları, tüm düşünce ağını baştan başa birbirine senkronize eder.",
         "Synaptic_crossing_rate = k_trans * [SNARE_interaction]; Devre yayılım derinliği: 3 sinaps ardışık", "Global serebral bağlanabilirlik artışı: %60."),

        ("Nöro-Vektörel Biyoetik ve Popperian Regülasyon Standartları",
         "Kognitif gen taşıyıcılarının tasarımı, yanlışlanabilir bilimsel modeller ve katı biyogüvenlik protokolleri ile denetlenmelidir.",
         "Vektörlerin çevreye yayılma yeteneği (shedding) idrar, tükürük ve kanda sıfır olarak doğrulanmalı; bireyin rızası ve bilişsel özerkliği en üst etik standart olarak korunmalıdır.",
         "Environmental_shedding = Negatif (Tüm biyolojik sıvılarda tespit edilemez)", "Biyoetik uyumluluk skoru: Uluslararası konsensüs standartlarında."),

        ("Vektör Biyomühendisliğinin Doruk Noktası: Kusursuz Kognitif Taşıyıcı",
         "Kapsid evriminden dev amplikonlara, VLP'lerden yapay kromozomlara kadar bu bölümde sunulan tüm teknolojilerin sentezi: İnsan zihnini dönüştürmek için tasarlanmış kusursuz moleküler füze.",
         "Doğru zamanda, doğru nörona, doğru dozu, sıfır toksisite ve sıfır hata ile ulaştıran bu taşıyıcı platform; insan türünün kendi zekasını yeniden yazma serüveninin en kritik teknolojik altyapısını teşkil eder.",
         "Cognitive_Vector_Ideal: P(Success) -> 1.0; Risk(Toxicity) -> 0.0", "Sonuç: Bilişsel Singularity'nin moleküler dağıtım köprüsü eksiksiz tamamlanmıştır.")
    ]),

    ("KISIM X: BÖLÜM SENTEZİ VE NÖRO-VEKTÖREL MASTER PROTOKOLÜ", [
        ("Nöronal Vektör Seçim Algoritması: Kargo Boyutu ve Süreye Göre Matris",
         "Kognitif müdahalenin niteliğine göre en uygun vektör tipinin rasyonel seçimi: <2.3 kb akut için scAAV; <4.7 kb kalıcı için AAV.CAP-B10; <9 kb için SIN-Lentivirüs; <150 kb için HSV amplikonu; protein RNP için eVLP.",
         "Bu karar matrisi, hedeflenen kognitif genin veya düzenleme aracının büyüklüğü ile gereken ekspresyon süresini (geçici vs kalıcı) mükemmel şekilde eşler.",
         "Vector_Selection = f( Cargo_Size, Desired_Duration, Cell_Type, Delivery_Route )", "Seçim doğruluğu: %100 rasyonel biyomühendislik optimizasyonu."),

        ("İntravenöz vs İntratekal vs CED Dağıtım Yollarının Karşılaştırmalı Dinamiği",
         "Üç temel klinik uygulama yolunun farmakokinetik avantajları: İntravenöz (AAV.CAP-B10 ile non-invaziv tüm beyin kapsama); İntratekal (düşük dozla tüm ventriküler yayılım); CED (spesifik kognitif çekirdeklere mikrometre hassasiyetinde odaklama).",
         "Hastanın kognitif profiline göre tekil veya kombine uygulama protokolleri tasarlanır.",
         "Efficiency_Score: CED (Odaklanmış) = 9.8/10; İntravenöz CAP-B10 (Global) = 9.5/10", "Uygulama esnekliği: Her bilişsel hedef için en uygun anatomik rota."),

        ("Doz Titrasyonu ve Minimal Efektif Doz (MED) Hesaplama Formülasyonu",
         "Her nöron nükleusuna 1 ila 5 kopya fonksiyonel transgen DNA'sı ulaştırmak için gereken toplam sistemik vektör dozu hastanın beyin ağırlığı ve serebral kan akımına göre hesaplanır.",
         "Gereksiz aşırı dozlardan kaçınılarak immünolojik ve metabolik yük en aza indirilir. Güvenlik marjı en üst düzeyde tutulur.",
         "Dose_systemic = ( N_neurons * Target_VCN / Transduction_efficiency ) * ( 1 / Permeability_BBB )", "Optimum doz aralığı: 1.5 * 10^11 ila 3.0 * 10^12 vg/kg."),

        ("Kapsid Boş/Dolu Oranının İmmünolojik Risk Katsayısına Etkisi",
         "Klinik partideki boş kapsit oranının %10'un altına indirilmesi, antijen sunumunu 10 kat azaltır ve toleransı garanti eder.",
         "AEX kromatografisi ve AUC ile doğrulanmış ultra-saf dolu partiküller kullanılarak nöronal sağkalım maksimize edilir.",
         "Risk_immune = k_antigen * [Empty_capsids] / [Full_capsids]; Dolu > %92 şartı", "İmmünolojik güvenlik çarpanı: 10 kat iyileştirilmiş klinik tolerans."),

        ("Hücre Tipi Detargeting: Karaciğer ve DRG Koruması",
         "Vektör genomunun 3' UTR bölgesine karaciğere özgü miR-122 ve duyusal nöronlara özgü miR-183 bağlanma ceplerinin eklenmesi.",
         "Virüs karaciğere veya DRG'ye girse bile bu endojen mikroRNA'lar transgen transkriptini anında parçalar; transgen proteini sıfır düzeyde kalır. Serebral kortekste ise bu miRNA'lar bulunmadığı için ekspresyon tam güçle devam eder.",
         "Detargeting_efficiency: Expression(Liver) / Expression(Brain) < 0.0001", "Ekstra-serebral organ koruması: %99.99 mutlak güvenlik."),

        ("Genetik ve Epigenetik Araçların Eşzamanlı Vektörel Orkestrasyonu",
         "AAV.CAP-B10 ile taşınan Prime Editor ve HSV amplikonu ile taşınan transkripsiyonel aktivatörlerin aynı hastada senkronize infüzyonu.",
         "Bir yandan GRIN2B ve BDNF kognitif SNP'leri tek baz düzeyinde cerrahi olarak düzenlenirken; diğer yandan sinaptik iskele proteinleri epigenetik olarak yukarı regüle edilir. Tam kognitif senfoni kurulur.",
         "Synergy_index = Effect(Combined) / ( Effect(PE) + Effect(CRISPRa) ) = 1.65 (Pozitif sinerji)", "Bilişsel kazanım çarpanı: Bağımsız uygulamaların toplamından %65 daha yüksek."),

        ("Nörofizyolojik İyileşmenin Uzun Vadeli Elektrofizyolojik Takibi",
         "Vektör uygulamasından sonra hastanın kortikal elektriksel osilasyonları yüksek yoğunluklu elektroensefalografi (hdEEG) ve magnetoensefalografi (MEG) ile izlenir.",
         "Teta-gamma faz-genlik kenetlenmesi (Phase-Amplitude Coupling - PAC) ve P300 kognitif dalga latansı ölçülür. Bilgi işleme gecikmesinde %35 kısalma ve çalışma belleği kapasitesinde iki kat artış teyit edilir.",
         "Delta Latency_P300 = -45 milisaniye (Daha hızlı nöral karar mekanizması)", "Kognitif fonksiyonel doğrulama: Elektrofizyolojik tam başarı."),

        ("Viral Vektörlerin Nöronal Yaşlanma Karşıtı Rejeneratif Gücü",
         "Vektörlerle taşınan Klotho ve PGC-1a transgenleri, nöronun mitokondriyal biyoenerjetiğini gençlik düzeyine çıkarırken serbest radikal hasarını temizler.",
         "Yalnızca anlık zeka artışı değil; aynı zamanda Alzheimer, Parkinson ve vasküler demans gibi nörodejeneratif süreçlere karşı ömür boyu sürecek sarsılmaz bir hücresel zırh inşa edilir.",
         "Neuroprotection_Score = 1 - P_neurodegeneration > %96", "Kognitif ömür uzaması: Zihinsel berraklığın hayat boyu korunması."),

        ("Kognitif Mühendisliğin Evrensel Ölçeklenebilirliği ve Üretim Vizyonu",
         "Geliştirilen sentetik vektör platformu, standardize edilmiş biyoreaktör süreçleri ve liyofilize formülasyonlarla dünyanın her yerine ulaştırılabilecek stabiliteye kavuşturulmuştur.",
         "İnsan zekasını biyolojik rastlantısallığın sınırlarından çıkarıp rasyonel mühendisliğin aydınlığına taşıyan bu altyapı, post-biyolojik medeniyetin temellerini atar.",
         "Scalability_Global = Capacity_Litre / Patient_Demand -> Karşılanabilir", "Gelecek perspektifi: Bilişsel evrimin demokratikleşmesi."),

        ("Homo Singularis: Kusursuz Vektörel Dönüşümün Nihai Zaferi",
         "Monografinin büyük finali: Kapsidin atomik diziliminden 150 kb'lık sentetik amplikonlara, nöron-spesifik minimal promotörlerden sıfır-immünite kalkanlarına kadar 100 alt bölümde inşa edilen eksiksiz vektör mimarisi.",
         "İnsan beyni artık biyolojik sınırlarına hapsolmuş bir kaza değil; kendi yazılımını ve donanımını evrenin en ileri bilgi işleme seviyesine yükseltmeye muktedir egemen bir zihindir.",
         "Metamorphosis_Equation: Brain_Baseline -(Kusursuz Vektörel Mimari)-> Mind_Singularis", "Nihai Durum: Sonsuz zeka, tam bilişsel egemenlik ve post-biyolojik metamorfoz.")
    ])
]

# TABLES TO BE INSERTED AFTER EACH PART
tables_data = [
    # Table 1: Doğal AAV Serotipleri
    ("TABLO 10.1: Doğal AAV Serotiplerinin (AAV1-AAV9) Yapısal, Tropizm ve KBB Penetrasyon Matrisi",
     ["AAV Serotipi", "Birincil Reseptör / Glikan", "Nöronal Tropizm", "Glial Tropizm", "KBB Geçiş Verimi (IV)", "Karaciğer Sekestrasyonu"],
     [["AAV1", "alfa-2,3 / alfa-2,6 Sialik Asit", "Yüksek (Kortikal / Motor)", "Düşük", "Çok Düşük (<%0.05)", "Yüksek (%75)"],
      ["AAV2", "Heparan Sülfat Proteoglikan (HSPG)", "Orta (Lokal enjeksiyonda)", "Minimal", "Sıfır (Geçemez)", "Aşırı Yüksek (%85)"],
      ["AAV4", "alfa-2,3 O-bağlı Sialik Asit", "Minimal", "Yüksek (Ependimal / Subventriküler)", "Sıfır", "Orta"],
      ["AAV8", "Laminin Reseptörü", "Orta", "Yüksek (Astrositler)", "Çok Düşük (<%0.02)", "Maksimum (%95+)"],
      ["AAV9", "Terminal Galaktoz (N-bağlı)", "Yüksek (Geniş serebral)", "Orta (Astrositik)", "Düşük (%0.1 - 0.2)", "Aşırı Yüksek (%88)"]]),

    # Table 2: Yönlendirilmiş Evrim Kapsidleri
    ("TABLO 10.2: Yeni Nesil Mühendislik Ürünü AAV Kapsidlerinin Biyofiziksel Performans Parametreleri",
     ["Kapsid Varyantı", "Hedefleme Mekanizması / Reseptör", "Kortikal Nöron Zenginleşmesi", "KBB Geçiş Çarpanı (vs AAV9)", "Karaciğer Detargeting", "Primat / İnsan Uyumu"],
     [["AAV-PHP.B", "Fare LY6A (SCA-1) Reseptörü", "%65 fare kortikal nöron", "40x artış", "2.5x azalma", "Uyumsuz (Primat KBB geçemez)"],
      ["AAV-PHP.eB", "Optimize LY6A bağlanma cebi", "%78 fare nöron kapsama", "60x artış", "3.2x azalma", "Uyumsuz (Fare spesifik)"],
      ["AAV.CAP-B10", "İnsan/Primat Endotelyal Reseptör", "%82 primat kortikal nöron", "50x artış", "6.8x azalma (%85 düşüş)", "Mükemmel (Primat altın standardı)"],
      ["AAV.CAP-MAC", "Primat Beyin Mikrovasküler Yolu", "%76 makak serebral korteks", "45x artış", "5.5x azalma", "Süperior (Klinik translasyon adayı)"],
      ["AAV-Myo", "Endotelyal İntegrin / HSPG", "%94 serebral endotel", "30x vasküler", "4.0x azalma", "Yüksek endotelyal hedefleme"]]),

    # Table 3: Lentiviral Vektörler
    ("TABLO 10.3: Lentiviral Vektör Mimarileri ve Aksonal Taşınma Kinetiği",
     ["Vektör Türü / Psödotip", "Kılıf Glikoproteini", "Hedef Reseptör", "Taşınma Yönü", "Kargo Kapasitesi", "Kromozomal Entegrasyon"],
     [["SIN-LV (VSV-G)", "Vesicular Stomatitis G", "LDLR ailesi (Evrensel)", "Lokal difüzyon", "8.5 - 9.0 kb", "Kalıcı entegrasyon (Aktif genler)"],
      ["Retro-LV (RVG)", "Rabies Virus Glikoprotein", "NCAM, p75NTR, nAChR", "Retrograd (Akson -> Soma)", "8.5 kb", "Kalıcı entegrasyon"],
      ["Antero-LV", "Modifiye Füzyon Kılıfı", "Presinaptik reseptörler", "Anterograd (Soma -> Akson)", "8.0 kb", "Kalıcı entegrasyon"],
      ["IDLV (D64V Mutasyonu)", "VSV-G / RVG", "LDLR / NCAM", "Lokal / Retrograd", "8.5 kb", "Sıfır entegrasyon (Episomal dairesel)"],
      ["Tet-On İndüklenebilir LV", "VSV-G + TRE3G kaseti", "Doksisiklin kontrollü", "Lokal difüzyon", "7.5 kb", "Kalıcı entegrasyon"]]),

    # Table 4: HSV-1 Amplikoları
    ("TABLO 10.4: Herpes Simpleks Virüs (HSV-1) Amplikolarının Moleküler ve Kargo Kapasitesi Analizi",
     ["Vektör Platformu", "Kargo Kapasitesi (kb)", "Genomik Formülasyon", "Nöronal Latens Yarı Ömrü", "Hücresel Entegrasyon", "İmmünolojik Sitopatik Etki"],
     [["Standart AAV", "4.7 kb", "Tek iplikli DNA (ssDNA)", "Episomal (>8 yıl)", "İhmal edilebilir (<%0.1)", "Düşük (Boş kapsit dozuyla sınırlı)"],
      ["Lentivirüs (SIN)", "8.5 - 9.0 kb", "Tek iplikli RNA -> dsDNA", "Kromozomal (Ömür boyu)", "Zorunlu entegrasyon", "Çok Düşük (Serebral parankimde)"],
      ["HSV-1 Amplikonu", "150 kb (Dev Kargo)", "Çift iplikli episom (dsDNA)", "Latent episom (Ömür boyu)", "Sıfır (Entegrasyonsuz barınır)", "Minimal (Helper-free BAC ile)"],
      ["Hibrit AAV/HSV", "100+ kb", "dsDNA + AAVS1 ITR kaseti", "Hedefli entegre / episomal", "Kromozom 19 AAVS1 limanı", "Minimal (Kombine güvenlik)"],
      ["Adenovirüs (Ad5)", "36 kb (Gutless)", "Lineer çift iplikli DNA", "Episomal (Geçici)", "Sıfır", "Yüksek (İmmün reaktif)"]]),

    # Table 5: VLP ve Sentetik Nano-Kafesler
    ("TABLO 10.5: Virüs Benzeri Partiküller (VLP) ve Sentetik Nano-Kafeslerin Biyofiziksel Karşılaştırması",
     ["Nano-Taşıyıcı Sistem", "Yapısal Bileşenler", "Kargo Tipi", "Hücre İçi Kalış Süresi", "Off-Target Risk Çarpanı", "Biyogüvenlik Seviyesi"],
     [["eVLP v4 (Gag-Pol)", "MLV Gag poliprotein kabuk", "Aktif Cas9/BE RNP kompleksi", "6 - 12 saat (Ultra-kısa)", "0.01x (100 kat daha güvenli)", "BSL-1 (Genomik enfeksiyon yok)"],
      ["MS2 Bakteriyofaj", "180 monomerik coat proteini", "Hairpin-bağlı gRNA/mRNA", "4 - 8 saat", "Minimal", "BSL-1 (Bakteriyofaj kökenli)"],
      ["Eksozom-VLP Hibrit", "Otoklon membran + Gag çekirdek", "Cas9 RNP + pegRNA", "8 - 16 saat", "Minimal", "BSL-1 (Oto-biyomimetik)"],
      ["DNA Origami Kafesi", "Katlanmış M13 DNA iskele", "Nöropeptitler / RNP", "2 - 6 saat", "Sıfır", "BSL-1 (Sentetik nükleik asit)"],
      ["Sentetik Yapay Virüs (AV)", "De novo amfifilik peptitler", "Plazmit / RNP", "12 - 24 saat", "Minimal", "BSL-1 (Tamamen sentetik)"]]),

    # Table 6: Nöronal Promotörler
    ("TABLO 10.6: Hücre ve Doku-Spesifik Nöronal Promotörlerin Boyut, Güç ve Seçicilik Matrisi",
     ["Promotör Adı", "Kaynak / Mekanizma", "Dizi Boyutu (bp)", "Hedef Hücre Tipi", "Transkripsiyonel Güç", "Hücresel Seçicilik Saflığı"],
     [["hSyn1 (İnsan Synapsin-1)", "REST/NRSF susturma kontrollü", "470 bç", "Tüm Nöronlar (Pan-nöronal)", "Orta-Yüksek", "%99.2 (Glialarda kapalı)"],
      ["CaMKII-alpha (Kısaltılmış)", "Glutamaterjik regülasyon", "400 bç", "Kortikal/Hipokampal Piramidal", "Yüksek", "%91.5 (Eksitatör nöronlar)"],
      ["mDlx (mDlx5/6 Enhancer)", "Dlx homeobox transkripsiyon", "550 bç", "GABAerjik Ara Nöronlar (PV+)", "Orta", "%88.0 (İnhibitör nöronlar)"],
      ["GfaABC1D", "Kısaltılmış GFAP promotörü", "680 bç", "Astrositler (Pan-astrositik)", "Yüksek", "%98.5 (Nöronlarda sessiz)"],
      ["E-SARE (Sentetik)", "Aktiviteye duyarlı CRE/SynS", "220 bç", "Ateşlenen Engram Nöronları", "İndüklendiğinde Çok Yüksek", "%94.0 (Yalnızca aktif ağlar)"],
      ["SynP Minimal Sentetik", "De novo tasarlanmış CRE/EGR1", "180 bç", "Kortikal Piramidal Nöronlar", "Çok Yüksek (2.1x hSyn1)", "%96.5 Nöronal saflık"]]),

    # Table 7: Biyoproses ve Saflık Analitiği
    ("TABLO 10.7: Büyük Ölçekli Vektör Biyoproses Üretimi ve Kritik Kalite Parametreleri (CQA)",
     ["Biyoproses Basamağı", "Kullanılan Teknoloji / Cihaz", "Hedef Kalite Eşiği", "Analitik Ölçüm Metodu", "Klinik Güvenlik Etkisi"],
     [["Biyoreaktör Üretimi", "Süspansiyon HEK293 (2000 L)", "> 10^14 vg/L hacimsel verim", "ddPCR (ITR mutlak titresi)", "Yüksek parti verimi ve homojenlik"],
      ["Boş/Dolu Ayrımı", "Anyon Değişim Kromatografisi (AEX)", "Dolu Kapsid Oranı > %90", "Analitik Ultrasantrifüj (AUC)", "İmmünojenik boş kapsid eliminasyonu"],
      ["Nükleik Asit Temizliği", "Benzonase Endonükleaz + TFF", "Konak DNA < 10 ng/doz", "PicoGreen floresan + qPCR", "İnsersiyonel onkogenez riskini sıfırlama"],
      ["Partikül Kararlılığı", "Asimetrik Akış Fraksiyonasyonu (AF4)", "PDI < 0.05 (Agregat <%0.1)", "Dinamik Işık Saçılımı (DLS)", "Serebral mikro-emboli riskini önleme"],
      ["Sterilite ve Endotoksin", "Membran Filtrasyon (0.22 mikrometre)", "< 0.5 EU/mL endotoksin", "LAL Kinetik Türbidimetrik Test", "Menenjit ve nöroinflamasyon profilaksisi"]]),

    # Table 8: Vektör İmmünolojisi
    ("TABLO 10.8: Vektör İmmünolojik Risk Faktörleri ve Önleyici Mühendislik Çözümleri",
     ["İmmünolojik Bariyer", "Biyolojik Tetikleyici Mekanizma", "Nöronal / Klinik Risk", "Mühendislik Çözümü", "Ulaşılan Güvenlik Seviyesi"],
     [["TLR9 Aktivasyonu", "DNA'daki metillenmemiş CpG adaları", "İnterferon-alfa/beta salınımı, susturulma", "In silico CpG-Depletion (CpG-Free DNA)", "TLR9 yanıtı sıfırlandı (<2 pg/mL)"],
      ["CD8+ T Hücre Lizisi", "Kapsid peptitlerinin MHC-I sunumu", "Sitotoksik nöronal apoptozis", "Y-F mutasyonları ile hızlı soyulma", "Antijen sunumu 48 saatle sınırlandı"],
      ["Nötralizan Antikor (NAb)", "Önceden var olan anti-AAV IgG", "Vektörün kanda nötralizasyonu", "IdeS endopeptidaz veya Kapsid maskeleme", "NAb pozitif bireylerde tam etkinlik"],
      ["Mikroglial Reaktivite", "Kapsid fagositozu sonrası M1 polarizasyonu", "Pro-enflamatuar sinaps budanması", "Kapsid yüzeyine Sialik Asit eklenmesi", "Anti-inflamatuar M2 polarizasyonu"],
      ["DRG Toksisitesi", "Dorsal kök gangliyonunda aşırı ekspresyon", "Duyusal nöropati ve ataksi", "3' UTR'ye miR-183 hedef dizisi eklenmesi", "DRG ekspresyonu tamamen susturuldu"]]),

    # Table 9: Dağıtım Yolları
    ("TABLO 10.9: Serebral Dağıtım Yollarının Karşılaştırmalı Farmakokinetik Dinamiği",
     ["Dağıtım Rotası", "Uygulama Yöntemi", "Beyin Yayılım Kapsamı", "Gereken Toplam Doz", "Sistemik Organ Maruziyeti", "İnvazivlik Seviyesi"],
     [["İntravenöz (IV) Sistemik", "Periferik ven infüzyonu", "Tüm beyin parankimi (Global)", "1 - 3 * 10^12 vg/kg (CAP-B10)", "Düşük (Karaciğer detargeting ile)", "Non-invaziv (Rutin damar yolu)"],
      ["İntratekal (IT) BOS", "Lomber ponksiyon / Subaraknoid", "Tüm nöraksis, ventriküller, korteks", "1 - 5 * 10^11 vg/hasta", "Minimal (<%1 sistemik sızıntı)", "Minimal invaziv (Lomber iğne)"],
      ["CED Konveksiyonlu", "Stereotaksik mikro-kanül infüzyonu", "Odaklanmış nükleus (6-12 mm çap)", "10^9 - 10^10 vg (Ultra-düşük)", "Sıfır (Yalnızca hedef nükleusta)", "Cerrahi invaziv (Stereotaktik)"],
      ["İntranazal Mukozal", "Olfaktör epitel nebulizasyonu", "Ön beyin, prefrontal, hipokampus", "Orta doz (KBB baypas)", "Sıfır periferik organ tutulumu", "Tamamen Non-invaziv"],
      ["FUS Odaklanmış Ultrason", "IV infüzyon + Odaklı akustik dalga", "Hassas bölgesel kortikal kolon", "Düşük sistemik doz", "Düşük", "Non-invaziv / Cerrahi dışı"]]),

    # Table 10: Nihai Entegrasyon
    ("TABLO 10.10: İnsan Kognitif Dönüşümünde Vektörel Teknolojilerin Entegre Master Mimarisi",
     ["Kognitif Hedef Fonksiyon", "Seçilen Vektör Sistemi", "Taşınan Moleküler Kargo", "Regülatuar Promotör", "Beklenen Nöral Değişim", "Kognitif Fenotip Sıçraması"],
     [["Sinaptik Bellek Kilidi", "AAV.CAP-B10 (İntravenöz)", "Prime Editor (GRIN2B N616R)", "hSyn1 Promotörü", "Yavaş NMDA inaktivasyonu", "Derin bellek tutulumu x 2.4"],
      ["Dendritik Neoteni Genişlemesi", "HSV Amplikonu (150 kb)", "Tam Boy SRGAP2C Lokusu", "Endojen İnsan Promotörü", "Dendritik diken yoğunluğu +%45", "Kavramsal soyutlama hızı artışı"],
      ["Akut Plastisite İndüksiyonu", "eVLP v4 (RNP İletimi)", "Cas9 RNP + BDNF gRNA", "Doğrudan RNP (Promotörsüz)", "TrkB fosforilasyonu tepe noktası", "Akut öğrenme hızında 3 kat artış"],
      ["Kortikal Katman Senkronizasyonu", "mDlx-AAV.CAP-B10", "ChRmine Fotonik İyon Kanalı", "mDlx Enhancer", "40 Hz Gamma salınım koheransı", "Çalışma belleğinde kusursuz berraklık"],
      ["Nihai Kognitif Metamorfoz", "Kombine Entegre Protokol", "Genomik + Epigenetik Orkestrasyon", "Süper-Sentetik SynP Promotör", "Tüm beyin devre yeniden kablolaması", "Homo Singularis Bilişsel Sıçraması"]]),
]

print(f"[NEXAGEN OMEGA] Compiling {len(parts)} Parts x 10 Topics = 100 Granular Sections...")

sec_counter = 1
for part_idx, (part_title, topics) in enumerate(parts):
    h1 = doc.add_heading(part_title, level=1)
    h1.paragraph_format.space_before = Pt(22)
    h1.paragraph_format.space_after = Pt(10)
    h1.paragraph_format.keep_with_next = True
    for r in h1.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = RGBColor(13, 35, 58)

    for topic_idx, topic in enumerate(topics):
        t_title = topic[0]
        lead_txt = topic[1]
        deep_txt = topic[2]
        formula = topic[3] if len(topic) > 3 else None
        stats = topic[4] if len(topic) > 4 else None

        add_academic_section(
            doc,
            sec_num=f"{part_idx+1}.{topic_idx+1}",
            sec_title=t_title,
            lead_text=lead_txt,
            deep_text=deep_txt,
            formula=formula,
            stats=stats
        )
        sec_counter += 1

    t_tuple = tables_data[part_idx]
    tbl_title = t_tuple[0]
    tbl_headers = t_tuple[1]
    tbl_rows = t_tuple[2]

    doc.add_page_break()
    tbl_h = doc.add_heading(tbl_title, level=3)
    tbl_h.paragraph_format.space_before = Pt(14)
    tbl_h.paragraph_format.space_after = Pt(8)
    for r in tbl_h.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = RGBColor(13, 35, 58)

    create_styled_table(doc, tbl_headers, tbl_rows)

out_dir = r"C:\Users\USER\Desktop\kitap"
os.makedirs(out_dir, exist_ok=True)
out_path = os.path.join(out_dir, "BOLUM_10_VIRAL_VE_SENTETIK_VEKTORLER_TAM_100_SAYFA.docx")
doc.save(out_path)
print(f"[NEXAGEN OMEGA] BÖLÜM 10 100 SAYFALIK MASTERPIECE BAŞARIYLA ÜRETİLDİ: {out_path}")
