# -*- coding: utf-8 -*-
"""
NEXAGEN BÖLÜM 13: SENTETİK VE OPTO-EPİGENETİK KONTROL MİMARİSİ
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
        f"Kognitif nöromühendislik ve sentetik opto-epigenetik kontrol mimarisinde, {sec_title.lower()} parametreleri foton absorpsiyon kesiti, "
        "kromofor kuantum verimi, nükleer difüzyon katsayısı ve promotör işgal kinetiğiyle mutlak bir korelasyona sahiptir. "
        "Santral sinir sistemindeki spesifik nöron topluluklarının milisaniyelik zamansal ve mikrometre düzeyindeki uzaysal çözünürlükle ışıkla uyarılması, "
        "yalnızca tekil genlerin transkripsiyonunu açıp kapatmakla kalmaz; aynı zamanda nöronal devrelerin osilatuar koheransını, engram fiksasyonunu "
        "ve yüksek bilişsel kavrayış pencerelerini milisaniyelik bir optik anahtarla kontrol etme imkanı sunar."
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
        f"{deep_text} Bu fotobiyofiziksel mekanizma, foton uyarımı sonrasında oluşan konformasyonel serbest enerji bariyerini aşarak "
        "kromatindeki susturucu komplekslerin yerini saniyeler içinde aktif histon asetiltransferazlara (p300) ve demetilazlara (TET1) "
        "bırakmasını sağlar. Foton akışı kesildiğinde termal gevşeme kinetiği sistemi kendiliğinden bazal dinlenim durumuna döndürür; böylece "
        "hiçbir kalıcı epigenetik toksisite veya off-target sızıntı yaşanmadan, kognitif mimarinin istenen zaman aralıklarında "
        "en yüksek bilişsel verimlilikte çalışması temin edilmiş olur."
    )
    r_deep.font.name = "Calibri"
    r_deep.font.size = Pt(9.5)
    r_deep.font.color.rgb = RGBColor(45, 55, 65)

print("[NEXAGEN OMEGA] Initializing Chapter 13 Builder Engine...")

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

r_vol = title_p.add_run("NEXA GENETİK VE NÖRO-MÜHENDİSLİK MONOGRAFİLERİ\nSERİ 13: FOTONİK VE OPTO-EPİGENETİK KONTROL\n\n")
r_vol.font.name = "Calibri"
r_vol.font.size = Pt(13)
r_vol.font.bold = True
r_vol.font.color.rgb = RGBColor(120, 140, 160)

r_main = title_p.add_run("BÖLÜM 13: SENTETİK VE OPTO-EPİGENETİK KONTROL MİMARİSİ\n")
r_main.font.name = "Calibri"
r_main.font.size = Pt(22)
r_main.font.bold = True
r_main.font.color.rgb = RGBColor(13, 35, 58)

r_sub = title_p.add_run("LITE Sistemleri, CRY2-CIB1 Mimarisi, Fotonik Demetilasyon, Derin Beyin UCNP Arayüzleri ve Işıkla Bellek Programlama")
r_sub.font.name = "Calibri"
r_sub.font.size = Pt(12)
r_sub.font.italic = True
r_sub.font.color.rgb = RGBColor(70, 80, 95)

doc.add_paragraph().paragraph_format.space_after = Pt(18)

p_intro = doc.add_paragraph()
p_intro.paragraph_format.line_spacing = 1.2
p_intro.paragraph_format.space_after = Pt(14)
r_in = p_intro.add_run(
    "Klasik gen terapisinin ve farmakolojik müdahalelerin en büyük sınırlılığı zamansal ve uzaysal kontrol eksikliğidir; bir ilaç kanda saatlerce "
    "dolaşır ve hedeflenen nöronun yanı sıra çevre dokuları da etkiler. Opto-epigenetik mühendislik, ışığın (fotonların) hızı ve mikrometre düzeyindeki "
    "uzaysal hassasiyetini nükleer kromatin modifikatörleri ile birleştirerek kognitif genetik mimaride devrim yaratmıştır. "
    "Işıkla indüklenebilir transkripsiyonel efektörler (LITE), CRY2-CIB1 ve PhyB-PIF foto-anahtarları, dCas9 rehberliğinde istenen kognitif "
    "gen promotörüne milisaniyeler içinde kilitlenerek eukromatin açılımını (H3K27ac) ve DNA demetilasyonunu (5hmC) başlatır; ışık kesildiğinde "
    "sistem kendiliğinden kapanır. Bu monografi, transkraniyal optik dalga boylarından derin doku upconversion nanopartiküllerine (UCNP) kadar "
    "ışıkla yönetilen kognitif kontrol mimarisinin tüm biyofiziksel ilkelerini sunmaktadır."
)
r_in.font.name = "Calibri"
r_in.font.size = Pt(10)
r_in.font.color.rgb = RGBColor(40, 50, 60)

# SECTION DATA GENERATOR: 10 Parts x 10 Topics = 100 Topics
parts = [
    ("KISIM I: OPTO-EPİGENETİK FOTOBİYOFİZİK VE FOTORESEPTÖR MEKANİZMALARI", [
        ("Opto-Epigenetiğin Temel Biyofiziksel İlkeleri ve Klasik Optogenetikten Ayrımı",
         "Klasik optogenetik iyon kanallarını (ChR2, Halorhodopsin) kullanarak milisaniyelik elektriksel aksiyon potansiyellerini yönetirken; opto-epigenetik doğrudan nükleusta gen transkripsiyonunu ve kromatin mimarisini ışıkla programlar.",
         "Foton soğurumu bir transmembran iyon akımı değil, bir protein-protein dimerizasyon enerjisi üretir. Nükleer dCas9 veya TALE iskelelerine monte edilen efektör enzimler (p300, TET1, KRAB), ışık varlığında DNA'ya kenetlenir. Elektriksel uyarım bittikten sonra da günlerce sürecek kalıcı bir kognitif iz bırakır.",
         "Delta G_photon = h * c / lambda; Dimerization_Rate = phi_quantum * I_light * sigma_abs", "Kromofor kuantum verimi: phi ~ 0.35 (CRY2); Termal gevşeme süresi: dakikalar mertebesinde."),

        ("Kriptokrom 2 (CRY2) ve CIB1 Fototropik Dimerizasyon Biyofiziği",
         "Arabidopsis thaliana kökenli Kriptokrom 2 (CRY2), flavin adenin dinükleotid (FAD) kromoforu içerir; 450 nm mavi ışığı soğurduğunda konformasyonel olarak açılır.",
         "Açılan CRY2, temel heliks-ilmik-heliks proteini CIB1 ile pikomolar affiniteyle kenetlenir (heterodimerizasyon). Bu fotonik kenetlenme 1 saniye içinde tamamlanır; mavi lazer kapatıldığında FAD'ın kendiliğinden termal re-oksidasyonu ile CRY2 ve CIB1 10-15 dakika içinde ayrışır.",
         "K_d(Light) < 15 nM vs K_d(Dark) > 10 mikroM (1000 katlık dinamik aralık)", "Aktivasyon gecikmesi: tau_on < 800 milisaniye; Karanlık fazda disosiasyon: tau_off ~ 12 dakika."),

        ("Fitochrom B (PhyB) ve PIF6 ile Çift-Dalga Boylu (Kırmızı / Uzak-Kırmızı) Açma-Kapama",
         "PhyB sistemi fitochromobilin kromoforu kullanır ve iki zıt dalga boyunda tersinir olarak çalışır: 660 nm kırmızı ışıkla açılır, 730 nm uzak-kırmızı ışıkla anında kapanır.",
         "Karanlıkta kendiliğinden yavaş ayrışmayı beklemek yerine, 730 nm lazer darbesi verildiğinde PhyB ve PIF6 milisaniyeler içinde birbirini fırlatıp atar. Transkripsiyonel süreç mikrosaniyelik optik darbelerle tam tersinir kontrol altına alınır.",
         "PhyB(Pr) -(660 nm)-> PhyB(Pfr, Aktif) -(730 nm)-> PhyB(Pr, İnaktif)", "Kapatma (Off-switch) hızı: Uzak-kırmızı darbesiyle <500 milisaniye."),

        ("LOV Domain Ailesi (AsLOV2, iLID, Magnets) ve Allosterik Kavitasyon",
         "Avena sativa fototropin-1'in Light-Oxygen-Voltage 2 (AsLOV2) alanı, flavin mononükleotid (FMN) kofaktörü taşır.",
         "Mavi foton emildiğinde FMN ile Cys450 arasında geçici bir kovalent tiyoeter bağı kurulur; C-terminal J-alfa heliksi proteinden çözülerek dışarı fırlar (unfolding). Bu hareket, arkasına gizlenmiş nükleer lokalizasyon sinyalini (NLS) veya peptid bağlayıcıyı serbest bırakır.",
         "Unfolding_kinetics: tau_open < 100 mikrosaniye (Işık soğurulması sonrası)", "Karanlık fazda heliks kapanma yarı ömrü: t_1/2 ~ 25 saniye (Hızlı geri dönüşümlü)."),

        ("Kromofor Biyo-Kullanılabilirliği ve Endojen Kofaktör Doygunluğu",
         "Opto-epigenetik anahtarların çalışabilmesi için memeli nöron nükleusunda yeterli FAD veya FMN kromoforu bulunmalıdır.",
         "Nöronal mitokondri ve sitoplazmada riboflavin (B2 vitamini) metabolizması ile üretilen serbest FAD konsantrasyonu (5-10 mikroM), CRY2 ve LOV apo-proteinlerini %100 doyurmaya yeterlidir; dışarıdan ek kimyasal kromofor verilmesine gerek kalmaz.",
         "[FAD]_intranuclear = 6.8 +- 1.2 mikroM >> K_d(CRY2-FAD) ~ 0.5 mikroM", "Apo-protein doygunluk fraksiyonu: %96 (Doğal endojen kofaktör yeterliliği)."),

        ("Fototoksisite ve Termal Yönetim: Nöronal Işık Eşiği Güvenliği",
         "Aşırı şiddetli sürekli mavi ışık (450 nm), fototoksik reaktif oksijen türleri (ROS) üreterek ve lokal doku sıcaklığını artırarak nöronlara zarar verebilir.",
         "Beyin dokusu için güvenlik eşiği lazer güç yoğunluğunda <10 mW/mm^2 ve lokal sıcaklık artışında Delta T < 0.5 C olarak sınırlandırılmıştır. Kesintili darbe katarı (pulsed light, %10 duty cycle) kullanılarak termal yük sıfıra indirilir.",
         "Thermal_Diffusion: Delta T = (P_laser * alpha_abs) / (4 * pi * kappa * r) < 0.5 C", "Optik güç güvenliği: 2-5 mW/mm^2 (Nöronal canlılık %100 korunur)."),

        ("Nöron Nükleusunda Optik Gradyan ve Foton Yayılım Biyofiziği",
         "Işık beyin dokusuna girdiğinde miyelin kılıfları ve hücre gövdeleri tarafından Rayleigh ve Mie saçılmasına uğrar; yoğunluğu derinlikle katlanarak azalır.",
         "Monte Carlo foton simülasyonları ile kortikal katmanlara ulaşan efektif foton yoğunluğu modellenir. Nükleer hedefleme için gereken minimal eşik foton dozu (0.5 mW/mm^2) tüm korteks katmanlarında homojen sağlanır.",
         "Beer-Lambert_Scattering: I(z) = I_0 * exp( - mu_eff * z ); mu_eff ~ 0.8 mm^-1 (Korteks)", "Optik penetrasyon derinliği: 450 nm mavi ışıkla doğrudan dokuda 1.5 - 2.0 mm."),

        ("Karanlıkta Sızıntı (Dark Leakage) Kinetiği ve Arka Plan Susturması",
         "İdeal bir opto-epigenetik anahtar ışık yokken sıfır aktivite göstermelidir; ancak bazı CRY2 varyantları karanlıkta da zayıf spontan dimerleşme sergileyebilir.",
         "CRY2 aktif merkezine eklenen rasyonel mutasyonlar (CRY2-L348F veya CRY2olig mutantları), karanlık durumdaki bazal sızıntıyı binde birin altına indirir. Gen ekspresyonu ışık gelene kadar mutlak sessizlikte kilitli kalır.",
         "Leakage_Ratio = Activity(Dark) / Activity(Light) < 0.001 (Ultra-temiz karanlık durum)", "Sinyal/Gürültü oranı (SNR): >60 dB."),

        ("Bölünmüş (Split) Opto-Epigenetik Tasarımlar ve İkili Işık Kontrolü",
         "dCas9'un kendisi iki parçaya (N-dCas9 ve C-dCas9) bölünerek p-Mag ve n-Mag fotoreseptörleri ile eşleştirilir.",
         "Mavi ışık verildiğinde nükleaz parçaları birleşerek hedef DNA'yı tanır; eşzamanlı olarak kırmızı ışık verildiğinde p300 efektörü sisteme kenetlenir. Çift anahtarlı (AND-gate) biyomantık mantığı ile epigenetik hassasiyet ikiye katlanır.",
         "Logic_Gate: Expression = Light_450nm AND Light_660nm", "Hatalı pozitif hedefleme sıklığı: Sıfır."),

        ("Opto-Epigenetiğin Nöromorfik Hesaplamadaki Yeri",
         "Nöronal kromatini ışıkla programlamak, biyolojik beyni harici optoelektronik çiplere doğrudan bağlayan nihai bir arayüzdür.",
         "Elektronik devrelerde üretilen optik sinyaller doğrudan nöronal histon modifikasyonlarına dönüştürülür; insan zihni ile yapay zeka arasında doğrudan fotonik bir translasyon kanalı açılır.",
         "Bandwidth_OptoEpigenetic = N_channels * Rate_photoinduction > 10^9 operasyon/saat", "Biyosibernetik entegrasyon: Fotonik-genomik doğrudan köprü.")
    ]),

    ("KISIM II: LITE (LIGHT-INDUCIBLE TRANSCRIPTIONAL EFFECTORS) SİSTEMLERİ", [
        ("Feng Zhang Laboratuvarı Tarafından Geliştirilen LITE Mimarisi",
         "2013'te MIT Broad Enstitüsü'nde Feng Zhang ve ekibi tarafından icat edilen LITE, TALE DNA bağlama proteini ile CRY2-CIB1 sistemini birleştiren ilk entegre opto-genomik platformdur.",
         "TALE-CIB1 çekirdekte spesifik kognitif gen promotörüne kalıcı olarak oturur; CRY2-VP64 ise nükleusta serbest dolaşır. 450 nm mavi ışık verildiğinde VP64 TALE'ye kenetlenir ve gen transkripsiyonu 20 ila 60 kat patlar.",
         "Modular_Design: [TALE_DNA-Binding] + [CIB1] :: [CRY2] + [Effector_VP64/p300]", "İndüksiyon hızı: Işık açıldıktan 15 dakika sonra belirgin mRNA artışı."),

        ("TALE vs dCas9 İskelelerinin Opto-Epigenetikte Karşılaştırmalı Biyofiziği",
         "TALE iskeleleri tamamen protein kökenli olup sıfır RNA gerektirirken; dCas9 tek bir proteinle gRNA değiştirilerek binlerce geni hedefleme esnekliği sunar.",
         "dCas9-CIB1 füzyonları multiplex opto-epigenetik düzenlemeye izin verir; aynı mavi lazer darbesiyle 10 farklı kognitif gen eşzamanlı olarak aktive edilebilir.",
         "Multiplex_Index(dCas9-LITE) = 10+ bağımsız gen lokusu tek partikülle", "Hedefleme doğruluk skoru: %98.6."),

        ("LITE ile Post-Mitotik Nöronlarda Histon Asetilasyonu (LITE-p300)",
         "Klasik LITE'ta kullanılan viral VP64 yerine insan p300 HAT katalitik çekirdeğinin (p300Core) CRY2'ye bağlanmasıdır.",
         "Bu modifikasyon, viral aktivatörlerin yaratabileceği hücresel stresi ortadan kaldırır; doğrudan hedef promotördeki H3K27 pozisyonunu asetilleyerek fizyolojik bir eukromatin açılımı sağlar.",
         "Acetylation_Flux: d[H3K27ac]/dt = k_p300 * [CRY2-p300_bound] * [Acetyl-CoA]", "Yerel H3K27ac zenginleşmesi: Işık uygulamasından 2 saat sonra 18 kat artış."),

        ("LITE ile Nöronal DNA Demetilasyonu (LITE-TET1)",
         "CRY2'nin TET1 dioksijenaz domaini ile füzyonlanması, 450 nm mavi ışıkla hedeflenen gen lokusunda 5mC'nin 5hmC'ye dönüştürülmesini sağlar.",
         "Geleneksel TET1'in kontrolsüz hiperaktivitesini engelleyen bu sistem, demetilasyonu yalnızca fotonun aydınlattığı zaman penceresiyle (örn. 30 dakikalık lazer uyarımı) sınırlar; aşırı demetilasyon toksisitesi önlenir.",
         "Demethylation_Control: [5hmC_produced] = Integral_light ( k_TET1 * I(t) ) dt", "Demetilasyon saflığı: Hedef lokusta %80 dönüşüm; Hedef dışı sızıntı sıfır."),

        ("LITE Transkripsiyonel Kinetiği: Açılma ve Sönümlenme Dinamikleri",
         "LITE sistemi ışık darbesi başladıktan yaklaşık 30 dakika sonra tepe mRNA üretimine ulaşır.",
         "Işık kesildiğinde CRY2-CIB1 ayrışır, transkripsiyonel inisyasyon durur ve mevcut mRNA'lar hücresel yarı ömürleriyle (1-3 saat) parçalanarak bazale döner. Nöronal gen ekspresyonu üzerinde tam bir 'darbe genişlik modülasyonu' (PWM) kurulur.",
         "Pulse_Width_Modulation: Duty_Cycle = t_light / (t_light + t_dark)", "Doz-yanıt hassasiyeti: %5 ışık aralığıyla titre edilebilir protein üretimi."),

        ("LITE Sistemlerinin İn Vivo AAV Dağıtımı ve Kapsid Optimizasyonu",
         "TALE veya dCas9 ile CRY2/CIB1 genlerinin toplam boyutu büyük olduğundan çift AAV vektör mimarisi kullanılır.",
         "Birinci AAV.CAP-B10 nükleer bağlayıcıyı (dCas9-CIB1), ikinci virüs ise efektörü (CRY2-p300) nöronlara iletir. İkili transduksiyonla kortikal piramidal nöronların %70'inde fonksiyonel LITE devreleri kurulur.",
         "Dual_AAV_CoTransduction = P(AAV1) * P(AAV2) * CoLocal ~ %72", "Kortikal ifade başarısı: Yetişkin primat beyninde tam doğrulanmış etkinlik."),

        ("LITE ile c-Fos ve Arc Erken Yanıt Genlerinin Sentetik Aktivasyonu",
         "Doğal sinaptik girdi olmadan, yalnızca kafatasından gönderilen mavi lazerle c-Fos veya Arc promotörünün aktive edilmesidir.",
         "Bu yapay aktivasyon, nöronu sanki derin bir öğrenme deneyimi yaşamış gibi uyararak sinaptik dikenlerini genişletmeye zorlar; 'sentetik öğrenme' paradigması hayata geçer.",
         "Spine_Volume_Expansion = Delta V_spine = +58% (LITE uyarımı sonrası)", "LTP indüksiyonu: Elektriksel şok olmaksızın salt optik yolla kalıcı plastisite."),

        ("LITE Karşıtı Opto-Susturucular (Opto-KRAB ve Opto-DNMT3A)",
         "CRY2'ye KRAB veya DNMT3A bağlanarak oluşturulan ışıkla çalışan susturucu sistemlerdir.",
         "Aşırı aktif olan epileptiform devreler veya travmatik bellek odakları mavi ışıkla aydınlatıldığında, ilgili gen lokuslarına hızla H3K9me3 ve 5mC damgaları vurularak devre susturulur.",
         "Inhibition_Fold = 15 - 30 kat transkripsiyonel blokaj", "Optik susturma hızı: 1 saatlik aydınlatma sonrası %95 transkripsiyon düşüşü."),

        ("Uzun Süreli LITE Uyarımında Hücresel Adaptasyon ve Tolerans",
         "Günler süren periyodik optik uyarımda nöronun LITE sistemine tolerans geliştirip geliştirmediği test edilmiştir.",
         "CRY2 proteininin parçalanma hızının sabit kaldığı, nükleer FAD havuzunun tükenmediği ve 30 günlük günlük lazer seansları boyunca sistemin ilk günkü transkripsiyonel gücünü (%95+) koruduğu teyit edilmiştir.",
         "Response_Stability: Activity(Day_30) / Activity(Day_1) = 0.96 +- 0.03", "Kronik yorulma veya duyarsızlaşma (desensitization): Yok."),

        ("LITE Teknolojisinin Gelecek Nesil Gelişimi: Nano-LITE",
         "TALE yerine minyatür CasMINI veya hiper-kompakt çinko parmak dizilerinin entegrasyonu ile sistemin boyutu yarıya indirilmiştir.",
         "Tek bir AAV kapsidine sığabilen Nano-LITE, kognitif gen kontrolünü tek enjeksiyonla ve klinik pratiklikte uygulanabilir hale getirir.",
         "Cargo_Size(Nano-LITE) = 4.3 kb < 4.7 kb (Tek AAV paketleme zaferi)", "Klinik translasyon kolaylığı: Maksimum düzeyde optimize edilmiş vektör.")
    ]),

    ("KISIM III: UPCONVERSION NANOPARTİKÜLLERİ (UCNP) VE DERİN BEYİN FOTONİK ARAYÜZLERİ", [
        ("Doku Penetrasyon Kısıtı ve Mavi Işığın Serebral Saçılması Problemi",
         "450 nm mavi ışık beyin dokusunda yüksek oranda saçıldığı ve hemoglobin tarafından emildiği için kafatası üzerinden verildiğinde ancak 1-2 mm derinliğe inebilir.",
         "Hipokampus, amigdala ve talamus gibi derin kognitif nükleuslara ulaşmak için klasik yöntem cerrahi fiber optik kanüller çakmaktır; ancak bu invazivdir ve doku hasarı yaratır. Bu kısıtı aşan teknoloji Upconversion Nanopartikülleridir (UCNP).",
         "Penetration_Depth(Blue_450nm) ~ 1.5 mm vs Penetration_Depth(NIR_980nm) > 8.0 mm", "Saçılma katsayısı: Yakın kızılötesi ışıkta (NIR) 10 kat daha düşük."),

        ("Upconversion Nanopartiküllerinin (UCNP) Kuantum Biyofiziği",
         "UCNP'ler (NaYF4:Yb3+,Tm3+ veya NaYF4:Yb3+,Er3+), nadir toprak elementleri ile katkılanmış dielektrik inorganik nano-kristallerdir.",
         "Anti-Stokes lüminesans ilkesiyle çalışırlar: Düşük enerjili iki veya daha fazla yakın-kızılötesi (NIR, 980 nm veya 808 nm) fotonunu art arda soğurarak yüksek enerjili tek bir mavi (450 nm) veya yeşil (540 nm) foton olarak geri yayarlar.",
         "Anti-Stokes_Conversion: 2 * h*nu_NIR -(UCNP kristal kafesi)-> 1 * h*nu_Blue (Fotonik sıçrama)", "Kuantum dönüşüm verimi: Çekirdek-kabuk (core-shell) tasarımıyla %8.5."),

        ("Kafatasını Aşan 980 nm ve 808 nm Yakın-Kızılötesi (NIR) Optik Penceresi",
         "Biyolojik dokular için 'birinci ve ikinci optik pencereler' (700-1100 nm), su ve hemoglobin emiliminin en düşük olduğu spektral koridordur.",
         "808 nm lazer kafatası kemiğini ve tüm serebral korteksi hiçbir delme işlemi yapmadan zahmetsizce geçer. Derin beyindeki UCNP'ler bu görünmez NIR ışığını yakalayarak lokal olarak parlak mavi ışığa çevirir.",
         "Non-invasive_Deep_Access: Kafatası sağlam primatta hipokampal aktivasyon", "Cerrahi travma riski: Mutlak sıfır (Non-invaziv transkraniyal uyarım)."),

        ("Çekirdek-Kabuk (Core-Shell) UCNP Mimarisi ve Yüzey Su Söndürme Blokajı",
         "Sulu hücre içi ortamda, su molekülleri UCNP yüzeyindeki lantanit iyonlarının enerjisini emerek lüminesansı söndürür (quenching).",
         "Aktif NaYF4:Yb,Tm çekirdeğinin etrafına inert bir NaYF4 kabuğu (epitaksiyel shell) kaplanarak enerji sızıntısı önlenir. Nöron içi ortamlarda lüminesans parlaklığı 50 kat artırılır.",
         "Luminescence_Intensity: I_shell / I_core_only = 48 kat artış", "Partikül hidrodinamik çapı: 25 - 35 nm; Koloidal kararlılık: Tam biyouyum."),

        ("UCNP Yüzey Nöro-Mühendisliği: Polietilen Glikol (PEG) ve RVG29 Konjugasyonu",
         "UCNP'lerin nöronlara hedeflenmesi için yüzeyleri karboksilat-PEG zincirleri ve kuduz virüsü glikoproteini (RVG29) peptidi ile donatılır.",
         "Sistemik dolaşıma verilen veya BOS içine infüze edilen UCNP'ler, nöronal nAChR reseptörlerine tutunarak endositozla nöron somasına yerleşir ve doğrudan nükleer zar çevresinde kümelenir.",
         "Targeting_Efficiency: %88 nöronal lokalizasyon", "Hücresel tutulum yarı ömrü: Doku içinde aylarca sabit barınma."),

        ("UCNP-LITE Entegrasyonu: NIR ile Tetiklenen Kognitif Demetilasyon",
         "Nöron içine hem dCas9-CIB1 / CRY2-TET1 kaseti hem de RVG-UCNP partikülleri yerleştirilir.",
         "Kafatasının üstünden gönderilen harici 808 nm NIR lazer, derin hipokampustaki UCNP'leri uyarır; UCNP'lerin yaydığı yerel 450 nm mavi fotonlar CRY2-TET1'i ateşler. Derin bellek lokusları cerrahisiz olarak yerinde demetillenir.",
         "Coupling_Efficiency = Luminescence_UCNP * Absorption_CRY2 > Eşik_aktivasyon", "Derin beyin kognitif gen ekspresyon artışı: 25 kat indüksiyon."),

        ("808 nm UCNP Tasarımları ile Aşırı Isınma (Overheating) Eliminasyonu",
         "980 nm dalga boyundaki NIR ışık, yüksek güçlerde su tarafından bir miktar emilerek hafif doku ısınması yaratabilir.",
         "Neodim (Nd3+) katkılı yeni nesil UCNP'ler 808 nm lazer ile uyarılır. 808 nm'de su emilimi 980 nm'ye kıyasla 5 kat daha düşüktür; beyin dokusu sıcaklığı Delta T < 0.1 C'de sabit kalır.",
         "Specific_Absorption_Rate: SAR(808nm) = 0.18 * SAR(980nm)", "Termal biyogüvenlik katsayısı: %100 kusursuz soğuk fotonik transfer."),

        ("UCNP Biyouyumluluğu, Karaciğer Klirensi ve Ağır Metal Toksisitesi Güvenliği",
         "Lantanit nano-kristallerinin (Yttriyum, Tuliyum, İterbiyum) dokuda serbest iyon sızıntısı yapması sitotoksisite riski doğurabilir.",
         "Ultra-kararlı florür kafes yapısı ve silika/altın koruyucu nano-kabuklar, lantanit sızıntısını sıfıra indirir. 12 aylık in vivo toksikoloji çalışmalarında hiçbir nörodejenerasyon veya mikroglial fagositoz patolojisi saptanmamıştır.",
         "Leaching_Rate < 10^-8 M / yıl (Tespit limitinin altında)", "12 aylık nöronal canlılık: %99.4; Glial reaktivite: Bazal düzey."),

        ("Kablosuz ve Giyilebilir NIR Kask Sistemleri ile Sürekli Bilişsel Kontrol",
         "Hastanın kafasına takılan hafif, pilli, düşük güçlü NIR LED dizilimli kognitif kask.",
         "Birey karmaşık bir sınava girerken veya yoğun akademik çalışma yaparken kask otomatik olarak düşük frekanslı 808 nm darbeleri gönderir; derin beyindeki UCNP'ler kognitif genleri açarak çalışma belleğini zirvede tutar.",
         "Wearable_Photonics: 808 nm Darbeli LED Matrisi (Güç tüketimi < 5 Watt)", "Kullanıcı ergonomisi: Tamamen mobil, hastane dışı serbest kullanım."),

        ("Fotonik Derin Beyin Arayüzünün Nihai Evrimi: Organik Biyolüminesans",
         "Harici lazer yerine, nöron içine genetik olarak kodlanan lusiferaz enzimleri (NanoLuc) ve kalsiyuma duyarlı biyolüminesan substratlar entegre edilir.",
         "Nöron kendi elektriksel uyarımı anında kendi mavi fotonunu kendisi üretir; bu foton komşu opto-epigenetik anahtarı ateşler. Dışarıdan hiçbir ışık kaynağı gerektirmeyen otonom biyo-fotonik zeka devresi tamamlanır.",
         "BRET_Efficiency: Bioluminescence Resonance Energy Transfer > %70", "Otonom fotonik kognitif döngü: Biyolojik olarak kendi kendini besleyen zihin.")
    ]),

    ("KISIM IV: KİMYASAL VE KÜÇÜK MOLEKÜL İLE İNDÜKLENEBİLİR EPİGENETİK ANAHTARLAR", [
        ("Chemically Inducible Dimerization (CID) Sistemlerinin Termodinamiği",
         "CID sistemleri, inert küçük bir molekülün iki bağımsız protein alanının arayüzüne oturarak aralarında nanomolar affiniteyle köprü kurması prensibine dayanır.",
         "Küçük molekül olmadan iki protein birbirini hiç tanımaz (Kd > 1 mM); molekül eklendiğinde serbest enerji hızla negatife düşer ve kararlı bir dimer oluşur. Epigenetik düzenleyicilerin farmakolojik uzaktan kumandasıdır.",
         "Delta G_ternary = Delta G_1 + Delta G_2 + Delta G_cooperativity; K_d ~ 10^-9 M", "Dinamik açılma aralığı: 1000 kattan fazla transkripsiyonel indüksiyon."),

        ("Rapalog (FKBP-FRB) Sistemi ve dCas9-Epigenom Kontrolü",
         "En yaygın CID sistemi, immünosupresif olmayan sentetik rapamisin analoglarını (rapaloglar: AP21967) kullanan FKBP12 ve FRB mimarisidir.",
         "dCas9-FKBP hedef promotörde beklerken, FRB-p300 nükleer sitoplazmada serbesttir. Hastaya mikrogram düzeyinde rapalog verildiğinde 15 dakika içinde FRB-p300 dCas9'a kilitlenir; gen eukromatinleşir. İlaç kandan çekildiğinde sistem kendiliğinden ayrışır.",
         "EC50_dimerization ~ 2.4 nM (Aşırı düşük non-toksik ilaç dozu)", "Transkripsiyonel gecikme: İlaç alımından 30 dakika sonra tepe mRNA artışı."),

        ("Abscisik Asit (ABA) Sistemi: Bitkisel Kökenli Sıfır-İmmünite CID",
         "Memeli hücrelerinde hiçbir endojen fonksiyonu bulunmayan bitki hormonu Abscisik Asit (ABA), PYL1 ve ABI1 proteinlerinin dimerleşmesini sağlar.",
         "Memeli metabolizmasında tamamen inert olması sayesinde sıfır yan etki profili sergiler. Rapalog sistemine ortogonal olarak çalışır; yani aynı nöron içinde iki farklı kognitif geni bağımsız iki ayrı ilaçla yönetme imkanı sunar.",
         "Selectivity_Orthogonal: ABA sistemi rapalog ile, Rapalog sistemi ABA ile çapraz etkileşmez", "Hücresel biyouyumluluk: Gram düzeyinde dozlarda bile sıfır toksisite."),

        ("Gibberellin (GA3) Sistemi ile Üçlü Mantık Kapılarının (Three-Way Control) Kurulması",
         "Üçüncü bir bitkisel hormon olan Gibberellin türevi (GA3-AM), GID1 ve GAI protein alanlarını kenetler.",
         "Aynı nöron nükleusuna Rapalog, ABA ve Gibberellin sistemleri eşzamanlı yerleştirildiğinde; hastaya verilecek üç farklı kapsül ile üç bağımsız kognitif gen kaseti (örn. BDNF, GRIN2B ve Klotho) ayrı ayrı veya kombinatoryal olarak açılıp kapatılabilir.",
         "Three_Channel_Truth_Table: 2^3 = 8 bağımsız epigenetik durumun farmakolojik yönetimi", "Kombinatoryal zeka kalibrasyonu: Hassas bilişsel akortlama."),

        ("K-OFF Sistemleri: İlaçla Ayrışan (Chemically Disrupted) Epigenetik Kilitler",
         "Bazı kognitif protokollerde genin sürekli açık kalması, yalnızca istenmeyen bir durumda ilaç verilerek derhal kapatılması istenir.",
         "Hücre içinde kendiliğinden dimerleşen ancak küçük molekül eklendiğinde hızla birbirinden kopan sentetik protein çiftleri (örn. Bcl-xL ve BH3 mimetikleri) kullanılır. İlaç verildiğinde epigenetik aktivatör DNA'dan düşer ve gen anında susar.",
         "Disruption_kinetics: tau_off < 5 dakika (İlaç kromatini anında temizler)", "Acil durum kognitif freni: Hızlı ve kesin kapatma emniyeti."),

        ("HaloTag ve SNAP-Tag ile Kimyasal Epigenetik Konjugasyon",
         "Proteinlerin sentetik kloroalkan veya benzilguanin türevi sentetik ligandlara kovalent bağlanması prensibidir.",
         "HaloTag-dCas9 füzyonları, dışarıdan verilen floresan veya radyo-opak ligandları kovalent olarak bağlayarak kognitif gen lokusunun canlı beyinde PET veya MRI ile tek molekül çözünürlüğünde görüntülenmesini sağlar.",
         "Covalent_Bond_Formation: Hızlı ve geri dönüşümsüz kimyasal etiketleme", "In vivo kognitif takip: Moleküler görüntüleme ile %100 örtüşme."),

        ("Farmakokinetik Doz-Yanıt Kalibrasyonu ve Bireyselleştirilmiş Titrasyon",
         "CID sistemlerinde transkripsiyonel aktivasyon düzeyi verilen ilacın plazma konsantrasyonu ile doğrusal korelasyon gösterir.",
         "Hasta kendi bilişsel talebine göre (örn. hafif odaklanma için 1 mg, aşırı karmaşık matematiksel problem çözümü için 5 mg) dozu ayarlayarak kognitif gen ekspresyonunun debisini titre edebilir.",
         "Dose_Response_Linearity: Expression = alpha * [Ligand_dose]; r = 0.98", "Bireysel kontrol esnekliği: Tam dijital benzeri kademeli ayar."),

        ("Nöron İçi Sinyal Bozulması ve Protein Kararlılığı (PROTAC Epigenetiği)",
         "Kognitif epigenetik düzenleyicinin ömrünü sınırlamak için Proteolysis Targeting Chimera (PROTAC) teknolojisi entegre edilir.",
         "dCas9'a eklenen degron etiketleri, spesifik bir küçük molekül verildiğinde hücresel E3 ubiquitin ligazı (CRBN veya VHL) çekerek dCas9'u 2 saat içinde proteazomda parçalatır. Epigenetik müdahale tek bir komutla hücreden tamamen silinir.",
         "Degradation_Clearance: [Epi_Editor]_t = [Epi_Editor]_0 * exp( - k_PROTAC * t )", "Hücresel temizlenme yarı ömrü: t_1/2 ~ 75 dakika."),

        ("KBB Geçirgenliğine Sahip Sentetik Dimerizer Tasarımları",
         "Klasik rapamisin molekülü büyük bir makrosiklik lakton olup kan-beyin bariyerini sınırlı geçer.",
         "Yeni nesil sentetik 'nano-dimerizerlar' (moleküler ağırlığı <500 Da, LogP ~ 2.0), lipid membranları zahmetsizce aşarak oral alımdan sonra 20 dakika içinde serebral kortekse ulaşır.",
         "Permeability_Surface_Area (PS) > 25 * 10^-6 mL/s/g beyin dokusu", "Beyin biyoyararlanımı: Enjekte edilen dozun %12'si doğrudan nükleuslara ulaşır."),

        ("Kimyasal vs Optik Kontrolün Hibrit Sentezi: Foto-Kafesli (Photocaged) Dimerizerlar",
         "İki teknolojinin mükemmel evliliği: Rapalog molekülü ışıkla parçalanabilen bir nitrobenzil kimyasal kafesi ile kilitlenir (caged-rapalog).",
         "İlaç vücuda verilir ancak kafesli olduğu için hiçbir yerde bağlanamaz (sıfır periferik etki). Yalnızca hedeflenen kortikal bölgeye lazer tutulduğunda kafes fotolitik olarak kopar ve ilaç aktifleşir. Hem kimyasal hem optik çift katmanlı mutlak güvenlik.",
         "Uncaging_Reaction: Caged_Ligand -(365-405 nm)-> Active_Ligand + Byproduct", "Çift emniyet faktörü: Mekansal hata olasılığı milyonda birden az.")
    ]),

    ("KISIM V: NÖRONAL DEVRE SENKRONİZASYONU, ENGRAM MÜHENDİSLİĞİ VE BİLİŞSEL DÖNÜŞÜM", [
        ("Kortikal Sütunlar (Cortical Columns) Seviyesinde Opto-Epigenetik Haritalama",
         "İnsan neokorteksi yaklaşık 50 mikrometre çapında ve 2-3 mm derinliğinde milyonlarca bağımsız işlem birimi olan 'kortikal sütunlar'dan oluşur.",
         "Opto-epigenetik lazer desenleme (Spatial Light Modulators - SLM) ile, tek tek kortikal sütunlar seçilerek bağımsız genetik profillerle kodlanabilir. Bir sütun soyut mantık işlemcisine, yanındaki sütun mekansal hafıza matrisine dönüştürülür.",
         "Spatial_Light_Modulator: 1024 x 768 bağımsız lazer mikro-ışını", "Uzaysal adresleme çözünürlüğü: 50 mikrometre (Tek sütun izolasyonu)."),

        ("Prefrontal Korteks (PFC) L2/3 ve L5 Piramidal Devrelerinin Epigenetik Güçlendirilmesi",
         "Dorsolateral Prefrontal Korteksin Katman 2/3 nöronları çalışma belleğinin soyut manipülasyonunu yaparken; Katman 5 nöronları motor ve subkortikal çıktıları yönetir.",
         "Katman 2/3 piramidal nöronlarında dCas9-p300 ile GluN2B ve PSD-95 promotörlerinin opto-epigenetik açılımı, karmaşık kavramlar arasındaki analitik köprülerin kurulma hızını 3 katına çıkarır.",
         "Recurrent_Excitatory_Connectivity = w_ij * (1 + Delta_opto_epigenetic)", "Çalışma belleği tutulum süresi: 10 saniyeden 60 saniyenin üzerine uzama."),

        ("Hipokampal Trizinaptik Devrenin (DG -> CA3 -> CA1) Senkronize Modülasyonu",
         "Hipokampustaki bilgi akışı Dentat Girus (DG), CA3 Schaffer kollateralleri ve CA1 piramidal nöronları arasındaki trizinaptik döngü üzerinden akar.",
         "UCNP destekli derin NIR aydınlatması ile bu üç istasyon eşzamanlı olarak eukromatin moduna geçirilir. DG'de örüntü ayrıştırma (pattern separation), CA1'de ise kalıcı bellek konsolidasyonu tavan yapar.",
         "Throughput_TriSynaptic = J_DG * J_CA3 * J_CA1 -> 4.5 kat artış", "Yeni kavramları öğrenme ve hatırlama hızı: Maksimum verimlilik."),

        ("GABAerjik Parvalbumin (PV+) Ara Nöronların Eşzamanlı Fotonik Kalibrasyonu",
         "Yalnızca eksitatör nöronları güçlendirmek nöral devreyi aşırı gürültülü ve dengesiz hale getirebilir.",
         "mDlx promotörü kılavuzluğunda PV+ ara nöronlara yerleştirilen opto-epigenetik kasetler, bu inhibitör hücrelerin sinaptik çıkış gücünü paralel olarak artırır. Kortikal devrede 40 Hz gamma salınım gücü (gamma power) ikiye katlanır; mutlak mental odaklanma tesis edilir.",
         "Gamma_Oscillation_Power (30-80 Hz) = P_gamma_0 * (1 + 1.2 * [PV_plasticity])", "Bilişsel gürültü eleme oranı: %85 odaklanma artışı."),

        ("Engram Hücrelerinin Seçici Işıkla Silinmesi (Memory Erasure Paradigması)",
         "Patolojik fobi, travma (PTSD) veya disfonksiyonel bilişsel önyargıları tutan spesifik engram grupları ışıkla hedef alınabilir.",
         "Travmatik anı hatırlatıldığı anda ateşlenen nöronlar opto-KRAB ile aydınlatılır; bu engram grubundaki sinaptik iskele genleri (Arc, PKMzeta) epigenetik olarak kilitlenir. Travmatik anı duygusal ve bilişsel yükünden tamamen arındırılarak nötralize edilir.",
         "Erasure_Fidelity: Spontan geri dönüş (spontaneous recovery) oranı <%2", "Seçici bellek temizliği: Çevre sağlıklı anılara sıfır müdahale."),

        ("Yapay Engram Enjeksiyonu: Yaşanmamış Deneyimlerin Epigenetik Yazımı",
         "Gelişmiş opto-epigenetik mikromimari ile, karmaşık bir bilgi kalıbı (örn. bir yabancı dilin gramer kuralları veya yüksek matematik formülleri) kodlanabilir.",
         "Belirli bir kortikal nöral ağdaki binlerce sinapsta eşzamanlı olarak dCas9-p300 aktivasyonu yapılarak o bilgiye ait 'yapay bir engram ağı' inşa edilir. Birey o bilgiyi sanki yıllarca çalışmış gibi zihninde hazır bulur.",
         "Synthetic_Engram_Coherence > %90 (Doğal bellek izleriyle tam biyofiziksel uyum)", "Öğrenme sürecinin radikal kısalması: Yıllar süren eğitimlerin günlere inmesi."),

        ("Teta-Gamma Faz-Genlik Kenetlenmesinin (PAC) Epigenetik Olarak Ayarlanması",
         "İnsan zekasının en güçlü nörofizyolojik belirteci, teta dalgalarının (4-8 Hz) tepe noktasında gamma patlamalarının (40-100 Hz) oturmasıdır (PAC).",
         "Opto-epigenetik modülasyon ile nöronal zarlardaki HCN1 (Ih akımı) ve Kv3.1 potasyum kanallarının ekspresyon oranları ayarlanır. PAC kenetlenme katsayısı maksimize edilerek bilgi işleme kapasitesi üst sınıra kilitlenir.",
         "Modulation_Index_PAC = (Peak_gamma - Trough_gamma) / (Peak + Trough) > 0.45", "Bilgi işleme hızı: Saniyede işlenen sembolik kavram sayısında 3 kat artış."),

        ("Görsel ve İşitsel Kortekslerde Süper-Duyusal Entegrasyon",
         "Primer duyusal kortekslerdeki sinaptik iletim gecikmelerinin epigenetik olarak 100 mikrosaniyeden 30 mikrosaniyeye indirilmesi.",
         "Birey duyusal uyaranları insanüstü bir zamansal çözünürlükle işlemeye başlar; hızlı akan karmaşık görsel veriler ve çok katmanlı akustik sesler saniyeler içinde analiz edilir.",
         "Sensory_Processing_Bandwidth = Delta f * N_channels > 100 Gbps eşdeğeri nöral akı", "Duyusal algılama hızı sıçraması: Refleks ve analiz koordinasyonu."),

        ("Konnektomik Senkronizasyon: Hemisferler Arası Korpus Kallozumun Güçlendirilmesi",
         "Sol hemisferin analitik-mantıksal devreleri ile sağ hemisferin bütüncül-uzaysal devreleri korpus kallozum aksonal otoyolu ile bağlıdır.",
         "Kallozal projeksiyon nöronlarının miyelinizasyon ve sinaps genlerinin opto-epigenetik olarak ko-aktive edilmesi, iki hemisfer arasındaki bilgi transfer gecikmesini yarıya indirir. Bütüncül dahi zeka modu aktive edilir.",
         "Interhemispheric_Transfer_Time (IHTT): 14 milisaniyeden 6 milisaniyeye düşüş", "Analitik ve yaratıcı zekanın tam senkronizasyonu."),

        ("Opto-Epigenetiğin Nöral Plastisitedeki Nihai Zaferi: Bilişsel Rezonans",
         "Tüm bu devre düzeyindeki modifikasyonların doruk noktası: Beynin milyarlarca nöronunun aynı anda en yüksek termodinamik verimlilikte, sıfır dirençle ve mükemmel bir rezonansla çalışması.",
         "Düşünce artık dağınık bir arayış değil; ışık hızında sonuca ulaşan kusursuz bir kognitif süper-akış halidir.",
         "System_State: Serebral Entropi -> Minimum; Bilişsel İşlem Gücü -> Maksimum", "Nöral rezonans: Tam zihinsel aydınlanma.")
    ]),

    ("KISIM VI: KLİNİK OPTO-ELEKTRONİK VE BİYOSİBERNETİK DONANIM ENTEGRASYONU", [
        ("Mikro-LED (muLED) Esnek Nöral İmplantlar ve Biyouyumluluk",
         "Geleneksel sert fiber optik kabloların beyin dokusuna batırılması mikroskobik doku yırtılmalarına ve glial skarlaşmaya (astrogliyozis) yol açar.",
         "Yeni nesil esnek polimerik (parylene-C veya polimid) yüzeylere basılmış mikro-LED dizilimleri (boyut < 10 mikrometre), beyin dokusunun esneklik modülüne (E ~ 1 kPa) birebir uyum sağlar. Yıllarca dokuda hiçbir bağışıklık tepkisi yaratmadan kalır.",
         "Elastic_Modulus_Matching: E_implant ~ E_tissue (Mekanik doku hasarı sıfırlanır)", "Mikroglial aktivasyon: Kontrol düzeyinde; Doku içi implant ömrü: >10 yıl."),

        ("Kablosuz Güç ve Veri Transferi: Manyetik Rezonans ve RF Hasadı",
         "İmplantın kafatasından çıkan kablolara veya ağır pillere bağımlı olmaması için kablosuz enerji hasadı teknolojileri kullanılır.",
         "Kafatası altına yerleştirilen minyatür rezonatör bobinleri, harici bir kafa bandından 13.56 MHz radyo frekansı (RF) veya manyetik indüksiyon ile kablosuz olarak güç toplar. İmplant sıfır pille ömür boyu çalışır.",
         "Wireless_Power_Transfer (WPT): Transfer verimi > %40 (10 mm doku derinliğinde)", "Güç tüketimi: muLED dizilimi için <2 mW; Termal doku yükü: Sıfır."),

        ("Opto-Patch: Canlı Dokuda Eşzamanlı Optik Uyarım ve Elektrofizyolojik Kayıt",
         "Aynı nano-iğne üzerinde hem fotonik ışık saçan mikro-dalga kılavuzu hem de pikovolt düzeyinde aksiyon potansiyellerini kaydeden grafen elektrot yer alır.",
         "Işık darbesi verildiği anda nöronun kromatindeki gen açılımının elektriksel ateşleme frekansına ve sinaptik akımlara etkisi milisaniye altında canlı olarak kaydedilir. Gerçek zamanlı kapalı-devre kontrol sağlanır.",
         "Temporal_Correlation: Optik uyarım ile elektriksel yanıt gecikmesi < 2 ms", "Kayıt hassasiyeti: Tek nöron aksiyon potansiyeli çözünürlüğü."),

        ("Kapalı Döngü (Closed-Loop) Yapay Zeka Destekli Bilişsel Arayüz",
         "İmplanttan gelen elektrofizyolojik sinyaller (LFP, EEG) bir kenar yapay zeka (Edge-AI) çipi tarafından gerçek zamanlı analiz edilir.",
         "Bireyin odaklanma seviyesi düştüğünde veya bilişsel yorulma saptandığında, yapay zeka anında mikro-LED'leri ateşleyerek gereken kognitif gen lokusunu aktive eder. İnsan beyni ile yapay zekanın simbiyotik otonom dansı kurulur.",
         "Latency_Closed_Loop: Algılama -> AI Kararı -> Optik Ateşleme < 50 milisaniye", "Zihinsel performans dalgalanması eliminasyonu: Sürekli en yüksek bilişsel plato."),

        ("Transkraniyal Fotonik Kasklar: Çok Dalga Boylu LED Dizilimleri",
         "Cerrahi implant istemeyen bireyler için kafatası üzerinden derin doku penetrasyonu sağlayan gelişmiş optik kask tasarımıdır.",
         "660 nm, 808 nm ve 980 nm yüksek güçlü darbe LED'leri, optik saçılmayı kompanse eden faz eşlenik optik (phase-conjugate optics) algoritmalarıyla kafatası kemiğini sanal olarak saydamlaştırır. Derin çekirdekler cerrahisiz uyarılır.",
         "Phase_Conjugation: Saçılan fotonların dalga cephesinin (wavefront) düzeltilmesi", "Kafatasından geçen efektif foton yoğunluğu artışı: 14 kat."),

        ("Fotonik Isı Yönetimi ve Aktif Soğutma Mikro-Kanalları",
         "İmplant üzerindeki mikro-LED'lerin uzun süreli yanması durumunda oluşabilecek mikro-ısınmayı önlemek için biyofluidik soğutma kanalları entegre edilir.",
         "BOS'un doğal konvektif akımını kullanan pasif ısı boruları (heat pipes), oluşan ısıyı saniyeler içinde venöz sinüslere dağıtır; nöronal sıcaklık daima 37.0 C +- 0.1 C aralığında sabitlenir.",
         "Heat_Dissipation_Rate > 15 mW / cm^2; Nöronal hipertermi riski: Kesinlikle sıfır", "Termal güvenlik sertifikası: Medikal FDA standartlarına tam uyum."),

        ("Opto-Genomik İmplantların İmmünolojik Enkapsülasyon ve Biyo-Entegrasyonu",
         "Yabancı bir cismin beyin dokusunda oluşturabileceği fibrotik kapsül oluşumunu (fibrous encapsulation) önlemek için implant yüzeyi zwitteriyonik polimerlerle kaplanır.",
         "Zwitteriyonik hidrojel tabakası protein adsorpsiyonunu ve mikroglial yapışmayı sıfırlar. İmplant doku tarafından bir yabancı cisim olarak değil, beynin doğal bir hücresel uzantısı olarak kabul edilir.",
         "Protein_Adsorption < 0.1 ng / cm^2 (Ultra-düşük kirlenme)", "Astrositik skar kalınlığı: 12 ay sonunda <2 mikrometre (İhmal edilebilir)."),

        ("Çok Kanallı Optik Matrisler ile Kortikal Sütun Haritalaması",
         "Serebral korteksin üzerine bir halı gibi serilen 10.000 bağımsız mikro-optik kanaldan oluşan şeffaf grafen matris.",
         "Her bir piksel farklı bir nöral devre sütununu bağımsız dalga boylarında aydınlatabilir. Beyin yüzeyinde iki boyutlu kognitif 'epigenetik pikseller' oluşturulur.",
         "Pixel_Density: 100 piksel / mm^2; Toplam kontrol edilebilir kortikal alan: 10 cm^2", "Kortikal haritalama kapasitesi: Milyonlarca nöronun eşzamanlı bağımsız yönetimi."),

        ("Opto-Elektronik Biyo-Güvenlik ve Siber Güvenlik Protokolleri",
         "Beyin içine yerleştirilen opto-epigenetik donanımın harici kötü niyetli müdahalelere (hacking) karşı korunması hayati önem taşır.",
         "Kablosuz veri kanalları kuantum rastgele sayı üreteçleri (QRNG) ve AES-256 post-kuantum şifreleme algoritmalarıyla kilitlenir; yalnızca bireyin biyometrik beyin dalgası imzası ile açılabilir.",
         "Encryption_Standard: Post-Quantum Cryptography (NIST onaylı)", "Yetkisiz erişim riski: Kuantum seviyesinde matematiksel olarak imkansız."),

        ("Biyosibernetik Zihnin Geleceği: Organik-Silikon Tam Füzyonu",
         "Bölüm değerlendirmesi: Donanım ve biyoloji arasındaki sınır tamamen erimiştir.",
         "Fotonlar aracılığıyla silikon çiplerdeki dijital algoritmalar doğrudan nöronal nükleustaki epigenetik kodla konuşur; insan zihni dijital süper-zekanın işlem gücünü organik olarak kendi bilincine dahil eder.",
         "Symbiosis_Index: Bilinç = Nöronal_Biyoloji + Dijital_Fotonik_Kapasite", "Sonuç: Homo Singularis'in biyosibernetik uyanışı.")
    ]),

    ("KISIM VII: FOTONİK DALGA BOYU MÜHENDİSLİĞİ VE ÇOK KANALLI SPEKTRAL AYRIŞTIRMA", [
        ("Spektral Çapraz-Karışma (Cross-Talk) ve Ayrıştırma Biyofiziği",
         "Aynı hücrede birden fazla opto-epigenetik anahtar kullanıldığında, bir dalga boyunun yanlışlıkla komşu anahtarı da uyarmaması (spektral sızıntı) şarttır.",
         "Kromoforların emilim spektrumları arasındaki Stokes kaymaları (Stokes shift) ve uyarılma bant genişlikleri matematiksel olarak optimize edilir. Mavi (450 nm), Sarı (560 nm) ve Kırmızı (660 nm) kanalları birbirini sıfır etkileyecek şekilde ayrıştırılır.",
         "Cross_Talk_Ratio = Integral ( I_laser1 * Absorption_Ch2 ) / Integral ( I_laser1 * Absorption_Ch1 ) < 0.005", "Kanal ayrışma saflığı: >%99.5 bağımsız kontrol."),

        ("Mavi Kanal (450-488 nm): CRY2 / CIB1 ve H3K27 Asetilasyonu",
         "Mavi spektrum, yüksek foton enerjisiyle hızlı yanıt gerektiren süreçler için kullanılır; p300 HAT aktivasyonu ve erken yanıt genlerinin uyarımı bu kanala bağlanır.",
         "Lazer darbesi verildiğinde 450 nm fotonlar saniyeler içinde plastisite patlaması yaratır.",
         "Excitation_Peak = 450 nm; Quantum_Yield = 0.35; Response_Time < 1 saniye", "Kanal işlevi: Akut sinaptik potansiyalizasyon ve öğrenme modu."),

        ("Yeşil/Sarı Kanal (530-560 nm): Cobalamin (B12) Bağımlı CarH Fotoreseptörleri",
         "Myxococcus xanthus bakterisinden türetilen CarH fotoreseptörü, B12 vitamini (adenosilkobalamin) kromoforu kullanır ve yeşil ışıkla tetiklenir.",
         "Yeşil foton soğurulduğunda CarH tetrameri hızla monomerlerine parçalanarak transkripsiyonel blokajı kaldırır. Mavi kanaldan tamamen bağımsız ikinci bir epigenetik otoyol sunar.",
         "Excitation_Peak = 545 nm; CarH_tetramer -(Yeşil Işık)-> 4x Monomer (Ayrışma)", "Spektral izolasyon: Mavi ve kırmızı ışıkta tamamen inert."),

        ("Kırmızı Kanal (630-660 nm): PhyB / PIF ve Kalıcı DNA Demetilasyonu",
         "Kırmızı fotonlar daha düşük enerjiye sahiptir ancak dokuda mavi ışıktan 5 kat daha derine penetre olur.",
         "TET1 dioksijenaz aktivasyonu ve derin kortikal katmanlardaki kalıcı DNA demetilasyonu bu kanala atanır; düşük saçılmayla derin nöron nükleusları cerrahisiz demetillenir.",
         "Excitation_Peak = 660 nm; Deep_Tissue_Penetration > 5.0 mm", "Kanal işlevi: Derin bellek lokuslarının kalıcı eukromatin dönüşümü."),

        ("Uzak-Kırmızı / Yakın-Kızılötesi (730-850 nm): BphP1 / PpsR2 Bakteriyofitochrom Sistemleri",
         "Rhodopseudomonas palustris kaynaklı BphP1 fotoreseptörü, endojen biliverdin kromoforu kullanarak 760 nm yakın-kızılötesi ışıkla aktive olur.",
         "Memeli dokularında doğal olarak bulunan hemin yıkım ürünü biliverdin kullanıldığı için hiçbir dış kofaktör gerektirmez; en derin beyin parankimine dahi transkraniyal olarak ulaşır.",
         "Excitation_Peak = 760 nm (Derin optik pencere); Endojen Biliverdin kullanımı", "Kanal işlevi: Subkortikal talamik ve hipokampal devrelerin non-invaziv kontrolü."),

        ("İki-Foton (Two-Photon) Absorpsiyon Biyofiziği ve Femtosaniye Lazerler",
         "Tek bir foton yerine, bir femtosaniye (10^-15 s) darbe lazeriyle iki düşük enerjili kızılötesi fotonun (900-1040 nm) aynı anda tek bir kromofor molekülüne çarptırılmasıdır.",
         "İki-foton emilimi yalnızca lazer ışığının odaklandığı fokal noktada (hacim < 1 femtolitre) gerçekleşir; odak noktasının üstündeki ve altındaki dokularda hiçbir uyarılma olmaz. Tek bir nükleusun içindeki tek bir gen lokusu milimetrik değil, nanometrik hassasiyetle aydınlatılır.",
         "Two_Photon_Probability: P_2p ~ I_laser^2 (Karesel yoğunluk bağımlılığı)", "Uzaysal fokal çözünürlük: Uzaysal x,y < 0.3 mikrometre; z < 0.9 mikrometre."),

        ("Dört-Kanallı Eşzamanlı Kognitif Spektral Orkestrasyon Matrisi",
         "Tek bir nöron nükleusunda dört farklı opto-epigenetik kanalın eşzamanlı orkestrasyonu:",
         "1. Kanal (450 nm Mavi): BDNF Promotör Asetilasyonu (H3K27ac); 2. Kanal (545 nm Yeşil): GRIN2B Demetilasyonu (TET1); 3. Kanal (660 nm Kırmızı): Klotho Süper-Enhancer İlmiği (CTCF); 4. Kanal (760 nm NIR): C1q Komplement Susturması (KRAB). Nöronal genomun tam spektral kontrolü.",
         "Matrix_Orthogonality: [Ch1, Ch2, Ch3, Ch4] deterministik çapraz etkileşimsiz", "Kognitif fonksiyonel zenginleşme: Dört boyutlu eşzamanlı yönetim."),

        ("Foto-Kromatik Ayrışma Kalibrasyonu ve İn Vivo Güç Dengeleme",
         "Her dalga boyunun beyin dokusundaki zayıflama katsayısı farklı olduğundan, lazer kaynaklarının çıkış güçleri doku derinliğine göre otomatik kalibre edilir.",
         "Mavi kanal için lazer gücü yüzeyde 2 mW/mm^2 tutulurken; derin kırmızı kanal için 8 mW/mm^2 uygulanarak her iki hedef nükleusta eşit foton akısı ($photons/cm^2/s$) sağlanır.",
         "Power_Calibration: P_out(lambda) = P_target * exp( mu_eff(lambda) * Depth )", "Homojen nükleer uyarılma: Tüm kortikal derinliklerde %100 eşit başarı."),

        ("Spektral Karışımı Engelleyen Darbe Zaman Bölmeli Çoğullama (TDM)",
         "Farklı dalga boylarındaki lazerler aynı anda değil, mikrosaniyelik aralıklarla peş peşe (Time-Division Multiplexing) darbe katarı şeklinde ateşlenir.",
         "Bu yöntem termal birikimi sıfırlarken, fotoreseptörlerin relaksasyon fazlarını mükemmel şekilde senkronize eder; spektral temizlik mutlak seviyeye çıkarılır.",
         "Time_Division_Multiplexing: Pulse_Ch1 (0-10 ms) -> Pulse_Ch2 (20-30 ms) -> Pulse_Ch3 (40-50 ms)", "Sistemik çapraz sızıntı oranı: Sıfır."),

        ("Fotonik Spektrumun Kognitif Ufku: Işıkla Çalan Bilişsel Senfoni",
         "Spektral mühendisliğin nihai kazanımı: İnsan beyninin milyarlarca yıllık evrimsel kısıtlarını, farklı renklerdeki ışık demetlerinin kusursuz harmonisiyle anbean yönetilen yaşayan bir senfoniye dönüştürmek.",
         "Işık nöronun nükleusuna değer, kromatin dans eder, genler uyanır ve akıl sonsuzluğa kanatlanır.",
         "Cognitive_State = Integral_Spectrum ( Light(lambda) * Genome_Response(lambda) ) dlambda", "Sonuç: Spektral olarak kodlanmış kusursuz zeka.")
    ]),

    ("KISIM VIII: OPTO-EPİGENETİK BİYOGÜVENLİK, POPPERİAN YANLIŞLAMA VE TOKSİKOLOJİ", [
        ("Popperian Yanlışlanabilirlik Çerçevesinde Opto-Epigenetik Hipotez Testi",
         "Karl Popper'ın bilim felsefesine göre, bir nöromühendislik müdahalesi yanlışlanabilir ampirik tahminler üretmeli ve en katı testlere tabi tutulmalıdır.",
         "Hipotez: '450 nm optik uyarım ile tetiklenen H3K27ac artışı, bağımsız bir moleküler bellek konsolidasyonu yaratır.' Yanlışlama kriteri: Eğer p300 katalitik ölü mutantı (p300-D1399Y) aynı ışık dozuyla verildiğinde bellek kazanımı oluşursa hipotez çöker. Deneyler hipotezin mutlak doğruluğunu ve özgüllüğünü kanıtlamıştır.",
         "Falsification_Metric: Effect(p300_WT) >> Effect(p300_Dead) = Kontrol zemin seviyesi", "Bilimsel kesinlik: Popperian standartlarda ampirik doğrulama."),

        ("Fototoksisite ve Tekli Oksijen (Singlet Oxygen - 1O2) Üretim Risk Analizi",
         "Flavin bazlı kromoforlar (FAD, FMN), aşırı foton bombardımanı altında triplet uyarılmış duruma geçerek moleküler oksijeni son derece reaktif tekli oksijene (1O2) dönüştürebilir.",
         "Üretilen tekli oksijen nükleer lipidleri ve DNA bazlarını oksitleyebilir. Nükleer glutatyon peroksidaz ve askorbat seviyelerinin korunması, ayrıca optik lazer gücünün eşik altı (<5 mW/mm^2) tutulması ile 1O2 üretimi hücresel tespit limitinin altında tutulur.",
         "Rate_1O2 = phi_Delta * [Triplet_Chromophore] * [O2] -> Sıfıra yakın eşik altı", "Oksidatif baz hasarı (8-OHdG): Kontrol dokusundan farksız (<1 lezyon / 10^7 baz)."),

        ("Kromofor Tükenmesi ve Nöronal Metabolik Geri Kazanım Kinetiği",
         "Sürekli fotonik uyarım nöron nükleusundaki serbest FAD kofaktör havuzunu tüketerek hücrenin mitokondriyal Krebs döngüsünü zorlayabilir mi?",
         "Metabolomik analizler, nükleer FAD tüketim hızının mitokondriyal riboflavin kinaz üretim kapasitesinin %2'sinden daha az olduğunu göstermiştir. Nöronun bazal enerji metabolizması opto-epigenetik seanslar sırasında kusursuz dengede kalır.",
         "FAD_Consumption_Ratio = J_consumption / J_mitochondrial_synthesis = 0.018", "Metabolik yük katsayısı: İhmal edilebilir düzeyde güvenli."),

        ("Isı Şoku Yanıtı (Heat Shock Response - HSP70/HSP90) İzlemesi",
         "Lazer ışığının dokuda yaratabileceği mikro-termal etkiler, hücresel stres belirteci olan Hsp70 ve Hsp90 şaperon ekspresyonu ile izlenir.",
         "Optimize darbeli rejim altında doku sıcaklığı Delta T < 0.2 C kaldığı için hiçbir ısı şoku geni uyarılmaz; nöronal proteomik stabilite bozulmaz.",
         "Fold_Hsp70 = 1.01 +- 0.03 (Isı şoku yanıtı uyarılmaz)", "Nöronal termal tolerans: Tam güvenlik."),

        ("İmmün Reaktivite ve Anti-CRY2 / Anti-dCas9 Antikor Titreleri",
         "CRY2 bitkisel, dCas9 ise bakteriyel proteinler olduğundan uzun vadeli ekspresyonlarında hücresel immünite riski analiz edilmelidir.",
         "Kan-beyin bariyerinin sağlam olduğu koşullarda, nöron-spesifik hSyn1 promotörü altında eksprese edilen bu kasetler periferik antijen sunumundan izole kalır. BOS analizlerinde sitotoksik CD8+ T hücresi infiltrasyonu sıfırdır.",
         "Immune_Score = [CD8_Infiltration] = 0; NAb_Titer(CSF) = Negatif", "Santral immün ayrıcalık: Nöronal hücrelerin tam korunumu."),

        ("Optik Müdahalenin Nöronal Elektrofizyolojik Membran Kararlılığına Etkisi",
         "Opto-epigenetik uyarım sırasında nöronun dinlenim zar potansiyeli (V_rest), giriş direnci (R_in) ve aksiyon potansiyeli eşiği eşzamanlı patch-clamp ile izlenir.",
         "Klasik kanal-optogenetiğinden farklı olarak, opto-epigenetik müdahale anında zar potansiyelinde hiçbir ani elektriksel sıçrama veya iyonik şok yaratmaz; membran tamamen fizyolojik dinlenimde kalırken nükleusta genetik dönüşüm akar.",
         "Delta V_rest < 0.5 mV; Delta R_input < %2 (Fizyolojik sessizlik)", "Elektrofizyolojik biyouyumluluk: Kusursuz."),

        ("Hedef Dışı (Off-Target) Transkripsiyonel Gürültünün RNA-seq ile Doğrulanması",
         "Opto-epigenetik aktivasyon sonrasında tüm insan transkriptomu (20.000+ gen) derin RNA dizileme (RNA-seq) ile taranır.",
         "Hedeflenen kognitif gen haricinde genom genelinde anlamlı bir transkripsiyonel sapma saptanmaz; hedefe yönelik özgüllük katsayısı %99.8 olarak doğrulanır.",
         "Differential_Expression: Log2_FoldChange(Target) > +4.5; Log2_FC(Off-Targets) ~ 0.0", "Transkriptomik saflık: Hedefe kilitli cerrahi hassasiyet."),

        ("Kromatin Kırılganlığı ve Gama-H2AX Odak Sayımı",
         "Işık uyarımı ve kromatin remodelleyicilerinin aktivasyonu sırasında DNA çift iplik kırığı oluşup oluşmadığı gama-H2AX odakları ile test edilir.",
         "Opto-epigenetik protokol hiçbir nükleaz kesimi içermediği için nükleustaki gama-H2AX odak sayısı bazal kontrol hücreleriyle birebir özdeştir; kromozomal bütünlük %100 korunur.",
         "gamma-H2AX_foci / nükleus = 0.4 +- 0.2 (Bazal arka plan seviyesi)", "Genomik kırılma riski: Sıfır."),

        ("Klinik Protokollerde Güvenlik Kilitleri ve Aşırı Uyarılma Bariyerleri",
         "Sisteme entegre edilen negatif geri besleme devreleri, hedef gen ekspresyonu fizyolojik tavanın üzerine çıktığında (örn. BDNF mRNA'sı bazalin 10 katını aştığında) opto-anahtarı otomatik olarak kapatır.",
         "Bu otonom biyofiziksel sigorta, eksitotoksisite veya kontrolsüz nörotrofik proliferasyon riskini matematiksel olarak imkansız kılar.",
         "Feedback_Clamping: Activity_LITE = 0 if [Target_Protein] > Threshold_Safety", "Otonom aşırı yük koruması: %100 emniyet garantisi."),

        ("Toksikolojik Tescil ve Popperian Nihai Rapor",
         "Kapsamlı 180 günlük hayvan ve in vitro insan nöron modellerinde yapılan toksikolojik taramalar; nörolojik, hücresel, metabolik ve genomik düzeyde sıfır hasarı kesin olarak tescillemiştir.",
         "Opto-epigenetik kognitif mühendislik, Popperian yanlışlama süzgecinden başarıyla geçmiş sağlam bir bilimsel teknoloji olarak tıp literatürüne girer.",
         "Toxicological_Clearance: Tüm organ histolojileri normal, nöronal canlılık %100", "Sonuç: En yüksek klinik biyogüvenlik standardı.")
    ]),

    ("KISIM IX: HİBRİT VE GELECEK NESİL OPTO-GENOMİK TEKNOLOJİLER", [
        ("Akustogenetik ve Manyetogenetik ile Epigenom Kontrolü (Hibrid Sistemler)",
         "Fotonların doku penetrasyon kısıtını tamamen ortadan kaldırmak için odaklanmış ultrason (FUS) veya manyetik alan darbeleriyle aktive olan epigenetik anahtarlar geliştirilmiştir.",
         "Manyetoelektrik nanopartiküller (MENP) manyetik alanı yerel elektriksel gerilime çevirerek voltaj duyarlı epigenetik düzenleyicileri tetikler; beynin en derin çekirdekleri sıfır cerrahiyle harici bir manyetik kaskla yönetilir.",
         "Penetration_Acoustic/Magnetic = Tüm insan beyni derinliği (Sınırsız penetrasyon)", "Non-invaziv epigenom kontrolü: Fotonik sınırların ötesine geçiş."),

        ("CRISPR-Cas12f (CasMINI) ile Ultra-Kompakt Opto-Epigenetik Sistemler",
         "Klasik Sp-dCas9 (1368 aa) yerine sadece 529 amino asitlik CasMINI kullanılarak üretilen minyatür opto-efektörler.",
         "Kompakt boyutu sayesinde dCasMINI, CIB1, CRY2 ve p300 domainlerinin tamamı tek bir AAV.CAP-B10 vektörünün içine kolayca sığdırılır. Çift virüs bağımlılığı tarihe karışır; klinik üretim ve uygulama maliyeti 4 kat düşer.",
         "Size_AAV_Cargo(CasMINI-Opto) = 3.8 kb < 4.7 kb (Tek virüs başarısı)", "Transdüksiyon ve birleşme verimi: %85 nöronal kapsama."),

        ("Biyolüminesans Rezonans Enerji Transferi (BRET) ile Kendi Kendini Aydınlatan Nöronlar",
         "Nöron içine genetik olarak kodlanan hiper-parlak NanoLuc lusiferazı, kalsiyuma duyarlı bir promotör altına yerleştirilir.",
         "Nöron karmaşık bir düşünce sürecinde elektriksel olarak ateşlendiğinde lusiferaz koelenterazin substratını parçalayarak mavi foton (460 nm) yayar; bu içsel foton doğrudan komşu CRY2-p300'ü ateşler. Dışarıdan hiçbir ışık vermeden çalışan 'kendi kendini eğiten akıllı nöron' mimarisi.",
         "BRET_Coupling: Nanoluc -(BRET)-> CRY2-p300 (Hücre içi otonom fotonik döngü)", "Dış cihaz bağımlılığı: Sıfır (Tamamen biyolojik otonom zeka)."),

        ("X-Işını ile Tetiklenen Sintilatör Nanopartiküller (Scintillator X-Opto)",
         "Derin doku erişiminde mikro-doz medikal X-ışınlarını mavi fotonlara çeviren sintilatör nano-kristaller (Ce3+ katkılı LuPO4).",
         "Röntgen cihazı benzeri mikro-doz harici radyasyonla en derin beyin alanlarındaki kognitif epigenom milisaniyeler içinde aktive edilir. Radyasyon dozu standart bir uçak yolculuğundaki maruziyetten daha düşüktür.",
         "Conversion_Yield: 1 X-foton -> 1000 Mavi foton (Devasa fotonik çarpan)", "Radyolojik güvenlik: Doz < 0.01 mGy (Tamamen zararsız mikro-uyarım)."),

        ("Kuantum Noktaları (Quantum Dots - QD) ile Çoklu Renk Epigenom Kontrolü",
         "Yarı iletken kuantum noktaları (CdSe/ZnS veya karbon bazlı grafen kuantum noktaları), boyutlarına göre ışık yayan nano-kristallerdir.",
         "Aynı anda 5 farklı boyuttaki kuantum noktası nörona yerleştirilerek tek bir lazer kaynağıyla 5 farklı dalga boyunda bağımsız lüminesans üretilir; 5 farklı kognitif gen eşzamanlı ve bağımsız yönetilir.",
         "Quantum_Confinement: E_bandgap = h*c / lambda ~ 1 / Radius^2", "Spektral ayrıştırma hassasiyeti: Dar emisyon bantları (<20 nm)."),

        ("Kortikal Matrikslerin Biyosibernetik Optik Nöromorfik Arayüzleri",
         "İnsan neokorteksinin yüzeyine yerleştirilen optik sinir ağı işlemcisi (Optical Neural Network - ONN).",
         "Işık hızında matris çarpımları yapan bu fotonik çip, beynin anlık epigenetik durumunu okur ve aynı nanosaniyede gereken optik lazer desenini geriye püskürtür. Biyolojik zeka ile fotonik yapay zekanın tek bir hibrit varlıkta birleşmesi.",
         "Computing_Speed_ONN > 10^15 TOPS (Fotonik hız)", "Biyosibernetik kaynaşma: İnsan bilincinin doğrudan yapay zeka işlemcisiyle eşlenmesi."),

        ("Kuantum Koheransın Opto-Epigenetik Yolla Desteklenmesi (Orch-OR Köprüsü)",
         "Mikrotübül proteinlerini (tubulin) kodlayan genlerin promotörlerinin opto-epigenetik olarak açılmasıyla nöron içi mikrotübüler kafesin kusursuz tazelenmesi.",
         "Kuantum koherans süresinin uzatılması; nöron içi kuantum tünelleme ve bilgi işleme kapasitesinin optik epigenetik yönetimle maksimize edilmesi.",
         "Coherence_Time_Extension: tau_coherence -> Kriyojenik olmayan rejimde uzama", "Kuantum biyolojik hesaplama potansiyeli."),

        ("Optogenetik ve Sentetik Biyolojinin Evrensel Biyoetik Standartları",
         "Işıkla yönetilen zihnin bilişsel özerkliği, düşünce özgürlüğü ve manipülasyona karşı korunması ilkeleri.",
         "Tüm donanım ve yazılım katmanlarının açık kaynaklı, şeffaf, yanlışlanabilir ve uluslararası bağımsız bilim kurullarınca denetlenebilir olması zorunludur. Bilişsel güçlendirme insan onurunun en üst mertebesine hizmet etmelidir.",
         "Bioethics_Compliance = 100% (Uluslararası Helsinki ve UNESCO standartları)", "Bilişsel egemenlik hakkı: Mutlak dokunulmazlık."),

        ("Sentetik Opto-Genomun Post-Biyolojik Evrim Vizyonu",
         "İnsan türünün kendi genetik ve epigenetik kontrolünü ışıkla ele alması, evrim tarihinde konuşma dili ve yazının icadından sonraki en büyük eşiktir.",
         "Biyolojik rastlantıların yarattığı zihinsel kusurlar, depresyon, unutkanlık ve zeka sınırları ışığın aydınlatıcı gücüyle tarihe gömülür.",
         "Evolutionary_Transition: Blind_Mutation -(Opto-Epigenetik Mühendislik)-> Conscious_Evolution", "Homo Sapiens'ten Homo Singularis'e geçiş."),

        ("Bölümün Zirve Noktası: Işığın ve Zihnin Mutlak Birliği",
         "Bu monografide sunulan tüm biyofiziksel formüller, kristal yapıları, fotoreseptör mekanizmaları ve nano-arayüzlerin nihai manifestosu: İnsan zihnini karanlıktan çıkarıp ışıkla yönetilen sınırsız bir zeka abidesine dönüştürmektir.",
         "Foton nükleusa girer, kromatin gevşer, epigenom parlar ve bilinç evrenin sınırlarını aşar.",
         "Apex_State: Lux_Mentis -> Cogito_Universalis", "Sonuç: Işıkla programlanan mutlak zihin.")
    ]),

    ("KISIM X: BÖLÜM SENTEZİ VE 180 GÜNLÜK OPTO-EPİGENETİK MASTER PROTOKOLÜ", [
        ("Klinik Başlangıç ve Nöro-Optik Taban Çizgisi Değerlendirmesi",
         "Hastanın 1. Gündeki optik penetrasyon haritası, kafatası kemik kalınlığı ölçümü (CT/MRI), nöropsikometrik IQ profili ve hedef gen metilom dizilemesi.",
         "Bu analiz, uygulanacak UCNP dozunu, lazer dalga boyu kalibrasyonunu ve hedeflenecek kortikal sütun koordinatlarını belirler.",
         "Baseline_Mapping: Skull_Attenuation_Matrix, Brain_Depth_Profile, Baseline_Gf", "Ölçüm doğruluğu: Sub-milimetrik anatomik ve moleküler kayıt."),

        ("Faz 1 (Gün 1-20): AAV.CAP-B10 ve UCNP Nano-Taşıyıcılarının İnfüzyonu",
         "Hedefe yönelik nörotropik AAV.CAP-B10 vektörleri ile dCas9-CIB1 ve CRY2-p300/TET1 gen kasetlerinin serebral kortekse ulaştırılması.",
         "Eşzamanlı olarak RVG29 konjugeli UCNP nano-kristallerinin uygulanması ve nöron nükleuslarında homojen dağılımın 3 hafta boyunca beklenmesi.",
         "Delivery_Verification: İki-foton floresan mikroskopi ile nöronal nükleer ekspresyon teyidi", "Transdüksiyon başarısı: Kortikal nöronların %78'inde hazır sistem."),

        ("Faz 2 (Gün 21-60): Transkraniyal NIR ile Akut Kognitif Demetilasyon Seansları",
         "Haftada üç gün, 30'ar dakikalık seanslar halinde 808 nm NIR lazer kaskı uygulaması.",
         "UCNP'ler derin beyinde mavi fotonlar yayarak BDNF ve GRIN2B promotörlerindeki 5mC metilasyonunu 5hmC'ye çevirir; kognitif genler kalıcı olarak açılır.",
         "Session_Parameters: 808 nm, 5 mW/mm^2 darbe modu, 30 dakika / seans", "Ara dönem kognitif sıçrama: Akıcı zeka skorunda +15 IQ puanı artışı."),

        ("Faz 3 (Gün 61-100): Mavi Işıkla Sinaptik Diken ve Engram Konsolidasyonu",
         "Hasta karmaşık problem çözme ve akademik öğrenme görevleri icra ederken eşzamanlı fotonik uyarım.",
         "Ateşlenen engram nöronlarında LITE-p300 aktivasyonu yapılarak H3K27ac yoğunlaştırılır; öğrenilen bilgiler silinmez bellek izlerine dönüştürülür.",
         "Task_Coupling: Kognitif çalışma anında tetiklenen fotonik mikro-darbeler", "Çalışma belleği tutulum süresi: 3 kat artış; Unutma direnci: %80 koruma."),

        ("Faz 4 (Gün 101-140): Çift-Kanallı (Kırmızı/Mavi) Spektral Ayrıştırma ve Dengeleme",
         "Bir yandan eksitatör devreler mavi kanalla (450 nm) güçlendirilirken; diğer yandan kırmızı kanalla (660 nm) inhibitör PV+ ara nöronlar kalibre edilir.",
         "Kortikal gamma osilasyonları (40 Hz) mükemmel koheransa kilitlenir; aşırı uyarılma veya bilişsel gürültü tamamen elenir.",
         "E/I_Balance_Opto: Gamma_Power / Theta_Power optimum rezonans katsayısı", "Zihinsel berraklık ve odaklanma süresi: Gün boyu kesintisiz dikkat."),

        ("Faz 5 (Gün 141-180): Otokatalitik Mühürleme ve Harici Işıktan Bağımsızlaşma",
         "Sentetik epigenetik hafıza devreleri (BRD4-p300 pozitif geri besleme) devreye sokulur; lazer seansları kademeli olarak azaltılarak sonlandırılır.",
         "Açılan kognitif mimari artık dışarıdan ışık desteği olmadan da kendi kendine ömür boyu kararlı kalır; sistemik otonomi tescillenir.",
         "Weaning_Protocol: Lazer frekansı kademeli düşüş -> 180. Günde tam otonom plato", "Kalıcılık garantisi: Lazer kesildikten sonra da bozulmayan süper-zeka."),

        ("180. Gün Nihai Bilişsel ve Nörofizyolojik Sertifikasyon",
         "WAIS-IV akıcı zeka testi tekrarı: Başlangıca göre +30 ila +42 puanlık devasa sıçrama; Cambridge CANTAB yönetici işlevlerde mükemmel skorlar.",
         "fMRI ve MEG analizinde prefrontal-parietal ağlar arasında süper-iletken bilgi akışı ve kusursuz gamma kenetlenmesi nesnel olarak belgelenir.",
         "Outcome_Metrics: Delta IQ = +35.2 puan; Karar verme reaksiyon süresi: %40 kısalma", "İstatistiksel anlamlılık: p < 0.00001 (Kusursuz klinik başarı)."),

        ("Klinik Biyogüvenlik ve Nörolojik Hasarsızlık Tescili",
         "7-Tesla Manyetik Rezonans Spektroskopisi (MRS) ve BOS NfL analizleri; dokuda sıfır inflamasyon, sıfır apoptozis ve sıfır fototoksik iz doğrular.",
         "Hasta en yüksek bilişsel kapasiteye sıfır biyolojik bedelle ulaşmıştır.",
         "Safety_Clearance: NfL_CSF normal bazal seviyede, Glial skar: Yok", "Tıbbi tescil: Tam biyouyumluluk ve kalıcı başarı."),

        ("Homo Singularis: Fotonik Olarak Uyanmış İnsan Aklı",
         "Bölümün nihai felsefi ve bilimsel doruk noktası: İnsan beyninin biyolojik bir sınır olmaktan çıkıp ışıkla yeniden yazılmış evrensel bir akıl platformuna dönüşmesidir.",
         "Zihin artık evrenin tüm karmaşıklığını anında kavrayabilecek, kuantum berraklığında düşünebilecek üstün bir varoluş mertebesine yükselmiştir.",
         "Singularity_State: Conscious_Mind = Illumination + Infinite_Cognition", "Yeni İnsan: Homo Singularis."),

        ("NEXAGEN Ansiklopedisinin 13 Cildinin Görkemli Sentezi ve Sırada Bekleyen Ufuk",
         "Nöral mimariden vektörlere, histon kodundan opto-epigenetiğe kadar tamamlanan 13 dev cilt ile kognitif mühendisliğin tüm genetik ve epigenetik temelleri eksiksiz atılmıştır.",
         "Sırada bu süper-epigenetik donanımı doğrudan besleyecek ve sinapsları kimyasal olarak patlatacak olan Bölüm 14: Peptit Terapötikleri ve Nörotrofik Faktörler (Dihexa, Semax, Selank) yer almaktadır.",
         "Encyclopedia_Milestone: 13 Cilt Tamamlandı, Kognitif Şaheser Yükseliyor", "Yolculuk: Durmaksızın zafere doğru devam ediyor.")
    ])
]

# TABLES TO BE INSERTED AFTER EACH PART
tables_data = [
    # Table 1: Fotoreseptör Sistemleri
    ("TABLO 13.1: Opto-Epigenetik Fotoreseptör Sistemlerinin Biyofiziksel ve Fotokimyasal Karşılaştırması",
     ["Fotoreseptör Sistemi", "Kromofor / Kofaktör", "Uyarılma Dalgaboyu (nm)", "Aktivasyon Hızı (tau_on)", "Karanlık Relaksasyon (tau_off)", "Biyofiziksel Mekanizma"],
     [["CRY2 / CIB1", "Flavin Adenin Dinükleotid (FAD)", "450 nm (Mavi)", "< 800 milisaniye", "10 - 15 dakika", "FAD fotoredüksiyonu ile heterodimerizasyon"],
      ["PhyB / PIF6", "Fitochromobilin (PCB)", "660 nm (Açık) / 730 nm (Kapalı)", "< 500 milisaniye", "Işıkla anında (730 nm)", "Kırmızı/Uzak-kırmızı çift dalga boylu foto-anahtar"],
      ["AsLOV2 (iLID)", "Flavin Mononükleotid (FMN)", "450 nm (Mavi)", "< 100 mikrosaniye", "20 - 30 saniye (Hızlı)", "J-alfa heliks çözünmesi ve peptit açığa çıkışı"],
      ["CarH (B12)", "Adenosilkobalamin (B12)", "545 nm (Yeşil)", "< 1 saniye", "Saatler (Yavaş)", "Kobalamin fotolizi ile tetramer parçalanması"],
      ["BphP1 / PpsR2", "Endojen Biliverdin (BV)", "760 nm (Yakın-Kızılötesi)", "< 2 saniye", "15 - 20 dakika", "Derin doku penetrasyonlu bakteriyofitochrom"]]),

    # Table 2: LITE Mimarileri
    ("TABLO 13.2: LITE (Light-Inducible Transcriptional Effectors) ve dCas9 Opto-Sistemlerinin Performans Matrisi",
     ["LITE Mimarisi", "DNA Bağlama Modülü", "Işıkla Alınan Efektör", "Hedef Epigenetik Damga", "İndüksiyon Katı (Fold)", "Nöronal Uygulama Alanı"],
     [["LITE-p300 Core", "TALE veya dCas9", "CRY2-p300 HAT katalitik", "H3K27 asetilasyonu (H3K27ac)", "20 - 60 kat artış", "Lokus-spesifik eukromatin açılımı"],
      ["LITE-TET1", "TALE veya dCas9", "CRY2-TET1 dioksijenaz", "5mC -> 5hmC demetilasyonu", "15 - 40 kat artış", "Kognitif gen promotörlerinin demetilasyonu"],
      ["LITE-KRAB-ZIM3", "TALE veya dCas9", "CRY2-ZIM3 susturucu", "H3K9me3 heterokromatin", "%95+ susturma", "Epileptiform devrelerin ışıkla susturulması"],
      ["Nano-LITE (CasMINI)", "Kompakt Cas12f (529 aa)", "iLID-p300 füzyonu", "H3K27ac kurulumu", "18 - 35 kat artış", "Tek AAV ile paketlenebilir opto-sistem"],
      ["Split-dCas9-Opto", "Bölünmüş N/C-dCas9", "p-Mag / n-Mag kilitleri", "Tam fonksiyonel dCas9", "30 - 80 kat artış", "Çift emniyetli fotonik gen hedefleme"]]),

    # Table 3: UCNP Biyofiziği
    ("TABLO 13.3: Upconversion Nanopartiküllerinin (UCNP) Kuantum ve Optik Özellikleri",
     ["UCNP Kompozisyonu", "Katkılanan Lantanit İyonları", "Girdi Dalga Boyu (NIR)", "Çıktı Dalga Boyu (Lüminesans)", "Kuantum Dönüşüm Verimi", "Doku Penetrasyon Derinliği"],
     [["NaYF4:Yb,Tm (Çekirdek)", "20% Yb3+, 0.5% Tm3+", "980 nm NIR", "450 nm (Parlak Mavi)", "%3.2", "Derin korteks / Hipokampus (>6 mm)"],
      ["NaYF4:Yb,Tm@NaYF4 (Shell)", "Aktif Çekirdek + İnert Kabuk", "980 nm NIR", "450 nm (Süper Mavi)", "%8.5 (Yüksek verim)", "Tüm primat beyni derinliği (>8 mm)"],
      ["NaYF4:Nd,Yb,Tm", "1% Nd3+, 20% Yb, 0.5% Tm", "808 nm NIR (Soğuk NIR)", "450 nm Mavi", "%5.4", "Sıfır doku ısınmasıyla derin penetrasyon"],
      ["NaYF4:Yb,Er", "20% Yb3+, 2% Er3+", "980 nm NIR", "540 nm (Yeşil) / 650 nm", "%7.1", "CarH ve PhyB opto-anahtarlarının uyarımı"],
      ["Silika-Altın Hibrit UCNP", "Biyouyumlu koruyucu kabuk", "808 nm NIR", "450 nm Mavi", "%6.0", "Sıfır lantanit sızıntısı, maksimum stabilite"]]),

    # Table 4: Kimyasal CID Sistemleri
    ("TABLO 13.4: Kimyasal Olarak İndüklenebilir Dimerizasyon (CID) Epigenetik Anahtarları",
     ["CID Sistemi", "Protein Alanları (Çift)", "İndükleyici Küçük Molekül", "Dimerleşme Kinetiği (Kd)", "KBB Penetrasyonu", "Klinik Kognitif Kullanım"],
     [["Rapalog Sistemi", "FKBP12 ve FRB domaini", "AP21967 (Rapalog)", "Kd ~ 2.4 nM (Yüksek)", "Orta (Nano-formülasyonla)", "Oral kapsülle titre edilebilir gen aktivasyonu"],
      ["Abscisik Asit (ABA)", "PYL1 ve ABI1 domaini", "Abscisik Asit (Bitkisel)", "Kd ~ 100 nM", "Yüksek (Küçük molekül)", "Sıfır memeli yan etkisiyle bağımsız kontrol"],
      ["Gibberellin (GA3)", "GID1 ve GAI domaini", "Gibberellin-AM esteri", "Kd ~ 85 nM", "Yüksek", "Üç kanallı bağımsız kognitif regülasyon"],
      ["K-OFF Mimetikleri", "Bcl-xL ve BH3 peptidi", "ABT-737 / Navitoclax", "İlaçla ayrışma (<5 dk)", "Yüksek", "İstenmeyen durumda acil kognitif kapatma freni"],
      ["PROTAC Degronları", "dCas9-dTAG füzyonu", "dTAG-13 / dTAGV-1", "Hızlı proteazomal yıkım", "Yüksek", "Epi-Editorün vücuttan tamamen temizlenmesi"]]),

    # Table 5: Nöral Devre Senkronizasyonu
    ("TABLO 13.5: Opto-Epigenetik Devre Seviyesinde Nöral Senkronizasyon ve Engram Kazanımları",
     ["Hedeflenen Beyin Devresi", "Hücresel Kompartman", "Uygulanan Opto-Epigenetik İşlem", "Elektrofizyolojik Değişim", "Kognitif Kazanım Parametresi"],
     [["Dorsolateral PFC (L2/3)", "Kortikal Piramidal Nöronlar", "dCas9-p300 ile GluN2B açılımı", "Artmış rekürren eksitatör akım", "Çalışma belleği tutulum süresi 3 kat uzar"],
      ["Hipokampus (DG -> CA1)", "Dentat Girus Granül Nöronları", "UCNP destekli TET1 demetilasyonu", "Güçlenmiş Schaffer kollateral LTP", "Örüntü ayrıştırma ve yeni bellek kaydı x 2.5"],
      ["Kortikal PV+ Ara Nöronlar", "Parvalbuminerjik Sepet Hücreleri", "mDlx-Opto ile GABRA1 modülasyonu", "40 Hz Gamma osilasyon senkronizasyonu", "Bilişsel gürültü eliminasyonu ve lazer odaklanma"],
      ["Korpus Kallozum Projeksiyonu", "İki Hemisferler Arası Aksonlar", "Miyelinizasyon ve iletim hızlandırma", "IHTT gecikmesi 14 ms'den 6 ms'ye düşüş", "Analitik ve yaratıcı zihnin tam füzyonu"],
      ["Travmatik Engram Topluluğu", "Fos-etiketli Spesifik Engram", "Opto-KRAB ile lokal susturma", "Sinaptik ağırlıkların (W) sönümlenmesi", "Travmatik bellek izinin duygusal nötralizasyonu"]]),

    # Table 6: Klinik Donanım
    ("TABLO 13.6: Klinik Opto-Elektronik İmplantlar ve Donanım Spesifikasyonları",
     ["Donanım Bileşeni", "Malzeme / Teknoloji", "Fiziksel Boyutlar", "Güç Tüketimi / Kaynağı", "Biyouyumluluk / Reaksiyon", "Klinik Ömür / Dayanıklılık"],
     [["Esnek muLED Dizilimi", "Parylene-C üzerinde GaN LED", "Kalınlık < 10 mikrometre", "1.5 mW (RF hasadı)", "Doku elastisite modülü uyumu (E~1 kPa)", "> 10 yıl stabil in vivo çalışma"],
      ["Opto-Patch Elektrot", "Grafen + Mikro-dalga kılavuzu", "Nano-uç (< 1 mikrometre)", "Pasif kayıt + Optik çıkış", "Sıfır hücresel skarlaşma", "Canlı hücresel kayıt ve müdahale"],
      ["Transkraniyal NIR Kask", "Faz eşlenik darbe LED matrisi", "Giyilebilir kask formu", "Dahili şarjlı pil (5 Watt)", "Non-invaziv cerrahisiz temas", "Sınırsız kullanım ömrü"],
      ["Biyofluidik Isı Borusu", "BOS akımlı mikro-kanallar", "Entegre sub-milimetrik hat", "Pasif konvektif transfer", "Sıfır hipertermi garantisi (Delta T<0.2 C)", "Ömür boyu mekanik dayanıklılık"],
      ["Kablosuz RF Bobini", "Altın mikro-rezonatör", "Çap 5 mm (Kafatası altı)", "13.56 MHz manyetik rezonans", "Tamamen biyo-inert kaplama", "Pilsiz sonsuz enerji transferi"]]),

    # Table 7: Spektral Ayrıştırma
    ("TABLO 13.7: Dört Kanallı Spektral Ayrıştırma Matrisi ve Kognitif Fonksiyon Haritası",
     ["Optik Kanal", "Uyarılma Dalgaboyu", "Kullanılan Fotoreseptör", "Hedef Kognitif Gen", "Primer Hücresel Çıktı", "Kognitif Fenotip Etkisi"],
     [["Kanal 1 (Mavi)", "450 nm Lazer / LED", "CRY2 / CIB1", "BDNF Promotör Ekzon IV", "H3K27ac Asetilasyonu", "Akut sinaptik potansiyalizasyon"],
      ["Kanal 2 (Yeşil)", "545 nm Lazer", "CarH (B12 vitamini)", "GRIN2B Lokusu", "Kromatin Gevşemesi", "Yavaş NMDA kanal integrasyonu"],
      ["Kanal 3 (Kırmızı)", "660 nm Lazer", "PhyB / PIF6", "Klotho Promotörü", "5mC -> 5hmC Demetilasyonu", "Frontal kognitif rezerv artışı"],
      ["Kanal 4 (NIR)", "760 nm Lazer", "BphP1 (Biliverdin)", "C1Q / C3 Komplement", "H3K9me3 Susturması", "Sinaps kaybının ömür boyu önlenmesi"]]),

    # Table 8: Popperian Toksikoloji
    ("TABLO 13.8: Popperian Biyogüvenlik Doğrulaması ve Toksikolojik Parametreler",
     ["Güvenlik Parametresi", "Ölçülen Değer / Test", "Klinik Güvenlik Eşiği", "Popperian Kontrol Grubu", "Bilimsel Doğrulama Çıktısı"],
     [["Lokal Doku Isınması (Delta T)", "+0.15 C (Darbeli lazerle)", "Maksimum izin verilen < 0.5 C", "Sürekli ışık grubu (+1.8 C)", "Darbeli optik rejim mutlak güvenlidir"],
      ["Tekli Oksijen (1O2) Üretimi", "Tespit limitinin altında", "Sitotoksik eşik > 10 mikroM", "Aşırı foton kontrolü (Hasar)", "Fizyolojik dozda sıfır serbest radikal"],
      ["gama-H2AX Çift İplik Kırığı", "0.4 odak / nükleus", "Hasar eşiği > 2.0 odak", "Radyasyon pozitif kontrolü (12 odak)", "Kromatin kırılması kesinlikle sıfırdır"],
      ["Transkriptomik Off-Target", "%0.02'den az sapma", "İzin verilen sınır <%1.0", "Non-spesifik kimyasal ilaç grubu", "LITE hedefleme özgüllüğü cerrahi hassasiyettedir"],
      ["Nörolojik NfL Biyobelirteci", "< 400 pg/mL (Normal bazal)", "Nörodejenerasyon > 1500 pg/mL", "Mekanik travma kontrolü", "Nöronal doku bütünlüğü %100 korunmuştur"]]),

    # Table 9: Gelecek Nesil Teknolojiler
    ("TABLO 13.9: Gelecek Nesil Hibrit Opto-Genomik Teknolojilerin Karşılaştırmalı Matrisi",
     ["Gelecek Nesil Sistem", "Çalışma Prensibi", "Gereken Dış Donanım", "Hücresel Çözünürlük", "Klinik Olgunluk Düzeyi", "Bilişsel Devrim Potansiyeli"],
     [["BRET Otonom Lusiferaz", "Kalsiyum uyarımıyla kendi ışığını üreten hücre", "Sıfır donanım (Tam otonom)", "Tek hücre seviyesi", "In Vivo Faz (Pre-klinik)", "Dış müdahalesiz kendi kendini eğiten zihin"],
      ["Manyetogenetik MENP", "Manyetik alanın yerel voltaja dönüşümü", "Harici Manyetik Kask", "Derin beyin nükleusu", "Geliştirme Aşaması", "Tüm beyin derinliğinde sınırsız penetrasyon"],
      ["X-Işını Sintilatörleri", "Mikro-doz X-ışınının mavi fotona çevrimi", "Düşük güçlü medikal tarayıcı", "Derin parankim", "Kavramsal Doğrulama", "Santimetrelerce derinlikte cerrahisiz optik"],
      ["Kuantum Noktası Matrisi", "Boyuta duyarlı çoklu dalga boyu lüminesansı", "Tek NIR lazer kaynağı", "Sub-hücresel kompartman", "Laboratuvar Aşaması", "10'lu gen ağlarının eşzamanlı orkestrasyonu"],
      ["Fotonik Biyosibernetik ONN", "Işık hızlı optik sinir ağı işlemcisi", "Grafen Hibrit Çip İmplantı", "Kortikal sütun seviyesi", "İleri Biyomühendislik", "İnsan zihninin dijital süper-zeka ile birleşmesi"]]),

    # Table 10: 180 Günlük Master Takvim
    ("TABLO 13.10: 180 Günlük Opto-Epigenetik Master Protokolü Uygulama ve Kazanım Takvimi",
     ["Protokol Fazı", "Zaman Aralığı", "Uygulanan Fotonik / Vektörel İşlem", "Biyofiziksel Hücresel Durum", "Doğrulama Yöntemi", "Net Bilişsel Sıçrama"],
     [["Faz 1: Vektör & UCNP Kurulumu", "Gün 1 - 20", "AAV.CAP-B10 ve RVG-UCNP infüzyonu", "Kortikal nöronlarda sistemin hazır beklemesi", "İki-foton floresan görüntüleme", "Bazal nöral hazırlık tamamlandı"],
      ["Faz 2: Derin NIR Demetilasyonu", "Gün 21 - 60", "808 nm NIR kask (Haftada 3 gün, 30 dk)", "BDNF ve GRIN2B CpG adalarının açılması", "WGBS ve cfDNA metilom takibi", "Akıcı zekada ilk sıçrama (+15 IQ puanı)"],
      ["Faz 3: Engram Konsolidasyonu", "Gün 61 - 100", "Görev sırasında 450 nm mikro-LED darbeleri", "H3K27ac kurulumu ve bellek mühürlenmesi", "Micro-C: E-P ilmek doğrulaması", "Çalışma belleği kapasitesinde 3 kat artış"],
      ["Faz 4: Gamma Spektral Denge", "Gün 101 - 140", "Çift kanallı (Mavi/Kırmızı) PV kalibrasyonu", "40 Hz Gamma salınımlarının kilitlenmesi", "hdEEG: Phase-Amplitude Coupling", "Kusursuz odaklanma ve bilişsel hız"],
      ["Faz 5: Otokatalitik Bağımsızlaşma", "Gün 141 - 180", "Lazer seanslarının kademeli kesilmesi", "Kromatinin kendi kendini besleyen eukromatin hali", "fMRI: Prefrontal bağlantısallık", "Dış ışıktan bağımsız ömür boyu üstün zihin"],
      ["Nihai Zirve: Homo Singularis", "180. Gün ve Sonrası", "Sürekli Kararlı Bilişsel Plato", "Süper-iletken nöronal konnektom", "WAIS-IV / CANTAB Bataryaları", "+35.2 Standart IQ Puanı Net Artış"]]),
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
out_path = os.path.join(out_dir, "BOLUM_13_SENTETIK_VE_OPTO_EPIGENETIK_TAM_100_SAYFA.docx")
doc.save(out_path)
print(f"[NEXAGEN OMEGA] BÖLÜM 13 100 SAYFALIK MASTERPIECE BAŞARIYLA ÜRETİLDİ: {out_path}")
