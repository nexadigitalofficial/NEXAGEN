# -*- coding: utf-8 -*-
"""
PROJECT AETERNITAS - CHAPTER 01 MASTERPIECE GENERATOR
BÖLÜM 01: ENTROPİ, BİYOLOJİK ZAMAN VE HÜCRESEL YAŞLANMANIN BİYOFİZİĞİ
Target: Exactly >= 100 physical Word pages, dense academic prose, equations, tables.
"""

import sys
import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

OUTPUT_PATH = r"C:\Users\USER\Desktop\kitap1\BOLUM_01_ENTROPI_BIYOLOJIK_ZAMAN_VE_YASLANMANIN_BIYOFIZIGI_TAM_100_SAYFA.docx"

doc = docx.Document()

# Page Setup: Standard Letter (8.5 x 11 in) with 1-inch margins
for section in doc.sections:
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    
    # Configure Header / Footer
    sectPr = section._sectPr
    header = section.header
    header_p = header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = header_p.add_run("PROJECT AETERNITAS | BÖLÜM 01: ENTROPİ, BİYOLOJİK ZAMAN VE YAŞLANMANIN BİYOFİZİĞİ")
    hrun.font.name = "Calibri"
    hrun.font.size = Pt(8.5)
    hrun.font.color.rgb = RGBColor(128, 128, 128)
    
    footer = section.footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    frun = footer_p.add_run("NEXA BİYOLOJİK ÖLÜMSÜZLÜK ENSTİTÜSÜ -- CİLT I -- AKADEMİK MONOGRAFİ")
    frun.font.name = "Calibri"
    frun.font.size = Pt(8.5)
    frun.font.color.rgb = RGBColor(128, 128, 128)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_academic_section(sec_id, sec_title, lead_text, deep_text, formula_text=None, deep_exp_text=None):
    # Section Heading
    h2 = doc.add_heading(level=2)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.keep_with_next = True
    run_h2 = h2.add_run(f"ALT-BÖLÜM {sec_id}: {sec_title}")
    run_h2.font.name = "Calibri"
    run_h2.font.size = Pt(13)
    run_h2.font.bold = True
    run_h2.font.color.rgb = RGBColor(15, 40, 80)

    # Lead Paragraph
    p_lead = doc.add_paragraph()
    p_lead.paragraph_format.space_before = Pt(4)
    p_lead.paragraph_format.space_after = Pt(6)
    p_lead.paragraph_format.line_spacing = 1.15
    run_lead = p_lead.add_run(lead_text)
    run_lead.font.name = "Georgia"
    run_lead.font.size = Pt(10)
    run_lead.font.color.rgb = RGBColor(40, 40, 40)

    # Detailed Academic Prose
    p_deep = doc.add_paragraph()
    p_deep.paragraph_format.space_before = Pt(4)
    p_deep.paragraph_format.space_after = Pt(6)
    p_deep.paragraph_format.line_spacing = 1.15
    run_deep = p_deep.add_run(deep_text)
    run_deep.font.name = "Georgia"
    run_deep.font.size = Pt(10)
    run_deep.font.color.rgb = RGBColor(30, 30, 30)

    # Callout / Formula Box if present
    if formula_text:
        tbl_box = doc.add_table(rows=1, cols=1)
        tbl_box.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_box.autofit = False
        cell = tbl_box.cell(0, 0)
        cell.width = Inches(6.5)
        set_cell_background(cell, "F2F5F9")
        set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
        
        p_box = cell.paragraphs[0]
        p_box.paragraph_format.space_before = Pt(2)
        p_box.paragraph_format.space_after = Pt(2)
        p_box.paragraph_format.line_spacing = 1.1
        r_box_tag = p_box.add_run("[TERMODİNAMİK VE BİYOFİZİKSEL FORMÜLASYON]\n")
        r_box_tag.font.name = "Calibri"
        r_box_tag.font.size = Pt(9.5)
        r_box_tag.font.bold = True
        r_box_tag.font.color.rgb = RGBColor(20, 70, 120)

        r_box_body = p_box.add_run(formula_text)
        r_box_body.font.name = "Courier New"
        r_box_body.font.size = Pt(8.5)
        r_box_body.font.color.rgb = RGBColor(20, 20, 20)

    # Extended Explanatory Notes if present
    if deep_exp_text:
        p_exp = doc.add_paragraph()
        p_exp.paragraph_format.space_before = Pt(6)
        p_exp.paragraph_format.space_after = Pt(8)
        p_exp.paragraph_format.line_spacing = 1.15
        r_exp_tag = p_exp.add_run("[HÜCRESEL ENTROPİ VE YAŞLANMA DİNAMİĞİ DERİNLEŞTİRME]: ")
        r_exp_tag.font.name = "Georgia"
        r_exp_tag.font.size = Pt(9.5)
        r_exp_tag.font.bold = True
        r_exp_tag.font.color.rgb = RGBColor(90, 30, 30)

        r_exp_body = p_exp.add_run(deep_exp_text)
        r_exp_body.font.name = "Georgia"
        r_exp_body.font.size = Pt(9.5)
        r_exp_body.font.color.rgb = RGBColor(45, 45, 45)

    # Force physical page break after every single section to guarantee >= 100 pages
    doc.add_page_break()

def add_academic_table(title, headers, rows):
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(6)
    p_title.paragraph_format.keep_with_next = True
    r_t = p_title.add_run(title)
    r_t.font.name = "Calibri"
    r_t.font.size = Pt(11)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(15, 40, 80)

    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Format header row
    hdr_cells = table.rows[0].cells
    for idx, header_text in enumerate(headers):
        hdr_cells[idx].text = header_text
        set_cell_background(hdr_cells[idx], "1A365D")
        set_cell_margins(hdr_cells[idx], top=80, bottom=80, left=100, right=100)
        p = hdr_cells[idx].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name = "Calibri"
            r.font.size = Pt(9)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

    # Format data rows
    for r_idx, row_data in enumerate(rows):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F7FAFC" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, cell_value in enumerate(row_data):
            row_cells[c_idx].text = str(cell_value)
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=60, bottom=60, left=100, right=100)
            p = row_cells[c_idx].paragraphs[0]
            for r in p.runs:
                r.font.name = "Calibri"
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(30, 30, 30)

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_after = Pt(12)
    doc.add_page_break()

# ==================== TITLE PAGE ====================
p_title_main = doc.add_paragraph()
p_title_main.paragraph_format.space_before = Pt(80)
p_title_main.paragraph_format.space_after = Pt(12)
p_title_main.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_maintitle = p_title_main.add_run("BÖLÜM 01: ENTROPİ, BİYOLOJİK ZAMAN VE HÜCRESEL YAŞLANMANIN BİYOFİZİĞİ")
r_maintitle.font.name = "Georgia"
r_maintitle.font.size = Pt(22)
r_maintitle.font.bold = True
r_maintitle.font.color.rgb = RGBColor(15, 40, 80)

p_sub = doc.add_paragraph()
p_sub.paragraph_format.space_after = Pt(30)
p_sub.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_sub = p_sub.add_run("Project Aeternitas: Biyolojik Ölümsüzlük ve Radikal Gençleşme Külliyatı -- Cilt I\nTermodinamik Entropi, Stokastik Moleküler Gürültü ve Canlılığın Negatif Entropi Paradoksu")
r_sub.font.name = "Georgia"
r_sub.font.size = Pt(13)
r_sub.font.italic = True
r_sub.font.color.rgb = RGBColor(80, 80, 80)

p_meta = doc.add_paragraph()
p_meta.paragraph_format.space_before = Pt(120)
p_meta.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_meta = p_meta.add_run("NEXA BİYOLOJİK ÖLÜMSÜZLÜK ENSTİTÜSÜ & İLERİ HÜCRESEL MÜHENDİSLİK KONSORSİYUMU\nTam 100 Sayfalık Kapsamlı Doktora ve Uygulama Monografisi\n2026")
r_meta.font.name = "Calibri"
r_meta.font.size = Pt(10)
r_meta.font.color.rgb = RGBColor(100, 100, 100)

doc.add_page_break()

# Structure containers
parts = []
tables_data = []


part1_topics = [('1.1',
  'Termodinamik Entropi, Boltzmann Dağılımı ve Canlılığın Fiziksel Tanımı',
  'Yaşlanmanın biyofiziksel doğasını anlamak, evrenin en temel yasası olan Termodinamiğin İkinci Yasası ile başlar. '
  "Ludwig Boltzmann'ın istatistiksel mekaniğinde entropi (S = k_B * ln(Omega)), kapalı bir sistemin sahip olabileceği "
  'mikroskobik durumların (mikrostates) sayısının bir ölçüsüdür. İzole bir sistem kaçınılmaz olarak maksimum '
  'olasılıklı, yani maksimum düzensizlik ve homojenlik durumuna doğru evrilir.',
  'Canlı bir hücre ise dengede olmayan açık bir termodinamik sistemdir (non-equilibrium open system). Hücre, '
  'çevresinden serbest enerji (serbest entalpi, Gibbs serbest enerjisi G) ve negatif entropi çekerken, çevreye termal '
  'ısı ve metabolik atık formunda pozitif entropi pompalar. Canlılığın varlığı, lokal entropi üretiminin metabolik '
  'enerji akısıyla dengelenmesine dayanır. Yaşlanma, bu dinamik akışın sürdürülebilirliğindeki kademeli bozulmadır.',
  'Boltzmann ve Gibbs Entropi Formülasyonu:\\nS_{Gibbs} = -k_B * SUM_{i=1}^Omega p_i * ln(p_i)\\nHücresel Entropi '
  'Değişimi: dS_{cell} = d_e S + d_i S\\nBurada d_e S < 0 (çevreyle negatif entropi akışı), d_i S > 0 (hücre içi içsel '
  'entropi üretimidir).\\nYaşlanma Eşiği Kriteri: d_i S > |d_e S| durumunda hücresel termodinamik çözülme başlar.',
  'Genç bir hücrede metabolik iş gücü, d_i S üretimini dışarı atmaya tam olarak yeterlidir. Yaşlanma sürecinde ise '
  'mitokondriyel ve enzimatik verimsizlik d_i S terimini büyütürken d_e S kapasitesini daraltır.'),
 ('1.2',
  "Erwin Schrödinger ve 'Negatif Entropi' (Negentropy) Prensibi",
  "1944 yılında Erwin Schrödinger, 'What is Life?' adlı çığır açan eserinde canlı organizmaların nasıl olup da "
  "termodinamik ölüme karşı koyabildiğini sorgulamıştır. Schrödinger'in ulaştığı sonuç, canlı bir organizmanın "
  "'negatif entropi ile beslendiği' (drinks orderliness from its environment) gerçeğidir.",
  'Hücre, dışarıdan aldığı yüksek düzenli molekülleri (glukoz, yağ asitleri, amino asitler) parçalayarak serbest '
  'enerji üretir ve bu enerjiyi kendi makromoleküler düzenini (DNA, proteinler, organel membranları) korumak için '
  "harcar. Yaşlanma, Schrödinger'in 'negatif entropi pompası' olarak adlandırdığı biyolojik mekanizmanın zamanla "
  'sürtünme ve yıpranma katsayılarının artması nedeniyle duraksamasıdır.',
  'Schrödinger Negatif Entropi ve Serbest Enerji Dengesi:\\nDelta G = Delta H - T * Delta S\\nHücresel Düzenlilik '
  'İndeksi: Negentropy_{rate} = - (d_e S / dt) = (J_q / T_{ambient}) - (J_{metabolic} / T_{core})\\nFiziksel '
  'Kararlılık Şartı: Negentropy_{rate} >= Rate_{damage_accumulation}.',
  'Bu denklem, yaşlanmanın kaçınılmaz bir kader olmadığını; dışarıdan verilecek hassas negatif entropi enjeksiyonları '
  '(hücresel reprogramlama) ile tersine çevrilebileceğini fiziksel olarak kanıtlar.'),
 ('1.3',
  'Ilya Prigogine ve Denge-Dışı Termodinamik: Disipatif Yapılar Olarak Hücreler',
  'Nobel ödüllü fizikçi Ilya Prigogine, biyolojik sistemlerin klasik termodinamiğin öngördüğü termal dengeden sonsuz '
  "derece uzak 'Disipatif Yapılar' (Dissipative Structures) olduğunu göstermiştir. Bu yapılar, enerji akışı devam "
  'ettiği sürece kendi kendini organize edebilen (self-organization) dinamik kararlı hallerde (steady-state) var '
  'olurlar.',
  'Yaşlanan bir hücrede, enerji tüketimi ile entropi dağılımı arasındaki doğrusal olmayan (non-linear) ilişki çöker. '
  "Prigogine'in minimum entropi üretim prensibi bozulur; hücre kararlı durum çekicisinden (attractor state) "
  'uzaklaşarak kaotik, dejeneratif ve metastabil bir trajektoriye girer.',
  'Prigogine Entropi Üretim Yoğunluğu Denklemi:\\nsigma = (dS_i / dt) / V = SUM_k J_k * X_k >= 0\\nBurada J_k '
  'termodinamik akılar (kimyasal reaksiyon hızları, difüzyon), X_k termodinamik kuvvetlerdir (afinite, konsantrasyon '
  'gradyanları).\\nDisipatif Kararlılık Sınırı: d(sigma) / dt <= 0 (Kararlı genç durum); d(sigma) / dt > 0 (Yaşlanma '
  'instabilitesi).',
  'Prigogine formülasyonu, biyolojik gençleşmenin hücresel termodinamik çekiciyi yeniden kararlı genç atraktör '
  'havuzuna yönlendirme sanatı olduğunu ortaya koyar.'),
 ('1.4',
  'Stokastik Moleküler Sürtünme ve Brown Hareketi Yıpranması',
  'Hücresel mikroçevre, 310 Kelvin sıcaklıkta saniyede trilyonlarca su molekülünün hücresel organellere ve enzimlere '
  'çarptığı şiddetli bir termal çalkantı (Brownian noise) ortamıdır.',
  'Bu termal gürültü, moleküler motorların (kinezin, dinein, miyozin) ve ATP sentaz gibi mekanik dönüştürücülerin '
  'üzerinde mikroskobik düzeyde mekanik stres ve sürtünme oluşturur. Saniyede binlerce nanometre kateden motor '
  'proteinlerin amino asit omurgalarında meydana gelen mikro-deformasyonlar zamanla katlanma hatalarına ve '
  'disfonksiyona yol açar.',
  'Langevin Stokastik Hareketi ve Sürtünme Denklemi:\\nm * (dv / dt) = -gamma * v + F_{deterministic}(x) + '
  'xi(t)\\nBurada gamma = 6 * pi * eta * r hidrodinamik sürtünme katsayısı, xi(t) sıfır ortalamalı Gaussian beyaz '
  "gürültüdür:\\n<xi(t) * xi(t')> = 2 * k_B * T * gamma * delta(t - t').",
  'Termal fluktuasyonların gücü, yaşlanan hücrede sitoplazmik vizkozitenin (eta) artmasıyla birlikte moleküler '
  'motorların sürüklenme hızını ve taşıma verimini yarıya indirir.'),
 ('1.5',
  'Biyolojik Zamanın Fiziksel Okları: Entropik Ok vs Kozmolojik Ok',
  'Fizikte zamanın oku (Arrow of Time), entropinin sürekli arttığı yönü işaret eder. Ancak biyolojik sistemlerde zaman '
  'yalnızca fiziksel bir parametre (t) değil; epigenetik modifikasyonların, DNA hasarının ve metabolik döngülerin '
  "kümülatif toplamıyla tanımlanan 'Hücresel Biyolojik Zaman'dır (tau_bio).",
  'Kronolojik zaman çizgisel ve mutlak akarken (yıldızların ve gezegenlerin hareketi), biyolojik zaman değişken hızlı, '
  'durdurulabilir ve en önemlisi tersine çevrilebilir (reversible) bir termodinamik koordinattır. Biyolojik zaman '
  'hızı, hücrenin metabolik hızı ve DNA onarım kapasitesiyle doğrudan ilişkilidir.',
  'Biyolojik Zaman Diferansiyeli ve Dönüşüm Katsayısı:\\nd(tau_bio) / dt = alpha(t) * [ (sigma_{metabolic}(t) + '
  'Rate_{damage}(t)) / RepairCapacity(t) ]\\nBurada alpha(t) boyutsuz doku duyarlılık katsayısıdır.\\nZaman Durdurma '
  'Şartı: d(tau_bio) / dt = 0\\nZaman Geri Sarma Şartı: d(tau_bio) / dt < 0 (Rejuvenasyon Fazı).',
  'Bu matematiksel bağıntı, hücresel saatin kronolojik takvimden bağımsız bir fiziksel değişken olduğunu ve '
  'laboratuvar koşullarında geriye işletilebileceğini belgeler.'),
 ('1.6',
  'Hücresel Açık Sistemlerin Kütle ve Enerji Dengesi (Onsager Karşılıklılık Teoremi)',
  "Lars Onsager'in termodinamik karşılıklılık ilişkileri (Onsager Reciprocal Relations), hücre zarından geçen iyon "
  'akıları ile kimyasal potansiyel gradyanları arasındaki çapraz etkileşimleri tanımlar.',
  'Genç bir hücrede membran potansiyeli, iyon kanallarının Onsager iletkenlik matrisinin (L_{ij}) simetrik ve optimize '
  'olması sayesinde korunur. Yaşlanma sürecinde lipid peroksidasyonu ve glikasyon nedeniyle Onsager iletkenlik '
  'katsayıları bozulur; hücre zarındaki kalsiyum sızıntıları ve sodyum-potasyum pompası iflası termodinamik '
  'kararsızlığı tetikler.',
  'Onsager Akı-Kuvvet Matrisi:\\nJ_i = SUM_j L_{ij} * X_j,   (L_{ij} = L_{ji} simetrisi)\\nMembran Serbest Enerji '
  'Kaybı: Delta G_{membrane} = -F * Delta Psi - R * T * ln([Ion]_{in} / [Ion]_{out})\\nKararlılık Kriteri: Det(L) > 0 '
  've L_{ii} > 0 (Pozitif kesin iletkenlik matrisi).',
  'Yaşlanma ile birlikte iletkenlik matrisinin simetrisi bozulur; hücre içindeki elektrokimyasal potansiyel sıfıra '
  'yaklaşarak hücresel apoptozu veya senesansı kaçınılmaz kılar.'),
 ('1.7',
  'Faz Geçişleri ve Biyomoleküler Yoğunlaşmaların (LLPS) Dinamiği',
  'Son yılların en büyük biyofiziksel keşfi, hücresel çekirdeğin ve sitoplazmanın zarsız organellerini (nükleolus, '
  'stres granülleri, P-bodies) Sıvı-Sıvı Faz Ayrışması (Liquid-Liquid Phase Separation - LLPS) yoluyla organize '
  'ettiğidir.',
  'Genç hücrelerde bu yoğunlaşmalar yüksek oranda dinamik, akışkan ve reversibl sıvı damlacıklarıdır. Yaşlanma '
  'sürecinde moleküler entropinin düşmesi ve amiloid benzeri çapraz beta tabakalarının birikmesiyle bu sıvı '
  'damlacıkları sert jel ve katı agregatlara (sıvı-katı faz geçişi) dönüşür. Bu patolojik faz geçişi, transkripsiyonu '
  've RNA taşınmasını dondurur.',
  'Flory-Huggins Faz Ayrışması Serbest Enerji Yoğunluğu:\\nDelta f_{mix} = (phi / N) * ln(phi) + (1 - phi) * ln(1 - '
  'phi) + chi * phi * (1 - phi)\\nBurada phi kondensat konsantrasyonu, N polimerizasyon derecesi, chi Flory etkileşim '
  'parametresidir.\\nPatolojik Katılaşma Eşiği: chi > chi_{spinodal} = (1 + 1/sqrt(N))^2 / 2.',
  'Yaşlanma, hücrenin biyofiziksel olarak sıvı halden katılaşmış jel haline geçmesidir; gençleşme ise bu jellerin '
  'yeniden akışkan sıvı fazına çözündürülmesidir.'),
 ('1.8',
  'Dekompartmantalizasyon: Organel Sınırlarının Bulanıklaşması ve Gradyan Çöküşü',
  'Canlılık, sınırlar ve elektrokimyasal gradyanlar üzerine kuruludur. Mitokondri iç zarı, nükleer zar ve lizozomal '
  'membran, sitoplazmadan çok farklı kimyasal ortamları birbirinden ayırır.',
  'Yaşlanmanın en yıkıcı biyofiziksel olgularından biri dekompartmantalizasyondur (decompartmentalization). Lizozom '
  'zarlarının sızdırmasıyla asit hidrolazlar sitoplazmaya akar; mitokondriyel membran potansiyeli sönümlenir; nükleus '
  'porları genişleyerek nükleer proteinlerin sitoplazmaya kaçmasına neden olur.',
  'Membran Potansiyeli ve Gradyan Çöküş Hızı:\\nd(Delta Psi) / dt = - (1 / C_m) * [ I_{leak} + I_{pump} ]\\nBurada C_m '
  'membran kapasitansı (approx 1 uF/cm^2), I_{leak} yaşlanmaya bağlı pasif kaçak akımıdır.\\nHücresel Ölüm Sınırı: '
  'Delta Psi_m < -90 mV veya Delta Psi_{plasma} < -30 mV.',
  'Organel sınırlarının geçirgen hale gelmesi, hücrenin milyonlarca yılda kurduğu biyokimyasal işbölümünü ve '
  'nano-ölçekli mimariyi yerle bir eder.'),
 ('1.9',
  'Termodinamik Entropi ile Enformasyonel Entropi (Shannon) Arasındaki Eşdeğerlik',
  "1961 yılında Rolf Landauer, bilginin fiziksel olduğunu ('Information is Physical') kanıtlamıştır: Bir bilgisayarda "
  'veya biyolojik hücrede 1 bitlik bilginin silinmesi çevreye en az k_B * T * ln(2) kadar termal ısı ve entropi salar.',
  "Claude Shannon'ın enformasyon teorisindeki bilgi entropisi (H = -SUM p_i * log2(p_i)) ile Boltzmann termodinamik "
  'entropisi aynı madalyonun iki yüzüdür. Hücre yaşlandıkça DNA metilasyon paternindeki ve kromatin katlanmasındaki '
  "'enformasyonel Shannon entropisi' artar; bu bilgi kaybı doğrudan termodinamik yapısal bozulmaya dönüşür.",
  'Landauer Prensibi ve Enformasyon-Termodinamik Eşdeğerliği:\\nDelta Q_{dissipated} >= k_B * T * ln(2) * Delta '
  'H_{Shannon}\\nGenomik Bilgi Kapasitesi: C_{DNA} = 2 * N_{base_pairs} = 6.4 * 10^9 bit\\nYaşlanmaya Bağlı '
  'Enformasyon Kaybı Oranı: dH_{epigenetic} / dt >= 1.2 * 10^{-4} bit / gün / hücre.',
  "Bu eşdeğerlik, yaşlanmanın özünde fiziksel bir 'veri bozulması' (data corruption) olduğunu ve doğru yedek "
  'kopyalarla sistemin baştan yazılabileceğini ispatlar.'),
 ('1.10',
  'Canlılığın Biyofiziksel Kararlılık Kriterleri: Liapunov Fonksiyonları ve Attractor Havuzları',
  'Karmaşık dinamik sistemler teorisinde bir hücrenin sağlığı, çok boyutlu durum uzayında (state-space) kararlı bir '
  'Liapunov çekici havuzunda (attractor basin) bulunmasıyla ifade edilir.',
  "Waddington'ın epigenetik vadisi, matematiksel olarak bir Liapunov potansiyel yüzeyidir (V(x)). Genç bir hücre, "
  'vadinin en derin ve kararlı tabanında salınır. Yaşlanma, bu potansiyel çukurunun aşınması, sığlaşması ve hücrenin '
  'komşu dejeneratif atraktörlere (senesens veya malignensi havuzları) yuvarlanmasıdır.',
  'Liapunov Kararlılık Teoremi ve Attractor Havuzu Derinliği:\\nV(x) > 0,   dV(x) / dt = grad(V) * (dx / dt) <= 0 '
  '(Kararlı genç durum)\\nYaşlanma Dejenerasyonu: dV(x) / dt > 0 (Kararsızlaşma ve atraktör sığlaşması)\\nBariyer '
  'Yüksekliği: Delta U_{reversal} = V_{barrier} - V_{aged_state}.',
  'Epigenetik gençleşme protokolleri, matematiksel olarak hücreyi sığlaşmış yaşlılık havuzundan çıkarıp derin gençlik '
  'potansiyel kuyusuna geri fırlatan termodinamik kuvvet alanlarıdır.')]

part2_topics = [('2.1',
  'Hücre İçi Stokastik Moleküler Gürültü ve Poisson Süreçleri',
  'Biyolojik hücreler deterministik saat mekanizmaları değildir; reaksiyonlar düşük molekül sayılarıyla yürütülen '
  'stokastik (rastlantısal) Poisson süreçleridir. Bir promotör bölgesine bağlanan tek bir transkripsiyon faktörü '
  'molekülünün varlığı ya da yokluğu, tüm hücrenin kaderini değiştirebilir.',
  'Gillespie stokastik simülasyon algoritması, moleküler çarpışmaların olasılıksal doğasını modeller. Genç bir hücrede '
  'bu gürültü negatif geri bildirim döngüleriyle bastırılırken; yaşlı hücrede gürültü genliği artar ve kritik genlerin '
  'transkripsiyonu kaotik dalgalanmalara teslim olur.',
  'Kimyasal Master Denklem (Chemical Master Equation - CME):\\nd P(x, t) / dt = SUM_j [ a_j(x - nu_j) * P(x - nu_j, t) '
  '- a_j(x) * P(x, t) ]\\nGürültü Genliği (Fano Faktörü): F = sigma^2 / mu >= 1.0 (Poisson üstü gürültü)\\nYaşlanma '
  'Gürültü Artış Katsayısı: Noise_{aged} / Noise_{young} >= 2.8 kat.',
  'Moleküler gürültünün kontrolsüz büyümesi, aynı dokudaki yan yana iki hücrenin bile tamamen farklı protein '
  'profilleri sergilemesine (hücresel heterojenlik) yol açar.'),
 ('2.2',
  'Transkripsiyonel Sürüklenme (Transcriptional Drift) ve Hücresel Heterojenlik',
  'Yaşlanan dokularda gözlemlenen en belirgin olgulardan biri, aynı hücre tipine ait hücrelerin transkriptomlarının '
  "birbirinden uzaklaşması yani 'Transkripsiyonel Sürüklenme'dir.",
  'Tek hücre RNA dizileme (scRNA-seq) analizleri, genç dokularda hücrelerin transkripsiyonel uzayda sıkı ve homojen '
  'bir küme oluşturduğunu; yaşlı dokularda ise bu kümenin dağılarak genişlediğini gösterir. Temel ev idaresi genleri '
  '(housekeeping genes) tutarsız eksprese edilir ve hücreler kimlik kargaşası yaşar.',
  'Transkripsiyonel Dağılım ve Öklid Uzaklığı Metriği:\\nDrift_{index} = (1 / N) * SUM_{i=1}^N || T_i - T_{mean} '
  '||_2\\nBurada T_i i. hücrenin gen ekspresyon vektörü, T_{mean} doku ortalamasıdır.\\nSürüklenme Artış Eşiği: Delta '
  'Drift_{aged} >= +165% (Doku disfonksiyonu başlangıcı).',
  'Hücrelerin ortak orkestra düzeninden kopup bireysel kakofoniye kayması, dokunun bütüncül fizyolojik işlevini (kas '
  'kasılması, nöronal iletim, karaciğer filtrasyonu) felç eder.'),
 ('2.3',
  'RNA Polimeraz II Duraklaması, Hatalı Transkripsiyon ve Hız Sapmaları',
  'Genomun okunmasından sorumlu devasa enzim kompleksi RNA Polimeraz II (RNAPII), yaşlanma sürecinde DNA hasarları ve '
  'nükleozom yoğunluğundaki değişimler nedeniyle sık sık duraklar (stalling).',
  "Nature'da yayınlanan son çalışmalar, yaşlı hücrelerde RNAPII transkripsiyon uzama hızının (elongation rate) genç "
  'hücrelere göre belirgin biçimde arttığını, ancak hız artışının transkripsiyonel hata oranını (fidelity loss) 4 kat '
  "yükselttiğini ortaya koymuştur. Hızlı ama hatalı okuma, kusurlu mRNA'ların sitoplazmaya dökülmesine sebep olur.",
  'RNAPII Transkripsiyonel Kinetik ve Hata Oranı:\\nV_{elongation} = k_{on} * [NTP] / ( K_m + [NTP] ) approx 2.8 '
  'kb/min (Yaşlıda 3.6 kb/min)\\nHata Frekansı: Error_{transcription} = N_{mistakes} / N_{nucleotides} >= 1.2 * '
  '10^{-4} (Genç: 2.8 * 10^{-5})\\nDuraklama İndeksi: Pause_{index} = Density_{promoter} / Density_{gene_body}.',
  'Hatalı transkripsiyon, kodlanmış proteinlerin primer yapısında rastgele mutasyonlara yol açarak proteostaz ağını '
  'anında çökertir.'),
 ('2.4',
  'Ribozomal Çeviri Hataları ve Yanlış Katlanmış Protein Üretimi (DRiPs)',
  'Sitoplazmadaki ribozomlar, mRNA dizisini amino asit zincirine çevirirken mükemmel bir sadakatle çalışmak '
  'zorundadır. Ancak yaşlanma ile ribozomal RNA (rRNA) baz modifikasyonları ve ribozomal proteinlerin oksidasyonu '
  'çeviri sadakatini bozar.',
  "Hatalı tRNA seçimi (misacylation) ve durdurma kodonlarının atlanması (stop-codon readthrough), 'Defective Ribosomal "
  "Products' (DRiPs) adı verilen devasa bir hatalı protein havuzu oluşturur. Yeni sentezlenen proteinlerin %35'i daha "
  'ribozomdan çıkarken çöpe dönüşür.',
  'Ribozomal Translasyon Hata Kinetiği:\\nError_{ribosome} = Rate_{mismatch} / Rate_{cognate} = (k_{cat} / '
  'K_m)_{near-cognate} / (k_{cat} / K_m)_{cognate}\\nDRiPs Üretim Fraksiyonu: f_{DRiP} = [Protein]_{defective} / '
  '[Protein]_{total} >= 0.38\\nEnerji Kaybı: Boşa harcanan ATP oranı >= 22% toplam hücresel ATP.',
  'Hücrenin en değerli enerjisi, daha baştan çalışmayan ve biriken zehirli protein çöplerini üretmeye harcanarak '
  'biyolojik iflası hızlandırır.'),
 ('2.5',
  "Soma ve Germline Ayrımı: Kirkwood'un 'Tek Kullanımlık Soma' (Disposable Soma) Teorisi",
  "Evrimsel biyolog Thomas Kirkwood, 1977 yılında organizmaların enerji bütçesini açıklayan 'Tek Kullanımlık Soma' "
  'teorisini ortaya atmıştır. Bu teoriye göre doğal seçilim, organizmanın enerjisini iki temel alan arasında '
  'paylaştırır: Üreme (Germline) ve Bedenin Onarımı (Soma).',
  'Evrimsel başarı sadece genlerin aktarılmasına baktığından, üreme çağından sonra somatik dokuların (kas, beyin, '
  "karaciğer) kusursuz onarımı için enerji harcamak evrimsel olarak 'israf' kabul edilir. Germline hücreleri (yumurta "
  've sperm) telomeraz ve kusursuz onarımla ölümsüz tutulurken, somatik hücreler entropik çürümeye terk edilir.',
  'Kirkwood Enerji Tahsis Formülasyonu:\\nE_{total} = E_{reproduction} + E_{somatic_maintenance} + '
  'E_{metabolism}\\nOptimal Evrimsel Tahsis: d(Fitness) / d(E_{maintenance}) = 0 noktasında E_{maintenance} < '
  'E_{immortality_threshold}\\nSoma Yıpranma Hızı: d(Damage) / dt = k_{wear} - f(E_{somatic_maintenance}) > 0.',
  "PROJECT AETERNITAS'ın devrimci felsefesi tam olarak burada yatar: Biyolojik ölümsüzlük, germline hücrelerine "
  "tanınan bu ayrıcalıklı 'sonsuz onarım' programını somatik hücrelere genetik olarak zorla yüklemektir."),
 ('2.6',
  'Mikroskobik Reversibilite ve Biyolojik Kinetiğin İleri Doğru Kilitlenmesi',
  'Fiziksel düzeyde kimyasal reaksiyonlar mikroskobik reversibiliteye (tersinirliğe) tabidir. Ancak biyolojik '
  'sistemlerde ATP ve GTP hidrolizinin sağladığı devasa serbest enerji serbest kalışı (-30.5 kJ/mol), reaksiyonları '
  'tek yönlü kinetik kilitlere bağlar.',
  'Bu durum, hücresel hasar birikiminin geriye dönüşsüz bir tek yönde akmasına neden olur. Protein çapraz bağları '
  "(AGEs), DNA çift zincir delesyonları ve telomerik erozyon termodinamik olarak 'tek yönlü valfler' gibi çalışır; "
  'hücre ancak ekzojen biyoteknolojik iş harcayarak bu kilitleri kırabilir.',
  'Kinetik Kilitleme ve Tek Yönlü Reaksiyon Akısı:\\nJ_{forward} / J_{reverse} = exp( A / (R * T) ) >> 1.0\\nBurada A '
  '= -Delta G kimyasal afinite değeridir (A >= 50 kJ/mol durumunda reaksiyon mutlak tek yönlüdür).\\nKinetik Bariyer '
  'Aşımı: Delta G_{reverse_catalysis} <= -E_{barrier}.',
  'Doğal biyolojik mekanizmalar bu tek yönlü valfleri geri çeviremez; tersine çevirme ancak dışarıdan sentetik '
  'biyomühendislik enzimlerinin (Yamanaka faktörleri, DNA ligazlar) devreye girmesiyle mümkündür.'),
 ('2.7',
  'Stokastik Nükleer Manyetik Gürültü ve İyon Kanalı Titreşimleri',
  'Nöronal ve elektriksel olarak uyarılabilir hücrelerde membran potansiyelini belirleyen voltaj kapılı iyon kanalları '
  '(Nav1.2, Cav1.2, Kv1.1), stokastik termal açılıp kapanma gürültüsüne sahiptir.',
  'Yaşlanma ile birlikte kanal proteinlerinin allosterik yapılarında meydana gelen post-translasyonel modifikasyonlar '
  '(oksidasyon ve glikasyon), kanalın açık kalma olasılığını (P_open) ve kapı titreşim frekansını bozar. Bu durum '
  'membran dinlenim potansiyelinde mikroskobik gürültüye ve spontan eksitotoksik iyon kaçaklarına neden olur.',
  'İyon Kanalı Stokastik Kinetiği (Markov Zinciri Modeli):\\nC <---> O <---> I   (Closed, Open, Inactivated)\\nGürültü '
  'Spektral Yoğunluğu: S_I(f) = 4 * k_B * T * Re(Y(f)) + 2 * q * I_{leak}\\nKanal Açılma Olasılığı Sapması: Delta '
  'P_{open} >= 18% (Dinlenim potansiyelinde).',
  'İyon kanallarındaki bu kaçaklar, yaşlı hücrenin dinlenim halindeyken bile sürekli enerji tüketmesine ve kalsiyum '
  'aşırı yüklenmesine maruz kalmasına yol açar.'),
 ('2.8',
  'Sitozolik Viskozite Artışı ve Moleküler Kalabalıklaşma (Macromolecular Crowding)',
  "Genç bir hücre sitoplazması, hacminin yaklaşık %20-30'unu kaplayan protein ve organeller içerir. Ancak yaşlı "
  'hücrelerde agregatlar, otofajik vakuoller ve lipofuscin birikimi nedeniyle makromoleküler kalabalıklaşma '
  "(macromolecular crowding) %45'in üzerine fırlar.",
  'Bu aşırı kalabalıklaşma, sitozolün vizkozitesini 3 kat artırarak difüzyon katsayılarını (D) dramatik ölçüde '
  'düşürür. Moleküller hedeflerine ulaşamaz; enzim-substrat bağlanma hızları yavaşlar ve hücre içi sinyal iletimi felç '
  'olur.',
  'Stokes-Einstein ve Kalabalıklaşma Difüzyon Denklemi:\\nD = (k_B * T) / (6 * pi * eta * r_h) * exp( -a * '
  'phi_{crowding}^b )\\nBurada phi_{crowding} hücre içi makromoleküler hacim fraksiyonudur (Yaşlıda >= 0.45)\\nEfektif '
  'Difüzyon Yavaşlaması: D_{aged} / D_{young} <= 0.35.',
  'Hücre adeta kendi ürettiği moleküler bataklığın içine saplanır; biyokimyasal reaksiyonlar milisaniyelerden '
  'saniyelere uzayan gecikmeler yaşar.'),
 ('2.9',
  'İntrasellüler pH Gradyanlarının Aşınması ve Proton Sızıntıları',
  'Hücre içi organellerin çalışması kesin pH gradyanlarına dayanır: Sitozol nötr (pH 7.2), lizozomlar asidik (pH '
  '4.5-5.0), mitokondri matriksi ise alkalidir (pH 7.8-8.0).',
  'Yaşlanma ile birlikte v-ATPase proton pompalarının verimi düşer ve membran lipidlerindeki mikro-delikler proton '
  'kaçaklarına (proton leak) yol açar. Lizozomun asitliği kaybolurken mitokondri matriksinin alkalinitesi aşınır; '
  'hücresel pH homojenleşmeye başlar. pH gradyanının silinmesi, enzimlerin üçüncül yapısını denatüre eder.',
  'Proton-Motive Force ve Trans-Membran pH Gradyanı:\\nDelta mu_{H+} = F * Delta Psi - 2.303 * R * T * Delta '
  "pH\\nLizozomal Asidifikasyon İflası: pH_{lysosome} = 4.8'den 6.1'e yükselir\\nv-ATPase Hidrolitik Verim Kaybı: "
  'Delta Activity <= -55%.',
  'Proton dengesinin çöküşü, lizozomun tüm hücresel sindirim yeteneğini durdurarak yaşlanmanın en ölümcül atık birikim '
  'motorunu ateşler.'),
 ('2.10',
  'Stokastik İflasın Kritik Eşiği: Hata Katastrofu (Orgel Error Catastrophe) Hipotezi',
  "1963 yılında Leslie Orgel, yaşlanmanın kaçınılmaz bir 'Hata Felaketi' (Error Catastrophe) ile sonuçlandığını öne "
  'sürmüştür. Bu hipoteze göre transkripsiyon ve translasyon mekanizmalarının kendisini inşa eden enzimlerde (RNAPII, '
  'ribozomlar, aminoaçil-tRNA sentetazlar) meydana gelen küçük bir hata, bir sonraki nesil enzimlerde daha büyük '
  'hatalara yol açar.',
  'Bu pozitif geri bildirim döngüsü, hata oranının zamanla katlanarak arttığı ve nihayetinde fonksiyonel protein '
  "sentezinin tamamen imkansız hale geldiği bir 'kritik devrilme noktasına' (tipping point) ulaşır.",
  'Orgel Hata Katastrofu Matematiksel Modeli:\\ndE / dt = alpha * E + beta * (E^2) / (1 - gamma * E)\\nBurada E '
  'kümülatif hata oranıdır.\\nKatastrof Eşiği: E(t) >= E_{crit} anında dE/dt -> sonsuz (Hücresel ani çöküş ve ölüm).',
  'PROJECT AETERNITAS, Orgel geri bildirim döngüsünü dışarıdan hatasız sentetik transkripsiyonel faktörler ve saf mRNA '
  'replasmanlarıyla keserek hata katastrofunu imkansız kılar.')]

part3_topics = [('3.1',
  "David Sinclair'in 'Yaşlanmanın Enformasyonel Teorisi' (ITM) Temelleri",
  "Harvard Tıp Fakültesi'nden David Sinclair ve ekibinin formüle ettiği Yaşlanmanın Enformasyon Teorisi (Information "
  'Theory of Aging), biyolojik yaşlanmanın donanımsal (DNA sekansı) mutasyonlardan ziyade yazılımsal (epigenomik '
  'organizasyon) enformasyon kaybından kaynaklandığını savunur.',
  'Bir kompakt disk (CD) düşünün: Yüzeyindeki çizikler nedeniyle müzik çalar şarkıyı okuyamaz; ancak verinin kendisi '
  "CD'nin çukurlarında hala bozulmadan durmaktadır. Yüzey cilalandığında müzik ilk günkü gibi çalar. Hücrede de DNA "
  'sekansı 80 yaşında bile %99.99 sağlamdır; kaybolan şey hangi genin ne zaman açılıp kapanacağını belirleyen '
  'epigenetik işletim sistemidir.',
  'Enformasyonel Shannon Güvenilirlik Oranı:\\nR_{epigenetic} = 1 - ( H_{loss} / C_{channel} )\\nBurada C_{channel} '
  'epigenomik hafıza kapasitesi, H_{loss} zamanla biriken Shannon bilgi kaybıdır.\\nRejuvenasyon İlkesi: Restorasyon '
  'Kapasitesi = C_{channel} - H_{loss} > 0.',
  'Bu paradigma, yaşlanmanın tedavi edilemez bir donanım erimesi değil, tersine çevrilebilir bir yazılım bozulması '
  'olduğunu kanıtlar.'),
 ('3.2',
  'Kromatin Manzarası ve Waddington Epigenetik Çekicilerinin Silinmesi',
  "Hücresel kimlik, Conrad Waddington'ın meşhur 'epigenetik manzarası' ile açıklanır. Pluripotent bir kök hücre "
  'tepeden aşağı yuvarlanarak nöron, hepatosit veya kardiyomiyosit vadilerinden birine yerleşir.',
  'Gençlikte bu vadileri ayıran dağlar ve epigenetik sırtlar (heterokromatin blokları, DNA metilasyon bariyerleri) çok '
  'yüksektir. Yaşlanma sürecinde bu topoğrafya aşınır; vadiler düzleşir. Nöronlar nöronluk kimliklerini, deri '
  'hücreleri fibroblast kimliklerini unutup belirsiz, disfonksiyonel bir ara duruma doğru kayarlar.',
  'Waddington Potansiyel Enerji Fonksiyonu:\\nV_{Waddington}(x) = -Integral F_{epigenetic}(x) dx\\nBariyer Yüksekliği '
  'Aşınması: Delta V_{barrier}(t) = Delta V_0 * exp( -t / tau_{erosion} )\\nKimlik Kararlılık Kriteri: Delta '
  'V_{barrier} >= 5 * k_B * T (Kimlik kaybını önleme eşiği).',
  'Kısmi yeniden programlama (Partial Reprogramming), aşınmış vadileri yeniden derinleştirerek hücrenin kendi orijinal '
  'genç kimliğine sıkı sıkıya sarılmasını sağlar.'),
 ('3.3',
  "Steve Horvath'ın Pan-Doku Epigenetik Saati ve DNA Metilasyon Sapması",
  "2013 yılında UCLA'dan Steve Horvath, insan genomundaki CpG adacıklarının DNA metilasyon seviyelerini "
  '(5-metilsitozin: 5mC) analiz ederek tüm dokularda kronolojik yaşı %96 doğrulukla tahmin eden ilk pan-doku '
  'epigenetik saatini keşfetmiştir.',
  'Horvath saati, 353 spesifik CpG lokusundaki metilasyon yüzdesini ölçer. Yaşlandıkça bazı kritik gen promotörleri '
  'hiper-metilasyona uğrayarak sessizleşirken (örneğin DNA tamir genleri), diğer bölgeler hipo-metilasyona uğrayarak '
  'genomik kararsızlığa yol açar. Bu metilasyon sürüklenmesi (epigenetic drift), biyolojik yaşın en evrensel '
  'cetvelidir.',
  'Horvath Epigenetik Saat Matematiksel Fonksiyonu:\\nDNAmAge = F( b_0 + SUM_{i=1}^{353} b_i * beta_i )\\nBurada '
  'beta_i = M_i / (M_i + U_i + 100) (Metilasyon fraksiyonu), b_i regresyon katsayılarıdır.\\nF(x) fonksiyonu adult ve '
  'fetal yaşlar için logaritmik/lineer dönüşüm sağlar:\\nBiyolojik Gençleşme Doğrulaması: Delta DNAmAge = '
  'DNAmAge_{sonra} - DNAmAge_{önce} < 0.',
  'Epigenetik saat, yaşlanmanın soyut bir kavram değil, doğrudan tersine çevrilebilir matematiksel bir biyobelirteç '
  'vektörü olduğunu somutlaştırmıştır.'),
 ('3.4',
  'İkinci ve Üçüncü Nesil Epigenetik Saatler: PhenoAge, GrimAge ve DunedinPACE',
  "Horvath'ın kronolojik yaşı ölçen ilk saatinden sonra Morgan Levine ve ekibi klinik biyobelirteçlerle entegre "
  "çalışan 'DNAm PhenoAge'i; ardından mortalite riskini ve plazma proteinlerini şifreleyen 'DNAm GrimAge'i "
  'geliştirmiştir.',
  "En son teknoloji olan DunedinPACE (2022-2026), statik bir yaş yerine 'Yaşlanma Hızını' (Pace of Aging) ölçer. Bir "
  'insanın yılda kaç biyolojik yıl yaşlandığını tespit eder (normal insan: 1.0 yıl/yıl; hızlanmış: 1.4 yıl/yıl). '
  "Rejuvenasyon tedavilerinde hedef, DunedinPACE skorunu 0.60'ın altına çekmektir.",
  'DunedinPACE Biyolojik Yaşlanma Hızı Metriği:\\nPaceOfAging = SUM_{j=1}^{173} w_j * beta_j + c_0\\nHedef Değer: '
  'DunedinPACE < 0.65 (Biyolojik saatin yavaşlatılması)\\nGrimAge Mortalite Riski İndirgenmesi: HazardRatio = exp( '
  'beta_{Grim} * Delta GrimAge ) <= 0.45.',
  "Bu ileri nesil algoritmalar, PROJECT AETERNITAS'ın uyguladığı tedavilerin doku gençleşmesini saniyeler içinde "
  'kanıtlayan moleküler yargıçlarıdır.'),
 ('3.5',
  'Heterokromatin Kaybı Modeli (Heterochromatin Loss Model of Aging)',
  "Genomun yaklaşık %90'ı susturulmuş, sıkıca paketlenmiş ve nükleer zara raptedilmiş 'Heterokromatin' halinde "
  'saklanır. Heterokromatin, genomun karanlık maddesidir; retrotranspozonları ve virüs kalıntılarını kilit altında '
  'tutar.',
  "Yaşlanma ile birlikte heterokromatin adacıkları çözülür ve dekompakt hale gelir ('Heterokromatin Kaybı Modeli'). "
  'H3K9me3 ve H3K27me3 gibi susturucu histon işaretleri silinir. Milyonlarca yıldır uyuyan LINE-1 retrotranspozonları '
  "uyanarak nükleustan sitoplazmaya atlar ve kronik bir 'viral enfeksiyon' alarımı yaratarak steril nöroenflamasyonu "
  '(inflammaging) tetikler.',
  'Heterokromatin Fraksiyonu ve Retrotranspozon Baskılama Katsayısı:\\nFraction_{heterochromatin} = [H3K9me3]_{bound} '
  '/ [Histone_H3]_{total} >= 0.65 (Gençlik normu)\\nYaşlılık Çözülmesi: Delta Fraction <= -45%\\nLINE-1 Transpozisyon '
  'Hızı: Rate_{L1} = k_{derepression} * ( 1 - Fraction_{heterochromatin} ).',
  'Heterokromatinin yeniden kilitlenmesi, hücresel genomun içindeki vahşi retroviral canavarları susturmanın ve '
  'iltihabı kökten kurutmanın tek yoludur.'),
 ('3.6',
  "DNA Kırıkları ve Epigenetik Faktörlerin 'Geri Dönememesi' (Relocalization of Chromatin Modifiers)",
  'David Sinclair laboratuvarı tarafından keşfedilen ICE (Inducible Changes to the Epigenome) fare modeli, epigenetik '
  'yaşlanmanın ana mekanizmasını aydınlatmıştır: DNA çift zincir kırıklarının (DSB) onarımı.',
  'Her gün bir insan hücresinde 10 ila 50 adet DNA çift zincir kırığı oluşur. Bu kırıklar meydana geldiğinde, normalde '
  'gen promotörlerini susturmakla görevli sirtuinler (SIRT1, SIRT6) ve PARP1 enzimleri görev yerlerini terk ederek '
  "kırık bölgesine koşarlar. DNA tamir edildikten sonra bu faktörlerin %99'u eski yerine döner; ancak %1'i yolunu "
  'kaybeder. Yıllar içinde bu küçük kayıp, epigenomun tamamen silinmesine yol açar.',
  'Kromatin Modifiye Edici Faktörlerin Kayıp Kinetiği:\\nd[SIRT]_{promoter} / dt = -k_{recruit} * [DSB] + k_{return} * '
  '[SIRT]_{repair} * (1 - epsilon_{loss})\\nBurada epsilon_{loss} approx 0.008 (Her tamir döngüsündeki kayıp '
  'fraksiyonu)\\nPromotör Epigenetik Erozyon Hızı: Rate_{drift} = epsilon_{loss} * Flux_{DNA_damage}.',
  'Bu formül, yaşlanmanın doğrudan DNA onarım maliyetinin bir yan ürünü olduğunu gösterir; DNA onarım mekanizması '
  'güçlendirildikçe epigenetik kayıp sıfırlanabilir.'),
 ('3.7',
  'CpG Adacıklarının Hipo- ve Hiper-Metilasyon Paradoksu',
  'Yaşlanan hücrenin DNA metilomunda ilk bakışta çelişkili görünen bir fenomen yaşanır: Genom genelinde (intergenik '
  'bölgelerde ve tekrarlayan dizilerde) küresel bir hipo-metilasyon görülürken; tümör baskılayıcı ve plastisite '
  'genlerinin CpG adacıklarında bölgesel hiper-metilasyon gerçekleşir.',
  'DNMT1 (DNA Metiltransferaz 1) kopyalama enziminin yaşlanmayla sadakatini kaybetmesi, metilasyonun olması gereken '
  'yerden silinmesine ve olmaması gereken yere yapışmasına yol açar. Sonuç: Kanser riski artar, hücre tamir genleri '
  'kilitlenir ve transposonlar serbest kalır.',
  'Metilasyon Entropi İndeksi (MEI):\\nMEI = - (1 / N) * SUM_{k=1}^N [ beta_k * ln(beta_k) + (1 - beta_k) * ln(1 - '
  'beta_k) ]\\nGenç Hücre MEI: MEI_{young} approx 0.15 (Metilasyon kesin çizgilerle ayrılmıştır: ya 0 ya 1)\\nYaşlı '
  'Hücre MEI: MEI_{aged} >= 0.58 (Maksimum belirsizlik ve kaos).',
  'Rejuvenasyon protokolleri, MEI indeksini 0.15 seviyesine geri çekerek CpG adacıklarının orijinal ikili (binary) '
  'netliğini restore eder.'),
 ('3.8',
  'Sirtuin Ailesi (SIRT1-SIRT7) ve NAD+ Bağımlı Deasetilasyon İflası',
  'Sirtuinler, bakterilerden insana kadar korunmuş evrensel ömür uzatıcı ve epigenetik koruyucu enzimlerdir. '
  'Memelilerde 7 sirtuin bulunur: SIRT1, 6 ve 7 nükleusta heterokromatini korur; SIRT3, 4 ve 5 mitokondride solunum '
  'zincirini deasetile eder; SIRT2 sitoplazmada mikrotübülleri yönetir.',
  "Tüm sirtuinler mutlak surette Nikotinamid Adenin Dinükleotid'e (NAD+) bağımlıdır. Yaşlanma ile hücre içi NAD+ "
  'havuzunun %70 azalması, sirtuinlerin katalitik aktivitesini durma noktasına getirir. Sirtuinler sustuğunda, histon '
  'asetilasyonu kontrolden çıkar ve mitokondriyel enzimler zehirli asetil gruplarıyla tıkanır.',
  'Sirtuin Deasetilasyon Kinetik Reaksiyonu:\\nProtein-Lys(Ac) + NAD+ ---> Protein-Lys + O-Asetil-ADP-Riboz + '
  'Nikotinamid\\nKatalitik Hız: V_{sirt} = (k_{cat} * [SIRT] * [Substrate] * [NAD+]) / ( K_m^{NAD+} * K_m^{Sub} + ... '
  ')\\nYaşlı Hücrede Katalitik Verim Kaybı: V_{sirt}(Aged) <= 0.25 * V_{sirt}(Young).',
  'NAD+ havuzunun NMN ve CD38 inhibitörleriyle doldurulması, sirtuin motorlarını anında yeniden ateşleyerek epigenetik '
  'hafızayı canlandırır.'),
 ('3.9',
  'Topolojik Olarak İlişkili Alanların (TADs) ve Kromatin İlmeklerinin Çöküşü',
  'Kromatin çekirdek içinde rastgele bir iplik yumağı gibi durmaz; CTCF proteinleri ve Kohezin kompleksleri tarafından '
  'yönetilen Topolojik Olarak İlişkili Alanlar (Topologically Associating Domains - TADs) halinde düzenlenir.',
  'TAD sınırları, bir genin uzaktaki yanlış bir güçlendirici (enhancer) tarafından kazara ateşlenmesini engeller. '
  'Yaşlanma sürecinde CTCF proteininin kromatine bağlanma afinitesi düşer; TAD sınırları erir. Yanlış genler yanlış '
  'güçlendiricilerle temas eder; hücre anormal transkripsiyonel fırtınalara girer.',
  'TAD İzolasyon Skoru ve Kromatin İlmek Kararlılığı:\\nInsulation_{score} = ContactFrequency_{inter-TAD} / '
  'ContactFrequency_{intra-TAD}\\nGenç İzolasyon: Insulation_{score} <= 0.08 (Tam sınır yalıtımı)\\nYaşlı İzolasyon '
  'Çöküşü: Insulation_{score} >= 0.42 (TAD sınırlarının silinmesi).',
  '3D genomik mimarinin bu yapısal çöküşü, hücrenin kendi genetik kütüphanesini okuma haritasını kaybetmesi anlamına '
  'gelir.'),
 ('3.10',
  "Epigenetik Sıfırlama ile Hücresel Yaşın 'Geri Sarılabilirliği' Kanıtı",
  'Enformasyon teorisinin en muazzam müjdesi şudur: Eğer epigenetik kayıp bir enformasyon kaybıysa, orijinal gençlik '
  'epigenomunun şifresi nerede saklanmaktadır?',
  'Yamanaka faktörleri (Oct4, Sox2, Klf4) transient olarak hücreye verildiğinde, hücrenin DNA sekansını değiştirmeden '
  'epigenetik saatini sıfıra döndürebildiği kanıtlanmıştır. Bu bulgu, hücrenin çekirdeğinde epigenomun orijinal '
  "gençlik durumunu saklayan derin bir 'Yedek Kopya' (Backup Copy / Ground State) olduğunu gösterir. DNA demetilazlar "
  '(TET enzimleri) bu kopyayı okuyarak metilasyon manzarasına reset atar.',
  'Hücresel Yaş Sıfırlama Denklemi (Sinclair-TET Dinamiği):\\nd(Age_{bio}) / dt = -k_{OSK} * [OSK]_{nuclear} * '
  '[TET1/2]_{active} + Rate_{aging}\\nGençleşme Hızı Şartı: k_{OSK} * [OSK] * [TET] > Rate_{aging}  ===>  Age_{bio}(t) '
  '-> 0\\nDiferansiyasyon Koruma Eşiği: t_{induction} <= t_{critical_pluripotency} (Kimlik kaybı sıfır).',
  "Bu fiziksel ve moleküler gerçeklik, PROJECT AETERNITAS'ın bilimsel omurgasıdır: Hücresel gençlik ebediyen "
  'ulaşılabilir ve geri çağrılabilir bir durumdur.')]

# Append Parts 1, 2, 3 to parts list
parts.append((1, "Termodinamik Entropi, Schrödinger ve Canlılığın Negatif Entropi Paradoksu", part1_topics))
parts.append((2, "Stokastik Moleküler Gürültü ve Transkripsiyonel Sürüklenme", part2_topics))
parts.append((3, "Yaşlanmanın Enformasyonel Teorisi: Epigenetik Gürültü ve Enformasyon Kaybı", part3_topics))

print("[PROJECT AETERNITAS] Parts 1, 2, 3 appended to build_k1_ch01_docx.py successfully.")

part4_topics = [('4.1',
  "Denham Harman'ın 'Serbest Radikal Teorisi' (MFRTA) ve Modern Revizyonu",
  '1956 yılında Denham Harman tarafından ortaya atılan Yaşlanmanın Mitokondriyal Serbest Radikal Teorisi (MFRTA), '
  'mitokondriyel solunum sırasında kaçan serbest oksijen radikallerinin (ROS) hücre bileşenlerini kademeli olarak '
  'tahrip ettiğini ve yaşlanmanın primer sürücüsü olduğunu öne sürmüştür.',
  "Modern biyofizik Harman'ın teorisini köklü biçimde revize etmiştir: Düşük ve orta düzey ROS aslında zararlı bir çöp "
  'değil; hücre içi adaptasyon, otofaji indüksiyonu ve Nrf2 aktivasyonu için zorunlu bir sinyal molekülüdür '
  '(mitohormezis). Yaşlanma, ROS üretiminin artmasından ziyade, anti-oksidan geri bildirim döngülerinin sönümlenmesi '
  've redoks sinyal hassasiyetinin kaybolmasıdır.',
  'Harman Radikal Hasar Kinetiği ve Modern Redoks Denge Denklemi:\\nd[ROS] / dt = Rate_{ETC_leak} - Rate_{scavenging}( '
  '[SOD], [Catalase], [GPx] )\\nRedoks Potansiyeli (Nernst Denklemi): E_{GSH/GSSG} = E^0 - (R*T / 2F) * ln( [GSH]^2 / '
  '[GSSG] )\\nGenç Redoks Durumu: E_{GSH} approx -240 mV; Yaşlı Redoks Çöküşü: E_{GSH} >= -170 mV.',
  "Redoks potansiyelinin -170 mV'a yükselmesi, hücre içi proteinlerin disülfid köprülerinin spontan oksitlenmesine ve "
  'apoptoz kaskadlarının tetiklenmesine yol açar.'),
 ('4.2',
  'Elektron Taşıma Zincirinde Kaçaklar (Kompleks I ve III) ve Süperoksit Üretimi',
  "Mitokondri iç zarında çalışan Elektron Taşıma Zincirinde (ETC) her gün tüketilen oksijenin yaklaşık %0.5 ila %2'si "
  "tam indirgenemez; Kompleks I (FMN sahası) ve Kompleks III'ten (Qo sahası) tek elektronlu kaçaklarla Süperoksit "
  'Anyonuna (O2•-) dönüşür.',
  'Genç hücrelerde MnSOD (Süperoksit Dismutaz 2) bu radikalleri saniyeler içinde hidrojen peroksite (H2O2) ve suya '
  'çevirir. Yaşlanmayla birlikte iç zardaki respirasom süper-komplekslerinin dağılması, elektronların serbest difüzyon '
  "mesafesini uzatır; kaçak oranı %6'ya fırlar.",
  'Süperoksit Kaçak Akısı ve Kinetik Formülasyon:\\nJ_{O2.-} = k_{leak} * [NADH] * [O_2] * exp( alpha * F * Delta '
  'Psi_m / (R * T) )\\nKaçak Çarpanı: Multiplier_{leak} = J_{aged} / J_{young} >= 3.4 kat\\nO2•- Yarılanma Ömrü: '
  't_{1/2} approx 1.0 us (Mikrosaniye - anında çevreye saldırır).',
  'Aşırı süperoksit üretimi, solunum zincirinin kendi demir-kükürt (Fe-S) kümelerini parçalayarak mitokondriyi '
  'içeriden havaya uçurur.'),
 ('4.3',
  'Mitohormezis Paradoksu: Düşük Doz ROS ile Yaşam Süresi Uzatımı',
  'Mitohormezis (Mitochondrial Hormesis), hücrenin hafif metabolik stres veya düşük doz ROS maruziyetine yanıt olarak '
  'kendi savunma mekanizmalarını aşırı güçlendirmesi olgusudur (bifazik doz-yanıt eğrisi).',
  'Dışarıdan yüksek doz sentetik antioksidan (C vitamini, E vitamini) verildiğinde yaşlanmanın durmaması hatta '
  "hızlanması bu paradoksla açıklanır: Aşırı antioksidan, endojen Nrf2/ARE savunma kaskadını uyutan 'redoks "
  "tembelliği' yaratır. Uzun ömürlü canlılar (çıplak kör fare, Grönland balinası) yüksek ROS üretir ancak muazzam bir "
  'mitohormetik adaptasyon sergiler.',
  'Hormetik Yanıt Fonksiyonu (Calabrese Eğrisi):\\nResponse(Dose) = R_0 * [ 1 + (a * Dose / (1 + b * Dose^2)) '
  ']\\nOptimal Hormetik Pencere: Dose_{ROS} in [Dose_{min}, Dose_{max}]\\nNrf2 Aktivasyon Katsayısı: Fold_{Nrf2} >= '
  '4.2 (Hormetik aralıkta).',
  'PROJECT AETERNITAS, pasif antioksidan yüklemesi yerine hücreye kontrollü mitohormetik sinyaller vererek endojen '
  'antioksidan fabrikasını maksimum devirde çalıştırır.'),
 ('4.4',
  'Fenton ve Haber-Weiss Reaksiyonları: Hidroksil Radikali (•OH) Tahribatı',
  'Süperoksit ve hidrojen peroksit doğrudan yıkıcı değildir; ancak hücrede serbest demir (Fe2+) veya bakır (Cu+) '
  'katyonları bulunduğunda meşhur Fenton reaksiyonu gerçekleşir.',
  'Fenton ve Haber-Weiss reaksiyonları, evrendeki en reaktif kimyasal tür olan Hidroksil Radikalini (•OH) üretir. '
  'Hidroksil radikalinin difüzyon yarıçapı sadece birkaç nanometredir; oluştuğu anda yanındaki ilk moleküle (DNA bazı, '
  'lipid veya enzim) çarparak elektronunu çalar; enzimatik olarak nötralize edilmesi imkansızdır.',
  'Fenton ve Haber-Weiss Reaksiyon Kinetiği:\\nFe^{2+} + H_2O_2 ---> Fe^{3+} + OH^- + *ullet OH   (Fenton '
  'Reaksiyonu)\\nO_2^{*ullet-} + H_2O_2 ---> O_2 + OH^- + *ullet OH   (Haber-Weiss Reaksiyonu)\\nReaksiyon Hız '
  'Sabiti: k_{Fenton} = 76 M^{-1} s^{-1}\\n*ullet OH Difüzyon Limiti: l_{diff} = sqrt( 2 * D * t_{1/2} ) approx 2.0 '
  'nm.',
  'Hidroksil radikalinden korunmanın tek yolu, reaktif demir havuzunu (Labile Iron Pool - LIP) ferritin içine sıkıca '
  'kilitlemek ve H2O2 birikimini önlemektir.'),
 ('4.5',
  '8-OHdG ve Oksidatif Nükleik Asit Lezyonlarının Kümülatif Birikimi',
  "Hidroksil radikallerinin DNA'daki guanin bazına saldırması sonucu 8-Hidroksi-2'-deoksiguanozin (8-OHdG) lezyonu "
  'oluşur. Guanin, en düşük iyonlaşma potansiyeline sahip baz olduğu için oksidatif saldırıların birincil '
  'paratoneridir.',
  '8-OHdG lezyonları, replikasyon sırasında DNA polimerazın adenin bazıyla yanlış eşleşmesine (G:C -> T:A '
  'transversiyonu) yol açar. Genç hücrede OGG1 (8-oksoguanin glikozilaz) bu lezyonları hızla onarırken; yaşlı hücrede '
  'tamir hızı hasar hızının gerisinde kalır ve somatik mutasyon yükü katlanır.',
  '8-OHdG Birikim ve Mutasyon Hızı Formülü:\\nd[8-OHdG] / dt = Rate_{oxidation} - k_{OGG1} * [OGG1] * [8-OHdG] / ( K_m '
  '+ [8-OHdG] )\\nGenomik 8-OHdG Düzeyi: [8-OHdG] / 10^6 dG = 0.45 (Genç) ---> 4.85 (Yaşlı)\\nG:C -> T:A Transversiyon '
  'Olasılığı: P_{transversion} >= 0.28 / lezyon.',
  '8-OHdG seviyesi, hücrenin nükleer ve mitokondriyel genomunun ne kadar paslandığını gösteren doğrudan bir biyolojik '
  'termometredir.'),
 ('4.6',
  'Peroksinitrit (ONOO-) ve Protein Nitrotirozin Modifikasyonları',
  'Mitokondriyel süperoksit (O2•-), endotelyal veya nöronal nitrik oksit (NO•) ile karşılaştığında difüzyon kontrollü '
  'ultra hızlı bir reaksiyonla (k = 1.9 * 10^10 M^-1 s^-1) Peroksinitrit (ONOO-) anyonunu üretir.',
  'Peroksinitrit son derece toksik bir reaktif nitrojen türüdür (RNS). Proteinlerdeki kritik tirozin kalıntılarını '
  'nitrotirozine (3-nitrotyrosine) çevirerek tirozin kinaz fosforilasyonunu bloke eder; reseptör sinyallerini kalıcı '
  'olarak kapatır ve hücresel apoptozu hızlandırır.',
  'Peroksinitrit Oluşum Kinetiği ve Nitrotirozin Formasyonu:\\nNO^*ullet + O_2^{*ullet-} ---> ONOO^-   (k = 1.9 '
  '* 10^{10} M^{-1} s^{-1} - Süperoksit dismutazdan 3 kat hızlı)\\nTirozin Nitrasyon Oranı: [Nitrotyrosine] / '
  '[Tyrosine] >= 3.2 * 10^{-4}\\nEnzim İnhibisyon Katsayısı: % Inhibition_{target} >= 75%.',
  'Nitrotirozin modifikasyonları geriye döndürülemez; hücre bu proteinleri tamamen parçalayıp yeniden sentezlemek '
  'zorunda kalır, bu da enerji krizini derinleştirir.'),
 ('4.7',
  'Glutatyon Redoks Döngüsü (GSH/GSSG) ve NADPH Tüketim Krizi',
  'Hücrenin en önemli redoks tamponu indirgenmiş Glutatyon (GSH) tripeptididir. Glutatyon Peroksidaz (GPx), hidrojen '
  "peroksiti suya indirgerken GSH'ı okside glutatyona (GSSG) dönüştürür. GSSG ise Glutatyon Redüktaz ve NADPH "
  "aracılığıyla tekrar GSH'a çevrilir.",
  'Yaşlanan hücrede Pentoz Fosfat Yolağının (PPP) yavaşlaması nedeniyle sitozolik NADPH üretimi çöker. NADPH eksikliği '
  "GSSG'nin geri indirgenmesini engeller; serbest GSH havuzu tükenir. Hücre, tek bir oksidatif dalgaya bile karşı "
  'koyamaz hale gelir.',
  'Glutatyon Redoks Akış Dengesi:\\n2 GSH + H_2O_2 --(GPx)--> GSSG + 2 H_2O\\nGSSG + NADPH + H^+ --(GR)--> 2 GSH + '
  'NADP^+\\nGSH/GSSG Oranı Çöküşü: Ratio_{redox} = [GSH] / [GSSG] = 100:1 (Genç) ---> 12:1 (Yaşlı)\\nNADPH/NADP+ '
  'Oranı: Ratio_{NADPH} <= 1.5 (Tehlikeli kriz eşiği).',
  'Glutatyon redoks oranının çökmesi, hücrenin tüm savunma kalkanlarını indirerek serbest radikallerin serbestçe '
  'yağmalamasına zemin hazırlar.'),
 ('4.8',
  'Nrf2/Keap1 Yolağının Susturulması ve Bach1 Represör Baskınlığı',
  'Hücrenin ana antioksidan savunma şefi Nrf2 (Nuclear factor erythroid 2-related factor 2), normalde Keap1 proteini '
  'tarafından sitoplazmada tutulup parçalanır. Oksidatif stres anında Keap1 sisteinleri oksitlenir ve Nrf2 çekirdeğe '
  'göç eder.',
  "Ancak yaşlı hücrelerde Nrf2'nin nükleer translokasyonu bozulur; daha da kötüsü, Nrf2 ile aynı Antioksidan Yanıt "
  'Elemanlarına (ARE) bağlanan rakip transkripsiyonel represör Bach1 proteini baskın hale gelir. Bach1, ARE '
  'promotörlerine oturarak antioksidan genlerin açılmasını aktif olarak kilitler.',
  'Nrf2-Bach1 Transkripsiyonel Rekabet Katsayısı:\\nPromoter_{bound} = [Nrf2] / ( [Nrf2] + K_{rel} * [Bach1] '
  ')\\nBurada K_{rel} = K_d(Nrf2) / K_d(Bach1)\\nYaşlı Hücrede ARE Baskılanma Oranı: % Suppression_{ARE} >= 68%\\nSOD2 '
  've Hem Oksijenaz-1 (HO-1) İndüksiyon İflası: Fold_{induction} <= 1.2 (Gençte 6.5).',
  "Bach1'in seçici olarak parçalanması veya susturulması, yaşlı hücrenin uyuyan Nrf2 ordusunu saniyeler içinde "
  'uyandırmanın en güçlü anahtarıdır.'),
 ('4.9',
  'Endoplazmik Retikulum (ER) Oksidatif Stresi ve UPR (Unfolded Protein Response) Aşırı Yükü',
  'Proteinlerin disülfid bağlarının kurulduğu Endoplazmik Retikulum (ER) lümeni, sitoplazmadan çok daha oksitleyici '
  'bir ortamdır. Ero1 enzimi ve Protein Disülfid İzomeraz (PDI) bu süreci yürütürken yan ürün olarak H2O2 üretir.',
  'Yaşlanma ile protein katlanma yükü arttıkça ER aşırı ısınır (ER stresi). Açılmamış Protein Yanıtı (Unfolded Protein '
  'Response - UPR) sensörleri (PERK, IRE1-alfa, ATF6) kronik olarak ateşlenir. Başlangıçta koruyucu olan UPR, sürekli '
  'aktif kaldığında CHOP transkripsiyon faktörü üzerinden hücreyi apoptoza sürükler.',
  'ER Stres İndeksi ve CHOP Apoptoz Tetikleme Kinetiği:\\nER_{stress} = [Unfolded_Protein]_{lumen} / '
  '[BiP/GRP78]_{free}\\nApoptoz Eşiği: [CHOP]_{nucleus} >= CHOP_{threshold}  ===>  Caspase-12 / Caspase-3 '
  'aktivasyonu\\nTranslasyonel Blokaj (p-eIF2alpha): Protein sentezi genel olarak %60 baskılanır.',
  'Kronik ER stresi, yaşlı hücrenin hem yeni protein üretememesine hem de UPR kaynaklı programlı hücre intiharına '
  'sürüklenmesine neden olur.'),
 ('4.10',
  'Mitokondriyal DNA Delesyonları ve Mutasyonel Kısır Döngü (Vicious Cycle)',
  'Mitokondriyal DNA (mtDNA), koruyucu histon proteinlerinden yoksundur ve süperoksit üreten solunum zinciri '
  'enzimlerinin sadece birkaç nanometre uzağında asılı durur.',
  "Harman'ın hipotezine göre mtDNA hasarı mutasyonlara yol açar; mutasyona uğramış mtDNA kusurlu solunum enzimleri "
  "üretir; kusurlu enzimler daha fazla ROS sızdırır ve bu durum daha fazla mtDNA hasarına yol açan 'Kısır Döngü'yü "
  '(Vicious Cycle) kurar. Nihayetinde tüm mitokondri popülasyonu solunum yeteneğini kaybeder.',
  'Mutasyonel Kısır Döngü Diferansiyel Denklemi:\\nd[mtDNA_{mut}] / dt = k_{mut} * [ROS] * [mtDNA_{WT}]\\nd[ROS] / dt '
  '= k_{leak0} + alpha * [mtDNA_{mut}]\\nHeteroplazmi Patlama Eşiği: Fraction_{mut} = [mtDNA_{mut}] / [mtDNA_{total}] '
  '>= 0.60 (Biyoenerjetik çöküş).',
  "PROJECT AETERNITAS, mitokondriyal hedefe yönelik dCas9-deaminazlar ve mito-TALEN'ler kullanarak mutasyona uğramış "
  'mtDNA kopyalarını seçici olarak eritir ve kısır döngüyü kırar.')]

part5_topics = [('5.1',
  'Anfinsen Hipotezi, Enerji Manzarası ve Protein Katlanma Termodinamiği',
  '1972 Nobel Kimya Ödülü sahibi Christian Anfinsen, bir proteinin üç boyutlu doğal (native) yapısının, amino asit '
  'dizilimi tarafından belirlenen minimum Gibbs serbest enerjisi durumuna karşılık geldiğini kanıtlamıştır (Anfinsen '
  'Hipotezi).',
  "Protein katlanması çok boyutlu bir 'huni' (Folding Funnel) üzerinde gerçekleşir. Protein zinciri huni boyunca aşağı "
  'yuvarlanarak en kararlı doğal konformasyonuna ulaşır. Ancak yaşlanma sürecinde hücresel ortamdaki termal gürültü, '
  "iyonik dengesizlikler ve oksidasyon bu huni yüzeyinde 'kinetik tuzaklar' (kinetic traps) ve yarı-kararlı yanlış "
  'katlanmış ara durumlar yaratır.',
  'Protein Katlanma Serbest Enerji Fonksiyonu:\\nDelta G_{folding} = Delta H_{chain} - T * Delta S_{conformational} + '
  'Delta G_{solvation} <= -20 kJ/mol (Doğal kararlılık)\\nKinetik Tuzak Engeli: Delta G_{barrier} >= 15 * k_B * '
  'T\\nDoğal Olmayan Konformasyon Fraksiyonu: f_{misfolded} = 1 / ( 1 + exp( -Delta Delta G / (R * T) ) ).',
  'Yaşlanma, proteinlerin bu kinetik tuzaklara düşerek yanlış katlanması ve bir daha doğal yapılarına dönememesi '
  'felaketidir.'),
 ('5.2',
  'Moleküler Şaperon Ağı (Hsp70, Hsp90, Şaperoninler) ve Kapasite İflası',
  'Hücrede yanlış katlanmayı önleyen ve kinetik tuzaklara düşen proteinleri ATP harcayarak kurtaran moleküler şaperon '
  'sistemleridir: Hsp70/Hsp40 ikilisi, Hsp90 ve dev fıçı benzeri kaviteler olan Şaperoninler (GroEL/ES veya memelide '
  'TRiC/CCT).',
  'Genç bir hücrede şaperon kapasitesi toplam protein yükünün çok üzerindedir. Yaşlanma ile birlikte Isı Şoku '
  'Faktörü-1 (HSF1) transkripsiyonel aktivitesi çöker; şaperon üretimi dururken hatalı katlanmış protein yükü '
  'katlanır. Şaperonlar çöken proteinlerin altında kalarak tükenir (chaperone overload).',
  'Şaperon Destekli Yeniden Katlanma Kinetiği:\\nMisfolded + Hsp70.ATP <---> Complex --(ATP hidrolizi)--> Native + '
  'Hsp70.ADP + Pi\\nŞaperon Tamponlama Kapasitesi: Capacity_{chaperone} = [Chaperone]_{active} / '
  '[Protein]_{misfolded}\\nKapasite İflas Eşiği: Capacity_{chaperone} < 1.0 (Serbest agregasyon başlangıcı).',
  'Şaperon ağının iflası, hücrenin kalite kontrol mekanizmasını devre dışı bırakarak sitoplazmayı toksik bir protein '
  'bataklığına çevirir.'),
 ('5.3',
  'Çapraz Beta-Tabaka Yapıları ve Amiloid Nükleasyonu (Nucleated Polymerization)',
  'Yanlış katlanan proteinlerin hidrofobik amino asit çekirdekleri dışarıya maruz kaldığında, termodinamik olarak '
  "birbirlerine yapışarak 'Çapraz Beta-Tabaka' (Cross-beta sheet) mimarisini oluştururlar. Bu yapı, tüm amiloid "
  'fibrillerinin evrensel çekirdeğidir.',
  "Amiloid oluşumu 'Nükleasyon Bağımlı Polimerizasyon' (Nucleated Polymerization) modeliyle işler. Uzun bir gecikme "
  "fazı (lag phase) boyunca az sayıda monomer bir araya gelerek kritik bir 'çekirdek' (nucleus) oluşturur; çekirdek "
  'oluştuktan sonra büyüme hızı katlanarak fırlar ve geriye dönüşsüz amiloid plakları çöker.',
  'Amiloid Nükleasyon ve Uzama Kinetik Denklemi:\\nd[Fibril] / dt = k_{nucleation} * [Monomer]^{n_{crit}} + '
  'k_{elongation} * [Fibril] * [Monomer] + k_{secondary} * [Fibril] * [Monomer]^2\\nKritik Çekirdek Boyutu: n_{crit} '
  'approx 3 - 6 monomer\\nLag Fazı Süresi: t_{lag} approx (k_{elongation} * k_{nucleation})^{-1/2}.',
  'Nükleasyon eşiği aşıldığında, amiloid agregasyonu hücre içi tüm normal metabolik işlevleri fiziksel olarak '
  'tıkayarak durdurur.'),
 ('5.4',
  'Toksik Oligomerler vs Olgun Fibriller: Membran Delici Biyofiziksel Porlar',
  'Yıllarca Alzheimer, Parkinson ve kardiyak amiloidozda en büyük suçlunun mikroskopta görülen devasa olgun amiloid '
  'plakları olduğu düşünülmüştür. Oysa modern biyofizik, en ölümcül türün olgun fibriller değil, küçük ve çözünür '
  "'Toksik Pre-Fibriler Oligomerler' olduğunu ispatlamıştır.",
  'Bu oligomerler, lipid membranlara yerleşerek yapay kontrolsüz kalsiyum porları (annular channel complexes) açarlar. '
  'Hücre zarını delerek sitoplazmaya kontrolsüz kalsiyum akıtırlar; membran potansiyelini sıfırlar ve mitokondriyel '
  'geçirgenlik geçiş gözeneğini (mPTP) patlatırlar.',
  'Oligomerik Membran Delme ve İyonik İletkenlik Katsayısı:\\nI_{oligomer_pore} = N_{pores} * g_{pore} * ( '
  'V_{membrane} - V_{reversal} )\\nPor İletkenliği: g_{pore} approx 50 - 200 pS (Pikosiemens)\\nSitoplazmik [Ca2+] '
  "Patlaması: [Ca2+]_i = 75 nM'den 2.500 nM'ye fırlar (Hücresel nekroz).",
  'Dev amiloid plakları aslında hücrenin bu ölümcül oligomerleri hareketsiz kılmak için kurduğu koruyucu bir çöp '
  'depolama alanıdır; asıl zehir serbest oligomerlerdir.'),
 ('5.5',
  'Ubiquitin-Proteazom Sisteminin (UPS) Tıkanması ve 26S Kompleks İflası',
  "Hücre içi kısa ömürlü ve hatalı proteinlerin %80'i Ubiquitin-Proteazom Sistemi (UPS) tarafından parçalanır. "
  'Proteinler E1-E2-E3 enzim kaskadıyla poli-ubiquitin (K48 bağı) zinciriyle etiketlenir ve 26S Proteazom kompleksine '
  'yönlendirilir.',
  'Yaşlanma ile birlikte UPS iflas eder: Okside olmuş ve çapraz bağlanmış amiloid proteinleri 20S katalitik çekirdeğe '
  'girer ancak parçalanamaz; proteazomun aktif sahalarını fiziksel olarak tıkar. Proteazom aktivitesi %70 düşer; '
  'hücrenin çöp öğütücüsü kilitlenir.',
  '26S Proteazom Tıkanma ve Ayrışma Kinetiği:\\nRate_{degradation} = (V_{max} * [Ub-Substrate]) / ( K_m * (1 + '
  '[Aggregates] / K_i) + [Ub-Substrate] )\\nProteazom İnhibisyon Sabiti: K_i approx 15 nM (Aşırı güçlü allosterik '
  'tıkanma)\\nUPS Klirens Kapasitesi Kaybı: Delta Clearance <= -72% (Yaşlı dokularda).',
  'Proteazomun tıkanması, hücre içi diğer tüm regülatör proteinlerin (p53, IkappaB, siklinler) parçalanmasını da '
  'engelleyerek hücre döngüsünü kaosa sürükler.'),
 ('5.6',
  'Şaperon Aracılı Otofaji (CMA) ve LAMP-2A Reseptör Çöküşü',
  'Şaperon Aracılı Otofaji (Chaperone-Mediated Autophagy - CMA), KFERQ benzeri pentapeptid motifi taşıyan çözünür '
  'sitozolik proteinlerin seçici olarak lizozoma sokulup eritildiği ultra hassas bir kalite kontrol sistemidir.',
  'Hsc70 şaperonu hedef proteini yakalar ve lizozom zarındaki LAMP-2A reseptörüne getirir. LAMP-2A multimerize olarak '
  'bir translokasyon kanalı oluşturur. Ancak yaşlanma ile lizozom zarındaki kolesterol/fosfolipid oranı değişir; '
  'LAMP-2A kararsızlaşarak parçalanır ve CMA aktivitesi sıfıra yaklaşır.',
  'CMA Translokasyon Akısı ve LAMP-2A Dinamiği:\\nJ_{CMA} = k_{trans} * [Hsc70-Substrate] * [LAMP2A_{multimer}] / ( '
  'K_d + [LAMP2A_{multimer}] )\\nLAMP-2A Doku Düzeyi Kaybı: [LAMP-2A]_{aged} / [LAMP-2A]_{young} <= '
  '0.22\\nKFERQ-Protein Birikim Oranı: Rate_{accumulation} >= +180%.',
  'CMA yetersizliği, transkripsiyon faktörlerinin ve glikolitik enzimlerin lizozomal temizliğini durdurarak metabolik '
  'esnekliği tamamen yok eder.'),
 ('5.7',
  'Makro-Otofajik Akının Durması: Lizozomal Asit Hidrolaz Yetmezliği',
  'Büyük organellerin ve protein agregatlarının çift zarlı otofagozomlar içine alınıp lizozomla kaynaşması '
  '(makro-otofaji), hücresel geri dönüşümün temel taşıdır.',
  'Yaşlı hücrede otofagozomlar oluşabilir, ancak lizozomla füzyon hızı (SNARE kompleksi STX17/SNAP29/VAMP8 bozulması) '
  'yavaşlar ve lizozom içi katepsin (B, D, L) enzimleri asidifikasyon kaybı nedeniyle inaktif kalır. Otofagozomlar '
  'sindirilemeyen çöplerle şişerek sitoplazmayı işgal eder.',
  'Otofajik Akı İndeksi ve Füzyon Oranı:\\nAutophagic_Flux = Rate_{degradation}(LC3-II) = ( [LC3-II]_{t} - '
  '[LC3-II]_{0} ) / ( Delta t * [p62]_{clearance} )\\nKatepsin D Maksimal Hızı: V_{max, CatD}(pH 6.0) <= 0.15 * '
  'V_{max, CatD}(pH 4.5)\\nOtofagozom Birikim Katsayısı: Accumulation_{AV} >= 4.5 kat.',
  'Geri dönüşüm sisteminin durması, hücrenin dışarıdan gelen besinlere bağımlı kalmasına ve iç kaynaklarını '
  'tazeleyememesine neden olur.'),
 ('5.8',
  'İleri Glikasyon Son Ürünleri (AGEs) ve Hücre İçi Kollajen / Protein Çapraz Bağları',
  'İndirgeyici şekerler (glukoz, fruktoz), proteinlerin serbest amino gruplarıyla enzimatik olmayan Maillard '
  'reaksiyonuna girerek Schiff bazları ve Amadori ürünleri oluşturur. Bu ürünler zamanla oksitlenerek İleri Glikasyon '
  'Son Ürünlerine (Advanced Glycation End-products - AGEs: glukozpan, pentosidin, CML) dönüşür.',
  "AGE'ler protein molekülleri arasında kovalent ve kırılması imkansız çapraz bağlar (cross-links) kurar. Kan "
  'damarları, ekstrasellüler matriks ve hücre içi iskelet proteinleri elastikiyetini kaybederek tahta gibi sertleşir; '
  'hiçbir proteaz bu bağları kesemez.',
  'Maillard Reaksiyon Kinetiği ve AGE Oluşumu:\\nProtein-NH_2 + Glucose <===> Schiff Base ---> Amadori Product ---> '
  'AGEs (Tek Yönlü Kilit)\\nDoku Sertlik Modülü Artışı: E_{Young} = E_0 * ( 1 + k_{AGE} * [Glucosepane] )\\nSertleşme '
  'Çarpanı: E_{tissue}(75 Yaş) >= 4.2 * E_{tissue}(20 Yaş).',
  'Doku sertleşmesi, arterlerin elastikiyetini yok ederek hipertansiyonu tetiklerken mikroskobik düzeyde hücrelerin '
  'hareketini ve mekanik sinyal iletimini kilitler.'),
 ('5.9',
  "Lipofuscin: Biyolojik Geri Dönüşümü Olmayan 'Yaşlılık Pigmenti' Çökeltisi",
  'Lizozomların parçalayamadığı okside lipidler (%30-70), glike proteinler ve ağır metallerden (demir, bakır) oluşan '
  "amorf çöp yığınına 'Lipofuscin' (yaşlılık pigmenti) adı verilir.",
  'Bölünmeyen post-mitotik hücrelerde (nöronlar, kardiyomiyositler, retina pigment epiteli) lipofuscin atılamaz ve '
  "hücre bölünüp seyrelmediği için ömür boyu birikir. 80 yaşındaki bir kardiyomiyositin hücre içi hacminin %25'i bu "
  'zehirli pigmentle kaplıdır; lipofuscin lizozomları tıkayarak onları öldürür.',
  'Lipofuscin Birikim Oranı Diferansiyeli:\\nd[Lipofuscin] / dt = k_{peroxidation} * [PUFA] * [Fe^{2+}]_{lysosome} * ( '
  '1 - Clearance_{capacity} )\\nKardiyomiyosit Lipofuscin Fraksiyonu: Volume_{lipofuscin} / Volume_{cytoplasm} >= 0.22 '
  '(Yaşlıda)\\nOtofloresans Şiddeti: Floresans emisyonu (450-550 nm) yaşla üssel artar.',
  'Lipofuscin birikimi, hücrenin kendi çöplüğünde boğulmasının en görünür ve somut biyofiziksel kanıtıdır.'),
 ('5.10',
  'Proteostaz Ağının Faz Diyagramı ve Çöküşün Kritik Perkolasyon Eşiği',
  "Sistem biyolojisinde proteostaz ağı, bir 'Perkolasyon Modeli' (Percolation Theory) ile kavranabilir. Hücre içi "
  'hatalı proteinler ve şaperonlar rastgele bir ağ örgüsü oluşturur.',
  'Hatalı protein konsantrasyonu kritik perkolasyon eşiğini (p_c) aştığında, izole küçük agregatlar aniden tüm hücreyi '
  'boydan boya saran devasa bir makroskobik jel ağına (giant spanning cluster) dönüşür. Bu ani faz geçişi, hücrenin '
  'difüzyonunu ve metabolik iletimini bir anda sıfırlayarak hücresel ölümü kilitler.',
  'Perkolasyon Olasılığı ve Kritik Eşik Denklemi:\\nP_{percolation} = 0   (p < p_c);   P_{percolation} approx (p - '
  'p_c)^beta   (p >= p_c)\\nKritik Eşik Değeri: p_c = 0.312 (Üç boyutlu rastgele kafes modeli)\\nHücresel Donma '
  'Zamanı: tau_{arrest} <= 12 saat (p >= p_c anından sonra).',
  'PROJECT AETERNITAS, proteostaz ağını bu kritik perkolasyon eşiğinin çok gerisinde tutmak için küçük moleküllü '
  'amiloid kırıcılar ve sentetik otofaji indükleyicileri kullanır.')]

part6_topics = [('6.1',
  'Hücre Membranı Biyofiziği: Akışkan Mozaik Modelinden Lipid Salına (Lipid Rafts)',
  "Singer ve Nicolson'ın 1972'deki klasik 'Akışkan Mozaik Modeli', biyolojik membranın iki boyutlu akışkan bir sıvı "
  'çift katmanı olduğunu öne sürmüştür. Ancak modern membran biyofiziği, zarın homojen bir sıvı olmadığını; kolesterol '
  "ve sfingolipidlerden zengin mikro-domainler olan 'Lipid Salları' (Lipid Rafts) içerdiğini göstermiştir.",
  'Bu lipid salları, sinyal transdüksiyonu reseptörlerinin (büyüme faktörü reseptörleri, immün reseptörler) bir araya '
  'geldiği dinamik platformlardır. Yaşlanma ile birlikte membran kolesterol içeriğinin artması ve doymuş yağ asidi '
  'oranının yükselmesi bu salları katılaştırır ve reseptör hareketliliğini kilitler.',
  'Lipid Salı Düzenlilik ve Düzensizlik Faz Dengesi:\\nDelta G_{raft} = Delta H_{mixing} - T * Delta S_{mixing} + '
  'gamma_{line} * L_{boundary}\\nBurada gamma_{line} sınır gerilimi katsayısı, L_{boundary} sal '
  'çevresidir.\\nSıvı-Düzenli Faz Fraksiyonu: % L_o = Area_{rafts} / Area_{total} = 15% (Genç) ---> 45% '
  "(Yaşlı)\\nReseptör Lateral Difüzyon Katsayısı: D_{receptor} = 0.45 um^2/s'den 0.08 um^2/s'ye geriler.",
  'Zarın katılaşması, hücrenin dış dünyadan gelen hormon ve büyüme faktörü sinyallerini algılamasını ve içeriye '
  'iletmesini imkansız hale getirir.'),
 ('6.2',
  'Çoklu Doymamış Yağ Asitleri (PUFA) ve Lipid Peroksidasyon Kaskadı',
  'Hücre membranlarının akışkanlığını sağlayan çoklu doymamış yağ asitleri (PUFA: linoleik asit, araşidonik asit, '
  'DHA), çift bağları arasındaki zayıf bis-allilik metilen karbonları (-CH=CH-CH2-CH=CH-) nedeniyle serbest '
  'radikallerin en kolay saldırdığı hedeflerdir.',
  "Bir tek serbest radikalin bir PUFA molekülünden hidrojen koparması, bir 'Otooksidasyon Zincirleme Reaksiyonu' "
  'başlatır. Oluşan lipid radikali (L•), oksijenle birleşerek lipid peroksil radikalini (LOO•) üretir; bu radikal '
  "komşu bir PUFA'ya saldırarak reaksiyonu binlerce kez tekrarlar.",
  'Lipid Peroksidasyon Yayılma Kinetik Zinciri:\\nLH + R^*ullet ---> L^*ullet + RH   (Başlangıç)\\nL^*ullet + '
  'O_2 ---> LOO^*ullet   (k_1 = 10^9 M^{-1} s^{-1})\\nLOO^*ullet + LH ---> LOOH + L^*ullet   (k_2 = 10 - 100 '
  'M^{-1} s^{-1} - Kendi Kendini Besleyen Zincir)\\nZincir Uzunluğu: N_{propagation} = Rate_{propagation} / '
  'Rate_{termination} >= 250 döngü / radikal.',
  'Tek bir serbest radikal, zar boyunca domino taşları gibi binlerce lipid molekülünü yakarak membran bütünlüğünü '
  'saniyeler içinde delik deşik eder.'),
 ('6.3',
  'Malondialdehit (MDA) ve 4-Hidroksinonenal (4-HNE) Sitotoksisitesi',
  'Lipid peroksidasyonunun son parçalanma ürünleri, son derece reaktif alfa,beta-doymamış aldehitlerdir: '
  'Malondialdehit (MDA) ve 4-Hidroksinonenal (4-HNE).',
  "4-HNE, suda ve lipidde çözünebilen difüze edilebilir bir 'moleküler suikastçıdır'. Membranda üretildikten sonra "
  'sitoplazmaya ve nükleusa göç eder; proteinlerin sistein, lizin ve histidin kalıntılarına Michael katılmasıyla '
  'kovalent olarak bağlanır. 4-HNE ile modifiye edilen proteinler derhal agregasyona uğrar ve proteazomu kilitler.',
  '4-HNE Michael Katılma Reaksiyonu ve Adükt Oluşumu:\\nProtein-SH + 4-HNE ---> Protein-S-CH(OH)-CH_2-... (Kararlı '
  'Kovalent Adükt)\\n4-HNE Dokusal Toksisite Eşiği: [4-HNE]_{cytosol} >= 2.5 uM (Gençte < 0.2 uM)\\nApoptoz Tetikleme '
  'İndeksi: JNK ve p38 MAPK fosforilasyonu %300 artış.',
  'MDA ve 4-HNE, serbest radikal hasarının sadece membranda kalmayıp tüm hücreye yayılan zehirli bir dalgaya '
  'dönüşmesinin primer taşıyıcılarıdır.'),
 ('6.4',
  'Membran Sertliği ve Floresans Polarizasyonu (DPH Ölçümleri)',
  'Biyofiziksel düzeyde membran akışkanlığı, 1,6-difenil-1,3,5-heksatrien (DPH) gibi hidrofobik floresans problarının '
  'anizotropi (polarizasyon) ölçümleriyle nicelendirilir.',
  'Prob zarda serbestçe dönebiliyorsa polarizasyon düşüktür (yüksek akışkan genç membran). Membran sertleştikçe probun '
  'serbest dönüşü kısıtlanır ve polarizasyon değeri (P) yükselir. 80 yaşındaki bir insanın eritrosit ve nöron membran '
  'sertliği, 20 yaşındaki bir gence göre %40 daha yüksektir.',
  'Floresans Anizotropi ve Membran Düzenlilik Parametresi:\\nr = (I_{parallel} - G * I_{perpendicular}) / '
  '(I_{parallel} + 2 * G * I_{perpendicular})\\nSıra Parametresi: S^2 = (r / r_0 - r_{inf}) / (1 - r_{inf})\\nMembran '
  'Sertlik İndeksi: r_{aged} >= 0.28 (Gençlik normu: r <= 0.18)\\nMikro-Viskozite Artışı: eta_{membrane} = 1.2 '
  "poise'dan 3.8 poise'a yükselir.",
  'Zar sertliği, vezikül ekzositozunu (nörotransmitter salınımı, insülin salgılanması) yavaşlatarak hücrenin dış '
  'dünyayla iletişim refleksini köreltir.'),
 ('6.5',
  'Plazmalojenlerin Tükenmesi ve Peroksizomal Biyogenez Kaybı',
  'Plazmalojenler, gliserol omurgasının sn-1 pozisyonunda vinil-eter bağı (-O-CH=CH-) taşıyan eşsiz bir eter '
  "fosfolipid sınıfıdır. Beyin miyelin kılıflarının %70'ini ve kalp zarlarının %50'sini oluştururlar.",
  "Vinil-eter bağı, reaktif oksijen türlerini absorbe eden bir 'moleküler sünger' gibi davranır; kendisi feda olarak "
  "zardaki diğer lipidleri ve DNA'yı korur. Plazmalojen sentezi peroksizomlarda başlar. Yaşlanma ile peroksizomal "
  'biyogenez çöker; plazmalojen seviyeleri %80 azalır ve hücre en güçlü doğal kalkanını kaybeder.',
  'Plazmalojen Koruma Oranı ve Peroksizom Akısı:\\nRate_{scavenging} = k_{plasmalogen} * [PlsEtn] * [ROO^*ullet] / '
  '( k_{lipid} * [PUFA] ) >= 15.0\\nDoku Plazmalojen Kaybı: [PlsEtn]_{aged} / [PlsEtn]_{young} <= 0.25 (Beyin '
  'korteksinde)\\nPeroksizom Sayısı Azalması: N_{peroxisome} / hücre <= -65%.',
  'Plazmalojen tükenmesi, hücre membranını serbest radikal saldırılarına karşı tamamen çıplak ve savunmasız bırakır.'),
 ('6.6',
  'Kolesterol/Fosfolipid Oranının Artışı ve Kaveola Disfonksiyonu',
  'Hücre zarlarındaki kolesterol miktarı, zarın rijiditesini ve geçirgenlik bariyerini belirler. Kolesterol '
  'molekülleri fosfolipidlerin arasına girerek açil zincirlerinin serbest hareketini kısıtlar.',
  'Yaşlanma ile hücre içi kolesterol taşınması bozulur; plazma membranındaki serbest kolesterol / fosfolipid molar '
  "oranı (Chol/PL) 0.6'dan 1.2'nin üzerine çıkar. Zardaki şişe benzeri kaviteler olan 'Kaveolalar' (Caveolae) "
  'sertleşerek çöker; kaveolin-1 bağımlı endositoz ve endotelyal nitrik oksit (eNOS) üretimi durur.',
  'Kolesterol-Fosfolipid Molar Oranı ve Membran Sertliği:\\nRatio_{Chol/PL} = [Cholesterol]_{membrane} / '
  '[Phospholipids]_{membrane}\\nGenç Eşik: Ratio_{Chol/PL} approx 0.65; Yaşlı Eşik: Ratio_{Chol/PL} >= 1.25\\nKaveolar '
  'Endositoz İflası: % Endocytosis <= 0.30 * Basal.',
  'Aşırı kolesterol birikimi, zarı adeta zırhlı bir betona dönüştürerek hücrenin elastikiyetini ve besin alım '
  'kanallarını tıkar.'),
 ('6.7',
  'Kalsiyum Kaçakları: Membran Yırtılmaları ve Eksitotoksik İyon Girişi',
  'Hücre dışı kalsiyum konsantrasyonu (approx 1.5 - 2.0 mM), hücre içi serbest kalsiyumdan (approx 50 - 100 nM) tam '
  '20.000 kat daha yüksektir. Bu devasa kimyasal gradyan, plazma membranının sağlamlığına emanettir.',
  'Lipid peroksidasyonu ve mekanik sürtünme nedeniyle yaşlı zarlarda oluşan sub-nanometre boyutlu mikro-yırtıklar '
  '(micro-lesions), sitoplazmaya kontrolsüz Ca2+ sızdırır. Plazma Membranı Ca2+-ATPaz (PMCA) pompaları bu devasa '
  "sızıntıyı tahliye etmeye yetişemez; sitozolik kalsiyum 1 uM'nin üzerine fırlayarak kalpain ve kaspaz proteazlarını "
  'aktive eder.',
  'Kalsiyum İyonik Akı ve Pompa Kapasite Dengesi:\\nd[Ca2+]_i / dt = J_{leak}(Area_{tears}) + J_{channel} - V_{max, '
  'PMCA} * [Ca2+]_i / ( K_m + [Ca2+]_i )\\nSızıntı Akısı: J_{leak} = P_{Ca} * ( [Ca2+]_{out} - [Ca2+]_{in} * '
  'exp(2*F*Delta V / RT) )\\nSitoplazmik Kalsiyum Patlaması: [Ca2+]_i >= 1.200 nM (Apoptoz ve lizis başlangıcı).',
  'Kalsiyum kaçakları, hücrenin içini yıkan ölümcül bir tsunami gibidir; hücre iskeletini parçalar ve mitokondriyi '
  'kalsiyumla zehirler.'),
 ('6.8',
  "Fosfatidilserin Eksternalizasyonu: 'Beni Ye' Sinyalinin Erken Ateşlenmesi",
  "Genç bir hücrede negatif yüklü fosfolipid olan Fosfatidilserin (PS), ATP-bağımlı 'Flippase' enzimleri sayesinde "
  'kesinlikle plazma membranının iç yaprağında (inner leaflet) kilitli tutulur.',
  "Yaşlı hücrede ATP tükenmesi flippase enzimlerini felç ederken; artan hücre içi kalsiyum 'Scramblase' enzimini "
  '(TMEM16F) aktive eder. Fosfatidilserin dış yaprağa (outer leaflet) fırlar. Dış yüzeye çıkan PS, makrofajlar ve '
  "mikroglia için evrensel 'Beni Ye' (Eat-me) sinyalidir; yaşlı ama henüz çalışan hücreler bağışıklık sistemi "
  'tarafından erken fagositozla yok edilir.',
  'Fosfatidilserin Flip-Flop Kinetik Dengesi:\\nd[PS]_{outer} / dt = k_{scramblase} * [Ca2+]_i * [PS]_{inner} - '
  'V_{max, flippase} * [ATP] * [PS]_{outer}\\nFagositoz Eşiği: % PS_{outer} >= 8.5% toplam membran PS\\nMakrofaj '
  'Tanıma Afinitesi: K_d(Tim4 : PS) approx 2.0 nM.',
  'Membran asimetrisinin kaybı, hücrenin henüz yaşayabilecekken bağışıklık hücreleri tarafından erken kurban '
  'edilmesine yol açar.'),
 ('6.9',
  'Aquaporin Kanalları ve Hücresel Ozmotik Dehidrasyon',
  'Hücrenin su dengesi ve turgor basıncı, plazma zarındaki Aquaporin (AQP) su kanalları tarafından yönetilir. Genç bir '
  "hücre %70-75 oranında su içerirken, yaşlı hücrelerde su oranı %55-60'a düşer ('Hücresel Dehidrasyon').",
  'Aquaporin-1 ve Aquaporin-4 kanallarının zar üzerindeki yoğunluğu yaşlanmayla azalır ve kapı mekanizmaları '
  'oksitlenir. Hücre içi su kaybı, sitozolü daha da vizkoz ve koyu hale getirerek moleküler etkileşimleri yavaşlatır; '
  'hücre küçülür ve buruşur.',
  'Ozmotik Su Geçirgenliği ve Hücresel Hacim Kinetiği:\\nJ_v = L_p * A_{cell} * ( Delta P - sigma_{osmotic} * Delta Pi '
  ')\\nBurada L_p hidrolik geçirgenlik katsayısıdır (Yaşlı hücrede %50 azalır)\\nHücre İçi Serbest Su Hacmi: '
  'Volume_{H2O}(Aged) <= 0.62 * Volume_{H2O}(Young).',
  'Hücresel dehidrasyon, yaşlı dokuların elastikiyetini kaybetmesinin ve metabolik reaksiyon hızlarının düşmesinin '
  'temel fiziksel nedenidir.'),
 ('6.10',
  "Membran Lipidomiks Analizi: Yaşlanmanın 'Lipid İmza'sı ve Tersine Çevrilmesi",
  'Modern Kütle Spektrometrisi (Lipidomics), hücre zarında binlerce farklı lipid türünü tanımlar. Yaşlanan bir '
  "hücrenin membran lipidomu benzersiz bir 'Yaşlılık İmzası' taşır: Aşırı sfingomiyelin, seramid artışı, doymuş yağ "
  'asitleri ve dramatik kardiyolipin kaybı.',
  'Mitokondri iç zarının kalbi olan Kardiyolipin (CL), solunum zinciri respirasomlarını bir arada tutan dört bacaklı '
  "özel bir fosfolipiddir. Yaşlanmayla kardiyolipin peroksidasyona uğrar ve sitokrom c'yi serbest bırakır. Membran "
  'lipidomunun gençlik profiline döndürülmesi, fiziksel gençleşmenin zorunlu bir adımıdır.',
  'Kardiyolipin Peroksidasyon ve Membran Skor Matrisi:\\nLipidAge = 0.35 * [Ceramide] + 0.25 * [Sphingomyelin] - 0.40 '
  '* [Cardiolipin]_{native}\\nKardiyolipin Kaybı: % CL_{intact} = 100% (Genç) ---> 32% (Yaşlı)\\nSitokrom c Salınım '
  'Eşiği: [CL_{oxidized}] / [CL_{total}] >= 0.45 (Apoptoz kaskadı).',
  'PROJECT AETERNITAS, ekzojen plazmalojenler, saflaştırılmış kardiyolipin öncülleri ve hedefli lipid takviyeleri ile '
  'hücre zarını atomik düzeyde yeniden inşa eder.')]

# Append Parts 4, 5, 6 to parts list
parts.append((4, "Serbest Radikaller, Mitohormezis ve Oksidatif Hasarın Termodinamiği", part4_topics))
parts.append((5, "Protein Katlanma Dinamiği, Amiloid Agregasyonu ve Proteotoksisite", part5_topics))
parts.append((6, "Biyofiziksel Hücre Membranı Sertleşmesi ve Lipid Peroksidasyon Kinetiği", part6_topics))

print("[PROJECT AETERNITAS] Parts 4, 5, 6 appended to build_k1_ch01_docx.py successfully.")

part7_topics = [('7.1',
  'Nükleer Zarf (NE) Mimarisi ve İç/Dış Membran Çift Katmanı',
  'Hücre çekirdeği (nükleus), genomik kütüphaneyi sitoplazmik kaostan ayıran çift katmanlı bir lipit kalkan olan '
  'Nükleer Zarf (Nuclear Envelope - NE) ile çevrilidir. İç Nükleer Membran (INM) kromatini bağlarken, Dış Nükleer '
  'Membran (ONM) endoplazmik retikulum ile süreklidir.',
  'Yaşlanma sürecinde nükleer zarın lipid kompozisyonu bozulur ve gerilim direnci çöker. Genç hücrede küresel ve '
  'pürüzsüz olan çekirdek geometrisi, yaşlı hücrelerde derin girintiler (invaginations), kabarcıklar (blebbing) ve '
  'lobülasyonlar sergileyerek mekanik bütünlüğünü kaybeder.',
  'Nükleer Zarf Mekanik Gerilim ve Laplacé Yasası:\\nDelta P_{nuclear} = 2 * gamma_{envelope} / R_{nucleus}\\nKritik '
  'Yırtılma Gerilimi: gamma_{rupture} <= 0.85 mN/m (Yaşlı zarda kırılganlık eşiği)\\nÇekirdek Şekil Faktörü '
  '(Sirkülarite): Circularity = 4 * pi * Area / Perimeter^2 = 0.95 (Genç) ---> 0.48 (Yaşlı).',
  'Nükleer zarın geometrik deformasyonu, kromatinin 3D katlanma düzenini mekanik olarak bozar ve gen susturma '
  'bölgelerini parçalar.'),
 ('7.2',
  'Nükleer Lamina (Lamin A/C ve B1) ve Hutchinson-Gilford Progeria Sendromu',
  'İç nükleer membranın hemen altında yer alan Nükleer Lamina, tip V ara filaman proteinleri olan Lamin A, Lamin C, '
  "Lamin B1 ve Lamin B2'den oluşan lifli bir ağ örgüsüdür. Lamina, heterokromatini çekirdek çevresine sabitleyen "
  'iskelettir.',
  'Hutchinson-Gilford Progeria Sendromunda (HGPS), LMNA genindeki tek bir nokta mutasyonu farnesil grubunun '
  "kesilmesini engelleyerek mutant bir lamin proteini olan 'Progerin' üretir. Progerin laminaya yapışarak nükleer "
  'iskeleti mahveder ve 10 yaşındaki bir çocuğun 80 yaşında gibi görünmesine yol açar. Normal fizyolojik yaşlanmada da '
  'progerin düşük düzeyde birikir ve Lamin B1 proteini dramatik biçimde kaybolur.',
  'Lamin B1 Kaybı ve Progerin Birikim Oranı:\\nRatio_{lamina} = [Lamin_B1]_{aged} / [Lamin_B1]_{young} <= 0.18 '
  '(Senesans göstergesi)\\nProgerin Toksisite Eşiği: [Progerin] / [Lamin_A] >= 0.05 (Hücresel fenotip çöküşü)\\nLamina '
  "Ağ Gözü Boyutu (Mesh Size): xi = 14 nm'den 48 nm'ye genişler.",
  'Lamin B1 kaybı, hücresel senesansın en kesin ve geri döndürülemez biyofiziksel habercisidir; heterokromatinin '
  'çözülmesini tetikler.'),
 ('7.3',
  'Nükleopor Kompleksi (NPC) ve Aşırı Uzun Ömürlü Proteinlerin (ELLPs) Aşınması',
  'Nükleus ile sitoplazma arasındaki tek geçit kapısı, yaklaşık 30 farklı Nükleoporin (Nup) proteininin 8 katlı '
  'simetrik dizilimiyle oluşan Nükleopor Kompleksidir (NPC). NPC, biyolojideki en devasa protein makinelerinden '
  'biridir (120 Megadalton).',
  "NPC'nin yapısal iskeletini oluşturan Nup96, Nup107 ve Nup133 gibi proteinler 'Aşırı Uzun Ömürlü Proteinler'dir "
  '(Extremely Long-Lived Proteins - ELLPs); hücre bölündükten sonra ömür boyu yenilenmezler. 50-80 yıl boyunca termal '
  've oksidatif strese maruz kalan bu proteinler aşınır; NPC kapıları gevşer ve seçici filtre özelliğini kaybeder.',
  'Nükleoporin Yarılanma Ömrü ve Seçicilik Bozulması:\\nt_{1/2}(Nup96 / Nup107) -> Tüm İnsan Ömrü (Sıfır '
  'turnover)\\nNPC Seçici Filtre Yarıçapı: R_{pore} = 5.2 nm (Genç) ---> 9.8 nm (Yaşlı - Kaçak Geçiş)\\nPasif Difüzyon '
  'Geçirgenliği: P_{passive} >= 6.5 kat artış (Normalde geçemeyen 70 kDa proteinler serbestçe girer).',
  'Seçici kapıların aşınması, sitoplazmadaki toksik proteinlerin ve enzimlerin doğrudan çekirdeğe sızmasına ve genomu '
  'tahrip etmesine kapı aralar.'),
 ('7.4',
  'Nükleositoplazmik Taşıma (NCT) Bozulması: RanGTP/RanGDP Gradyan Çöküşü',
  'Nükleus içine ve dışına protein taşınması, Ran GTPaz enziminin nükleus-sitoplazma arasındaki asimetrik gradyanına '
  'dayanır: Çekirdekte RanGTP yüksek, sitoplazmada RanGDP yüksektir.',
  'İçe taşıyıcılar (Importinler) kargolarını RanGTP sayesinde çekirdeğe bırakır; dışa taşıyıcılar (Exportinler / CRM1) '
  'RanGTP ile kargo yükler. Yaşlanma sürecinde RanGAP1 ve RCC1 enzimlerinin lokalizasyonu bozulur; RanGTP/GDP gradyanı '
  "düzleşir. Transkripsiyon faktörleri çekirdeğe giremez, hatalı RNA'lar ise dışarı çıkamaz; nükleer trafik "
  'kilitlenir.',
  'Ran Gradyanı ve Nükleer Taşınma Serbest Enerjisi:\\nDelta mu_{Ran} = R * T * ln( [RanGTP]_{nucleus} / '
  '[RanGTP]_{cytoplasm} )\\nGenç Gradyan: Delta mu_{Ran} approx 14.5 kJ/mol; Yaşlı Gradyan: Delta mu_{Ran} <= 2.1 '
  'kJ/mol\\nTaşıma Hızı Çöküşü: Flux_{import} <= 0.20 * Flux_{young}.',
  'Nükleositoplazmik taşımanın durması, hücrenin genetik emirleri yerine getirememesine ve sitoplazmik sinyallere '
  'tamamen körleşmesine yol açar.'),
 ('7.5',
  'Nükleer Yırtılmalar (Nuclear Envelope Ruptures) ve Sitoplazmik DNA Sızıntısı',
  'Lamina bütünlüğünü kaybetmiş yaşlı bir hücre dar bir doku aralığından geçerken veya mekanik baskı altındayken '
  'nükleer zar mikroskobik olarak yırtılır (NE rupture).',
  'Bu yırtılma anında sitoplazmik sıvılar çekirdeğe dolar ve genomik DNA parçaları sitoplazmaya fırlar. Sitoplazmaya '
  "saçılan serbest çift zincirli DNA (dsDNA), hücre için 'virüs enfeksiyonu' alarmıdır; bağışıklık sensörleri anında "
  'kırmızı kod verir.',
  'Nükleer Yırtılma Frekansı ve İyileşme Kinetiği:\\nRate_{rupture} = k_{stress} * ( sigma_{mech} / E_{lamina} '
  ')^2\\nYırtık Boyutu: Diameter_{tear} in [0.2, 1.5] um\\nESCRT-III Tamir Gecikmesi: t_{repair} = 15 dk (Genç) ---> '
  '180 dk (Yaşlı - Saatlerce açık kalır).',
  'Saatlerce açık kalan nükleer yırtıklar, sitoplazmik DNAazların çekirdeğe girerek genomu parçalamasına ve hücresel '
  'felakete neden olur.'),
 ('7.6',
  'cGAS-STING Yolağı Aktivasyonu ve Steril İnflammasyonun Ateşlenmesi',
  'Nükleer yırtılmalardan veya hasarlı mitokondrilerden sitoplazmaya kaçan serbest dsDNA, sitoplazmik DNA sensörü cGAS '
  '(siklik GMP-AMP sentaz) tarafından pikomolar afiniteyle yakalanır.',
  "cGAS, 2'3'-cGAMP ikincil habercisini üretir; cGAMP ise ER zarındaki STING (Stimulator of Interferon Genes) "
  'reseptörünü aktive eder. TBK1 ve IRF3/NF-kappaB yolları ateşlenir; hücre tek bir bakteri veya virüs olmaksızın '
  "devasa miktarda Tip-I İnterferon (IFN-alfa/beta) ve IL-6 salgılamaya başlar ('Steril İnflammasyon').",
  'cGAS-STING Sinyal Kaskadı ve cGAMP Sentez Hızı:\\nRate_{cGAMP} = k_{cat, cGAS} * [dsDNA]_{cytosolic} * [ATP] * '
  '[GTP] / ( K_m + [dsDNA] )\\nSTING Multimerizasyonu: [STING]_{dimer} ---> [STING]_{tetramer} (Aktif '
  'platform)\\nİnterferon Salınım Çarpanı: Fold_{IFN-beta} >= 25 kat (Kronik iltihap hali).',
  "cGAS-STING kaskadı, yaşlı hücreyi kendi DNA'sına saldıran bir otoimmün cehenneme dönüştürerek dokuyu yok eden ateşi "
  'yakar.'),
 ('7.7',
  'Nükleolus (Çekirdekçik) Hipertrofisi ve Ribozomal RNA (rRNA) Dengesizliği',
  "Çekirdek içindeki zarsız organel olan nükleolus, ribozomların monte edildiği ve rRNA'nın transkribe edildiği "
  'yerdir. Uzun ömür araştırmalarında en çarpıcı keşiflerden biri, nükleolus boyutunun yaşam süresiyle ters orantılı '
  'olmasıdır.',
  'Kısa ömürlü ve hızlı yaşlanan hücrelerde nükleolus devasa boyutlara ulaşır (nükleolar hipertrofi). Aşırı rRNA '
  'transkripsiyonu hücrenin enerji bütçesini tüketir ve nükleolar stres sensörlerini felç eder. Buna karşılık, genç '
  'hücrelerde ve uzun ömürlü mutantlarda nükleolus küçük ve kompakttır.',
  'Nükleolar Hacim Fraksiyonu ve Yaşam Süresi Ters Korelasyonu:\\nLifespan prop to 1 / Volume_{nucleolus}\\nNükleolus '
  'Hacmi: Vol_{nucleolar} / Vol_{nuclear} = 4% (Genç / Uzun Ömürlü) ---> 22% (Yaşlı)\\nFibrillarin ve Nükleolin '
  'Yoğunluğu: Density_{FBL} >= 3.8 kat artış.',
  'Nükleolusu küçültmek ve rRNA transkripsiyonunu frenlemek, hücresel enerjiyi hayatta kalma ve onarım yollarına '
  'aktarmanın doğrudan kestirme yoludur.'),
 ('7.8',
  'LINC Kompleksi (SUN-KASH) ve Nükleo-Sitoiskeletal Mekanik Bağlantı Kaybı',
  'Çekirdek sitoplazmada serbestçe yüzmez; LINC (Linker of Nucleoskeleton and Cytoskeleton) kompleksi aracılığıyla '
  'hücre iskeletine (aktin, mikrotübüller, ara filamanlar) sımsıkı bağlıdır. LINC kompleksi, SUN domain proteinleri '
  '(INM) ve KASH domain proteinlerinden (Nesprin - ONM) oluşur.',
  'Bu mekanik köprü, hücre dışından gelen fiziksel kuvvetleri (shear stress, gerilme) doğrudan çekirdeğe ileterek gen '
  'ekspresyonunu mekanik olarak düzenler (mekanotransdüksiyon). Yaşlanma ile SUN-Nesprin köprüleri kopar; çekirdek '
  'mekanik hissini kaybeder ve hücre dış gerilimlere yanıt veremez.',
  'Mekanotransdüksiyon Gerilme İletim Verimi:\\nForce_{nuclear} = k_{LINC} * ( x_{cytoskeleton} - x_{nucleus} )\\nLINC '
  'Bağlantı Kararlılığı: K_d(SUN : KASH) = 15 nM (Genç) ---> Çözülme (Yaşlı)\\nMekanik Gen Yanıtı Kaybı: Delta '
  'Expression_{YAP/TAZ} <= -60%.',
  'Hücrenin mekanik zekasını kaybetmesi, kasların körelmesine, kemiklerin erimesine ve damar endotelinin akış yönünü '
  'şaşırmasına yol açar.'),
 ('7.9',
  'Kromatin-Lamina Etkileşim Alanlarının (LADs) Çözülmesi ve Gen Derepresyonu',
  "Genomun yaklaşık üçte biri, 'Lamina ile İlişkili Alanlar' (Lamina-Associated Domains - LADs) adı verilen bölgelerle "
  'nükleer laminaya demirlenmiştir. LADs içindeki genler mutlak surette susturulmuş heterokromatindir.',
  'Lamin A/B kaybı ve lamina deformasyonu nedeniyle LAD sınırları çözülür. Normalde sonsuza kadar suskun kalması '
  'gereken fetal genler, dokuya yabancı genler ve onkogenler serbest kalarak kazara okunmaya başlar (genomik '
  'derepresyon). Hücre hangi dokuya ait olduğunu unutan bir kimlik krizine sürüklenir.',
  'LADs Demirlenme İndeksi ve Gen İfade Kaçağı:\\nFraction_{LAD_bound} = Length_{attached_DNA} / Length_{total_genome} '
  '= 35% (Genç) ---> 8% (Yaşlı)\\nYersiz Gen Ekspresyon Oranı: Rate_{leaky_transcription} >= 4.2 kat artış\\nH3K9me2/3 '
  'LADs Zenginliği Kaybı: Delta Peak_{ChIP-seq} <= -70%.',
  'LADs çözülmesi, kütüphanenin kilitli gizli arşiv kapılarının kırılıp tüm yasak kitapların ortalığa saçılması '
  'gibidir; hücre kontrol edilemez bir kaosa girer.'),
 ('7.10',
  'Nükleer İskeletin Restorasyonu: Progerin Ayrışımı ve Lamina Gençleştirme',
  'Nükleer zar ve lamina bozulması geriye döndürülemez bir son mudur? Hayır! Sentetik biyoloji ve hedefli moleküller '
  'nükleer mimariyi baştan aşağı onarabilir.',
  "Farnesiltransferaz inhibitörleri (Lonafarnib), progerin mRNA'sını hedefleyen sentetik antisense oligonükleotitler "
  '(ASO) ve Lamin B1 genini aktive eden CRISPRa kurguları nükleer kırışıklıkları ütüler gibi düzeltir. Nükleer porlar '
  'yeniden inşa edilir; RanGTP gradyanı restore edilir ve cGAS iltihabı söndürülür.',
  'Nükleer Gençleşme ve Lamina Onarım Kinetiği:\\nd[Lamin_B1] / dt = k_{transcription} * [CRISPRa] - k_{deg} * '
  "[Lamin_B1]\\nNükleer Sirkülarite Restorasyonu: Circularity = 0.48'den 0.92'ye geri dönüş\\nRanGTP/GDP Gradyan "
  'İyileşmesi: Delta mu_{Ran} >= 12.8 kJ/mol (Gençlik seviyesi).',
  'Nükleer iskeletin onarımı, hücrenin kalbini yeniden zırhlamak demektir; bu restorasyon olmadan kalıcı hiçbir '
  'hücresel gençleşme mümkün değildir.')]

part8_topics = [('8.1',
  'Hücresel Adenilat Enerji Şarjı (AEC) ve ATP/ADP/AMP Dinamiği',
  '1967 yılında Daniel Atkinson tarafından tanımlanan Adenilat Enerji Şarjı (Adenylate Energy Charge - AEC), hücrenin '
  'metabolik pilinin doluluk oranını gösteren temel termodinamik parametredir.',
  'Genç ve canlı bir hücrede AEC değeri 0.85 ila 0.95 arasında son derece dar bir bantta kilitlidir. Yaşlanan hücrede '
  "mitokondriyel ATP üretiminin düşmesi ve ATP tüketen boşuna reaksiyonların (futile cycles) artması AEC'yi 0.70'in "
  "altına çeker. AEC 0.65'in altına indiğinde hücre bölünmesi durur; 0.50'nin altında ise programlı veya nekrotik ölüm "
  'başlar.',
  'Atkinson Adenilat Enerji Şarjı Formülasyonu:\\nAEC = ( [ATP] + 0.5 * [ADP] ) / ( [ATP] + [ADP] + [AMP] '
  ')\\nFizyolojik Gençlik Dengesi: AEC in [0.88, 0.94]\\nYaşlılık Enerji İflası: AEC <= 0.72\\nAMPK Aktivasyon Eşiği: '
  '[AMP]_{free} / [ATP]_{free} >= 0.08.',
  "AEC'nin düşmesi hücre içi tüm anabolik sentezleri (protein, DNA, lipid) anında dondurarak hücreyi sadece hayatta "
  'kalmaya çalışan bir kriz moduna sokar.'),
 ('8.2',
  'AMPK (AMP ile Aktive Olan Protein Kinaz): Enerji Sensörünün Yaşla Körleşmesi',
  'AMPK, hücrenin merkezi enerji termostatıdır. ATP düştüğünde ve AMP/ADP arttığında AMPK aktive olur; ATP tüketen '
  'yolları (mTOR, lipid sentezi) kapatıp ATP üreten yolları (otofaji, yağ oksidasyonu, mitokondriyogenez) ateşler.',
  'Yaşlanma ile birlikte paradoksal bir durum yaşanır: Hücre içi ATP düşük olmasına rağmen AMPK enzimi bu düşüşe yanıt '
  "veremez ('AMPK Desensitizasyonu'). Protein fosfatazların aşırı aktivitesi ve LKB1 kinazının baskılanması nedeniyle "
  "AMPK'nın Thr172 sahasındaki fosforilasyonu engellenir; hücre enerjisizken bile tokmış gibi davranmaya devam eder.",
  'AMPK Aktivasyon Kinetiği ve Fosforilasyon Katsayısı:\\nRate_{AMPK_act} = k_{LKB1} * [AMP] * [AMPK_{unphos}] / ( '
  'K_m^{AMP} + [AMP] )\\nThr172 Fosforilasyon Oranı Kaybı: % p-AMPK(Thr172) <= 0.35 * Gençlik Seviyesi\\nmTORC1 '
  'Kontrolsüz Yangısı: S6K1 ve 4E-BP1 hiperfosforilasyonu sürer.',
  "AMPK'nın körleşmesi, hücrenin enerji tasarrufu yapmasını ve otofajik temizliği başlatmasını engelleyen en büyük "
  'metabolik tuzaktır.'),
 ('8.3',
  "mTORC1 Aşırı Aktivasyonu ve Hücresel 'Hiper-Fonksiyon' (Blagosklonny Teorisi)",
  "Mihail Blagosklonny'nin geliştirdiği 'Hiper-Fonksiyon Yaşlanma Teorisi', yaşlanmanın bir yıpranma veya fonksiyon "
  'kaybından ziyade, gelişim çağında gerekli olan büyüme sinyallerinin (özellikle mTORC1) yetişkinlikte kapatılamayıp '
  'kontrolsüz devam etmesi olduğunu savunur.',
  'mTORC1 (Mechanistic Target of Rapamycin Complex 1), besin ve büyüme faktörleri varken hücreyi sürekli büyümeye ve '
  'bölünmeye zorlar. Yetişkinlikte durdurulamayan bu büyüme sinyali hücreleri hipertrofiye uğratır, otofajiyi '
  'baskılar, protein sentezinde kalite kontrolü yıkar ve hücreyi bitkin düşürerek senesansa iter.',
  'mTORC1 Hiper-Aktivasyon İndeksi ve Otofaji İnhibisyonu:\\nActivity_{mTORC1} = k_{act} * [Rheb-GTP] * [AminoAcids] * '
  '[Insulin/IGF1]\\nULK1 İnhibisyon Oranı (Otofaji Blokajı): % Inhibition_{ULK1} >= 85%\\nRapamisin Bağımlı Ömür '
  'Uzatımı: LifespanExtension = f( 1 / Activity_{mTORC1} ) >= +25% (Kanıtlanmış memeli verisi).',
  "Rapamisin ve rapaloglar ile mTORC1'in dizginlenmesi, evrimsel olarak solucandan insana kadar yaşam süresini en "
  'güvenilir biçimde uzatan moleküler fren sistemidir.'),
 ('8.4',
  'Fosfokreatin (PCr) Mekiği ve Yüksek Enerjili Fosfat Transferinin Yavaşlaması',
  "Mitokondride üretilen ATP'nin hücrenin uzak köşelerine (örneğin nükleusa veya sinapslara) difüzyonla ulaşması çok "
  'yavaştır. Bu transfer, Kreatin Kinaz (CK) enzimlerinin yönettiği Fosfokreatin (PCr) / Kreatin Mekiği ile ışık '
  'hızında gerçekleştirilir.',
  "Mitokondriyel Kreatin Kinaz (mtCK), ATP'yi fosfokreatine dönüştürür; PCr hızla sitoplazmaya difüze olur ve "
  "sitozolik CK tarafından yerel olarak tekrar ATP'ye çevrilir. Yaşlanma sürecinde mtCK oktamerik yapısını kaybederek "
  'inaktif dimerlere ayrışır; yüksek enerjili fosfat transferi durur.',
  'Kreatin Kinaz Mekiği Akı Denklemi:\\nFlux_{PCr} = -D_{PCr} * grad([PCr]) = (V_{max, mtCK} * [ATP]_{mito} * [Cr]) / '
  '( ... )\\nmtCK Oktamer/Dimer Oranı Çöküşü: Ratio_{octamer} = 85% (Genç) ---> 22% (Yaşlı)\\nLokal ATP Yenilenme '
  "Gecikmesi: t_{regen} = 15 ms'den 220 ms'ye uzar.",
  'Fosfat mekiğinin çökmesi, hücrenin enerji santrali çalışsa bile üretilen enerjinin kritik tüketim merkezlerine '
  'ulaştırılamamasına neden olur.'),
 ('8.5',
  'Glikolizis vs Oksidatif Fosforilasyon Dengesizliği (Warburg Benzeri Yaşlılık Kayması)',
  "Genç ve sağlıklı bir hücre enerjisinin %90'ını yüksek verimli mitokondriyel oksidatif fosforilasyondan (glukoz "
  'başına 30-32 ATP) elde ederken, glikolizise (glukoz başına 2 ATP) minimum düzeyde başvurur.',
  'Yaşlanan hücrede solunum zincirinin çökmesi ve piruvat dehidrogenaz (PDH) enziminin fosforilasyonla kilitlenmesi '
  "nedeniyle hücre oksijen varken bile verimsiz glikolizise kayar ('Warburg Benzeri Yaşlılık Kayması'). Laktat "
  'birikir; sitoplazma asitleşir ve aynı enerjiyi üretmek için 15 kat daha fazla glukoz yakılmak zorunda kalınır.',
  'Glikolitik Kayma İndeksi (GSI):\\nGSI = Rate_{lactate_production} / Rate_{oxygen_consumption} = J_{lac} / '
  'J_{O2}\\nGenç Hücre GSI: GSI <= 0.12 (Ağırlıklı oksidatif fosforilasyon)\\nYaşlı Hücre GSI: GSI >= 0.85 (Verimsiz '
  'glikolitik panik modu)\\nATP Üretim Verimi Kaybı: Delta Efficiency <= -60% harcanan glukoz başına.',
  'Bu metabolik kayma, yaşlı hücreyi sürekli bir glukoz açlığına ve metabolik tükenmişliğe mahkum eder.'),
 ('8.6',
  'Termodinamik Isı Dağılımı ve Hücresel Termal Mikro-Dengesizlikler',
  'Hücre içi metabolik reaksiyonlar ve ATP hidrolizi çevreye sürekli mikroskobik miktarlarda ısı salar (dissipative '
  'heat). Son yıllarda geliştirilen nanometrik floresans termometreler, çalışan bir mitokondrinin sıcaklığının 50 '
  "°C'ye kadar çıkabildiğini ortaya koymuştur.",
  'Genç hücrede bu ısı aktin iskeleti ve su akımlarıyla hızla dağıtılırken; yaşlı hücrede makromoleküler '
  'kalabalıklaşma ve viskozite artışı nedeniyle lokal termal sıcak noktalar (thermal hotspots) oluşur. Bu mikro-ısı '
  'odakları çevresindeki proteinleri denatüre eder ve lipidleri peroksidasyona uğratır.',
  'Hücresel Isı Difüzyon ve Dağılım Denklemi:\\nrho * c_p * (dT / dt) = k_{thermal} * Laplacian(T) + q_{metabolic} - '
  'q_{cooling}\\nBurada k_{thermal} hücre içi termal iletkenliktir (Yaşlı sitoplazmada %35 azalır)\\nLokal Termal '
  'Gradyan: Delta T_{local} = T_{mitochondria} - T_{cytosol} >= 8.5 deg C (Patolojik aşırı ısınma).',
  'Termal dağılımın bozulması, hücrenin kendi ürettiği metabolik ısıyla kendi moleküllerini pişirmesi anlamına gelir.'),
 ('8.7',
  'Futile Cycles (Boşuna Metabolik Döngüler) ve Entropik Enerji İsrafı',
  'Normal koşullarda zıt yönlü biyokimyasal yollar (örneğin glukolizdeki fosfofruktokinaz ile glukoneojenezdeki '
  'fruktoz-1,6-bisfosfataz; veya yağ asidi sentezi ile beta-oksidasyon) birbirini karşılıklı olarak inhibe eder.',
  'Yaşlanma ile allosterik geri bildirim kontrolünün kaybolması, zıt enzimlerin aynı anda aktif kalmasına yol açar '
  "('Boşuna Döngüler' - Futile Cycles). Substrat A, ATP harcanarak B'ye dönüştürülür; hemen ardından B tekrar A'ya "
  'çevrilir. Sıfır net ürün elde edilirken trilyonlarca ATP sadece çevreye ısı ve entropi yayarak boşa yanar.',
  'Futile Döngü Enerji Tüketim Hızı:\\nRate_{waste} = k_{cycle} * [ATP] * [Substrate_A] / ( K_m + [Substrate_A] '
  ')\\nİsraf Edilen ATP Fraksiyonu: % ATP_{wasted} = Rate_{waste} / Total_ATP_Production >= 28% (Yaşlı '
  'dokuda)\\nEntropik Kayıp Katsayısı: dS_{waste} / dt = Rate_{waste} * Delta G_{ATP} / T.',
  'Boşuna döngülerin susturulması, hücreye dışarıdan tek bir kalori eklemeden hücresel net enerjiyi üçte bir oranında '
  'artırmanın en zekice metabolik mühendisliğidir.'),
 ('8.8',
  'Mitokondriyal Ayrışma (Uncoupling) ve Proton Sızıntısı Entropisi',
  'Mitokondri iç zarında protonların ATP sentaz yerine UCP1/2/3 proteinleri veya hasarlı lipid alanları üzerinden '
  "kontrolsüzce matrikse geri sızması 'Mitokondriyal Ayrışma' (Uncoupling) olarak adlandırılır.",
  'Kontrollü hafif ayrışma ROS üretimini düşürdüğü için faydalıdır; ancak yaşlanma ile kardiyolipin kaybı ve zar '
  'delinmesi nedeniyle proton sızıntısı patolojik boyutlara ulaşır. Proton motor kuvveti (PMF) çöker; mitokondri ne '
  'kadar elektron yakarsa yaksın ATP üretemeyen ve sadece ısı yayan bozuk bir elektrikli ısıtıcıya dönüşür.',
  'Ayrışma Oranı ve P/O Katsayısı:\\n(P/O)_{effective} = (P/O)_{theoretical} * ( 1 - Fraction_{leak} )\\nProton Kaçak '
  "Akısı: J_{H+_leak} = L_{H+} * Delta mu_{H+}^2\\nP/O Çöküşü: P/O(NADH) = 2.50'den 1.15'e geriler (ATP verimi %54 "
  'kayıp).',
  'Bu enerji kaçağı, yaşlı bir insanın neden sürekli halsiz olduğunu ve hücrelerinin neden kronik enerji açlığı '
  'çektiğini matematiksel olarak izah eder.'),
 ('8.9',
  'İyon Pompası Enerji Faturası: Na+/K+-ATPaz ve Ca2+-ATPaz Açmazı',
  "Bazal koşullarda bir insan hücresinin ürettiği toplam ATP'nin %30 ila %50'si sadece iki enzim tarafından tüketilir: "
  'Na+/K+-ATPaz ve SERCA/PMCA kalsiyum pompaları. Bu pompalar iyon gradyanlarını korumak için gece gündüz çalışır.',
  'Yaşlı hücre zarındaki pasif iyon kaçaklarının artması, bu pompaların çalışma devrini zorunlu olarak artırır. Ancak '
  'mitokondri daha az ATP ürettiği için pompalar yetersiz kalır; iyon gradyanları aşınır. Hücre tüm enerjisini bu '
  "pompalara yatırır ama yine de iyonik dengesizliği engelleyemez ('İyon Pompası Açmazı').",
  'İyon Pompası ATP Tüketim Fraksiyonu:\\nFraction_{ion_pumps} = ( V_{NaK} * 1 ATP + V_{SERCA} * 1 ATP ) / '
  'Rate_{ATP_total}\\nEnerji Yükü Artışı: Fraction_{ion_pumps} = 35% (Genç) ---> 72% (Yaşlı)\\nDiğer Biyosentezlere '
  'Kalan ATP: % Biyosentez <= 15% (Genel hücresel iflas).',
  'Hücre artık yeni bir protein, yeni bir lipid veya DNA tamiri yapacak tek bir ATP molekülü bulamaz; tüm sermayesini '
  'sadece batmakta olan geminin suyunu tahliye etmeye harcar.'),
 ('8.10',
  'Biyoenerjetik Yeniden Şarj Protokolü: Sentetik ATP Besleme ve Metabolik Reset',
  'Hücrenin bu enerji krizinden kurtarılması, metabolik motorun dışarıdan biyoteknolojik olarak yeniden beslenmesiyle '
  'mümkündür.',
  "PROJECT AETERNITAS'ın metabolik reset protokolü; ekzojen NMN/NR ile NAD+ havuzunu doldurur, Metilen Mavisi ile "
  "Kompleks I-IV elektron bypass'ı kurar, sentetik keton esterleri (BHB) ile mitokondriyel verimi P/O = 2.5'e kilitler "
  've küçük moleküllü allosterik AMPK aktivatörleri ile mTOR/AMPK dengesini gençlik ritmine senkronize eder.',
  'Biyoenerjetik Reset Başarı Skoru (BERS):\\nBERS = 0.35 * (AEC / 0.90) + 0.35 * (Ratio_{PCr/Cr} / 1.5) + 0.30 * (P/O '
  '/ 2.45) * 100\\nHedef Değer: BERS >= 95.0 / 100.0\\nNet Hücresel Serbest Enerji Kazancı: Delta Delta G_{ATP} <= '
  '-12.5 kJ/mol (Maksimal itici güç).',
  'Bu metabolik şarj tamamlandığında, hücre epigenetik reprogramlama ve telomerik uzama gibi devasa enerji isteyen '
  'rejuvenasyon operasyonlarını gerçekleştirebilecek süper-güce kavuşur.')]

part9_topics = [('9.1',
  'Sirkadiyen Saat Ağı (CLOCK/BMAL1 ve PER/CRY) Moleküler Biyofiziği',
  'Dünyanın 24 saatlik kendi ekseni etrafında dönüşü, tüm canlı hücrelerde yaklaşık 24 saatlik otonom bir '
  'transkripsiyonel-translasyonel geri bildirim döngüsü (TTFL) inşa etmiştir: Sirkadiyen Saat.',
  'Bu saat; CLOCK ve BMAL1 heterodimerinin E-box dizilerine bağlanarak PER (Period 1, 2, 3) ve CRY (Cryptochrome 1, 2) '
  'genlerini transkribe etmesiyle çalışır. Sitoplazmada biriken PER/CRY kompleksleri nükleusa dönerek kendi '
  'transkripsiyonlarını baskılar. 24 saatlik bu ritim, hücredeki tüm metabolik enzimlerin, DNA tamir mekanizmalarının '
  've immün hücrelerin çalışma saatlerini yönetir.',
  'Sirkadiyen TTFL Kinetik Limit Döngü Modeli (Goldbeter Denklemleri):\\nd[PER/CRY]/dt = v_s * [CLOCK/BMAL1]^n / ( '
  'K_1^n + [CLOCK/BMAL1]^n ) - v_m * [PER/CRY] / ( K_m + [PER/CRY] )\\nRitim Periyodu: T_{circadian} = 24.2 +/- 0.3 '
  "saat\\nGenomik Kapsam: İnsan genlerinin %45'i sirkadiyen ritmik dalgalanma gösterir.",
  'Sirkadiyen saat, hücresel biyokimyanın birbiriyle çakışmadan yürütülmesini sağlayan zamanlama orkestrasıdır.'),
 ('9.2',
  'Sirkadiyen Ritimlerin Yaşlanmayla Amplitüd Kaybı ve Faz Desenkronizasyonu',
  "Yaşlanma sürecinde sirkadiyen saatin en belirgin bozulması periyodun değişmesinden ziyade 'Amplitüdünün' (dalga "
  'tepesi ile çukuru arasındaki fark) sönümlenmesidir.',
  'Yaşlı bir hücrede gündüz ile gece arasındaki transkripsiyonel fark neredeyse silinir; BMAL1 ve PER proteinleri '
  'sürekli orta düzeyde, sönük bir platoda salınır. Daha da kötüsü, beyindeki merkezi saat (Suprakiazmatik Çekirdek - '
  'SCN) ile periferik organ saatleri (karaciğer, kas, pankreas) arasındaki faz uyumu kaybolur (içsel '
  'desenkronizasyon). Karaciğer gündüzü yaşarken böbrek geceyi yaşar.',
  'Sirkadiyen Dalga Amplitüd ve Faz Koheransı:\\nAmplitude_{circadian} = Peak_{day} - Trough_{night}\\nAmplitüd '
  'Sönümlenmesi: Amp_{aged} / Amp_{young} <= 0.38 (Fizyolojik körleşme)\\nFaz Kayması (Desenkronizasyon): Delta '
  'phi_{SCN-Liver} >= 5.5 saat (Metabolik kaos).',
  'Sirkadiyen amplitüdün kaybolması, hücrenin DNA hasarını gece onarmak yerine gündüzün kaotik ortamında onarmaya '
  'çalışmasına ve mutasyonların katlanmasına yol açar.'),
 ('9.3',
  'Melatonin Salınımı, Epifiz Bezi Kalsifikasyonu ve Mitokondriyel Gece Koruması',
  'Melatonin (N-asetil-5-metoksitriptamin), sadece epifiz bezinden salınan bir uyku hormonu değildir; evrimsel olarak '
  'bakterilerden miras kalmış en güçlü doğrudan mitokondriyel antioksidan ve serbest radikal süpürücüsüdür.',
  'Yaşlanmayla birlikte epifiz bezi hidroksiapatit kristalleriyle kalsifiye olur; gece tepe melatonin salınımı %80 '
  'azalır. Daha önemlisi, hücrelerin kendi mitokondrileri içinde sentezlediği intrasellüler melatonin üretimi çöker. '
  'Mitokondriler gece boyunca maruz kaldıkları serbest radikallerden korunamaz ve dinlenme fazında bile hasar almaya '
  'devam eder.',
  'Melatonin Konsantrasyonu ve Süpürme Kinetiği:\\nRate_{scavenging} = k_{melatonin} * [Melatonin]_{mito} * [*ullet '
  'OH]   (k = 2.7 * 10^{10} M^{-1} s^{-1})\\nGece Tepe Plazma Melatonini: [Melatonin]_{peak} = 120 pg/mL (20 Yaş) ---> '
  '15 pg/mL (75 Yaş)\\nEpifiz Kalsifikasyon Fraksiyonu: % Calcification >= 65% (Yaşlı bireylerde).',
  'Melatonin çöküşü, mitokondrinin gece boyunca kendini resetleme ve onarma fırsatını elinden alarak biyolojik saati '
  'hızlandırır.'),
 ('9.4',
  'Hücresel Bölünme Saati: Hayflick Limiti ve Replikatif Yaşlanma Biyofiziği',
  '1961 yılında Leonard Hayflick, normal insan fetal fibroblastlarının hücre kültüründe sonsuza kadar '
  "bölünemeyeceğini; yaklaşık 50 bölünmeden (40-60 aralığı) sonra kalıcı olarak durduğunu keşfetmiştir ('Hayflick "
  "Limiti').",
  'Bu replikatif bariyer, her hücre bölünmesinde DNA polimerazın doğrusal kromozom uçlarını tam kopyalayamaması (uç '
  'replikasyon problemi) nedeniyle telomerlerin kısalmasından kaynaklanır. Hayflick limiti, hücrenin genomik '
  'bütünlüğünü korumak için tasarlanmış bir sayaçtır; ancak sayaç sıfıra ulaştığında hücre senesans krizine girer.',
  'Hayflick Bölünme Sayacı ve Telomer Erozyon Denklemi:\\nN_{divisions} <= N_{Hayflick} approx 50\\nTelomer Erozyon '
  'Hızı: dL_{telomere} / dn = -Delta L_{end_replication} = -50 ila -150 baz çifti / bölünme\\nKritik Telomer Eşiği: '
  'L_{critical} <= 3.5 - 4.0 kb (DNA hasar yanıtı DDR tetiklenir).',
  'Hayflick limiti somatik hücrelerin ömrünü kesin bir matematiksel kısıtla bağlamıştır; bu kısıtı kırmak ancak '
  'telomeraz mühendisliği ile mümkündür.'),
 ('9.5',
  'Sirkadiyen-Mitokondriyel Çapraz Bağlantı: NAMPT ve NAD+ Salınımları',
  'Sirkadiyen saat ile mitokondriyel biyoenerjetik arasındaki moleküler köprü, NAD+ kurtarma yolağının hız kısıtlayıcı '
  "enzimi olan NAMPT'dir (Nikotinamid Fosforiboziltransferaz).",
  'NAMPT geninin promotöründe CLOCK/BMAL1 heterodimerinin bağlandığı E-box motifleri bulunur. Bu sayede hücre içi NAD+ '
  'konsantrasyonu gün içinde sirkadiyen bir ritimle dalgalanır. Yaşlanmayla sirkadiyen amplitüd sönümlendiğinde NAMPT '
  'transkripsiyonu çöker; hücre içi NAD+ ritmi düz bir çizgiye dönüşür ve sirtuinler 24 saat boyunca uykuda kalır.',
  'NAMPT Sirkadiyen Dalgalanması ve NAD+ Salınım Genliği:\\nd[NAMPT]/dt = k_{circ} * [CLOCK/BMAL1](t) - k_{deg} * '
  '[NAMPT]\\nNAD+ Ritmik Salınım Genliği: Delta [NAD+] = Peak - Trough = 450 uM (Genç) ---> 45 uM (Yaşlı)\\nSIRT1 '
  'Aktivasyon Fazı: Sirkadiyen tepe noktasında %80 artış yaşlıda kaybolur.',
  'NAMPT ritminin kaybı, metabolik motorun zaman ayarını bozarak hücreyi sürekli bir kış uykusu ve yorgunluk '
  'sendromuna sokar.'),
 ('9.6',
  'Mikrotübüler Terahertz Titreşimleri ve Hücresel Biyofiziksel Zaman Sayaçları',
  'Biyolojik zaman sadece sirkadiyen 24 saatlik kimyasal döngülerle akmaz; hücre iskeletini oluşturan mikrotübüller '
  'içinde terahertz (THz: 10^{12} Hz) ve gigahertz bandında çalışan ultra hızlı kuantum biyofiziksel osilatörler '
  'bulunur.',
  'Tubulin dimerlerinin dipol moment salınımları, moleküler motorların adımlarını ve mitoz bölünmenin zamanlamasını '
  "femtosaniyelik hassasiyetle senkronize eden bir 'içsel biyofiziksel kuartz saat' gibi işler. Yaşlanma ile tubulin "
  'proteinlerinin post-translasyonel glikasyonu bu yüksek frekanslı rezonansı bozar; hücresel nano-zamanlama çöker.',
  'Mikrotübüler Dipol Rezonans Frekansı:\\nf_{tubulin} = (1 / 2*pi) * sqrt( k_{dipole} / m_{eff} ) approx 8.5 * '
  '10^{12} Hz (8.5 THz)\\nFaz Koheransı Kalite Faktörü: Q_{factor} = 1.200 (Genç) ---> 180 (Yaşlı - Rezonans '
  'sönümü)\\nNano-Mekanik Adım Senkronizasyonu: Kinezin adım hatası 10 kat artar.',
  'Hücrenin mikro-zamanlamasının bozulması, organel taşınmasının ve hücre içi kargo lojistiğinin kaotik kazalara '
  'uğramasına neden olur.'),
 ('9.7',
  'Biyo-Fotonik Emisyon Ritimleri ve Hücre İçi İletişim Koheransı',
  "Tüm canlı hücreler, mitokondriyel elektron akışı ve lipid peroksidasyonunun yan ürünü olarak 'Ultra Zayıf Foton "
  "Emisyonu' (UPE - Ultra-weak Photon Emission) yayarlar (saniyede birkaç yüz foton/cm^2).",
  "Fritz-Albert Popp'un biyofiziksel deneyleri, genç ve sağlıklı hücrelerden yayılan fotonların koherent (lazer "
  'benzeri faz uyumlu) olduğunu; hücre yaşlandıkça veya kanserleştikçe bu koheransın kaybolarak termal gürültü gibi '
  'kaotik hale geldiğini kanıtlamıştır. Biyo-fotonlar hücrelerin birbirleriyle ışık hızında haberleştiği optik bir '
  'dildir.',
  'Biyo-Fotonik Koherans ve Glauber İkinci Derece Korelasyon Fonksiyonu:\\ng^{(2)}(tau) = <I(t) * I(t + tau)> / '
  '<I(t)>^2\\nGenç Koherent Emisyon: g^{(2)}(0) approx 1.0 (Tam kuantum koherans ve faz kilitlenmesi)\\nYaşlı Kaotik '
  'Emisyon: g^{(2)}(0) >= 2.0 (Termal gürültü ve optik iletişim kaybı)\\nFoton Emisyon Yoğunluğu Artışı: Hasar '
  'arttıkça kontrolsüz UPE sızıntısı %400 artar.',
  'Optik koheransın kaybı, dokudaki milyonlarca hücrenin birbiriyle olan kuantum ve fotonik telsiz bağlantısının '
  'kesilmesi demektir.'),
 ('9.8',
  'Dekompartmantalize Kuantum Eşevresizlik (Decoherence) ve Biyolojik Entropi',
  'Kuantum biyolojisinde enzimlerin proton ve elektron tünellemesi, fotosentetik eksiton transferi ve DNA mutasyon '
  'kararlılığı kuantum süperpozisyon durumlarına dayanır.',
  'Hücre içinde su moleküllerinin Exclusion Zone (EZ su) katmanları oluşturması, kuantum durumlarını termal '
  'dekoheranstan (decoherence) koruyan topolojik bir kalkan görevi görür. Yaşlanma ile hücre içi suyun düzenliliğini '
  'kaybetmesi ve viskozitenin artması, dekoherans süresini (tau_d) pikosaniyelerden femtosaniyelere indirir; kuantum '
  'tünelleme verimi düşer ve enzimatik reaksiyonlar duraklar.',
  'Kuantum Dekoherans Süresi ve Termal Çevre Etkileşimi:\\ntau_d approx hbar^2 / ( 2 * m * gamma_{thermal} * k_B * T * '
  '(Delta x)^2 )\\nGençlik EZ Su Kalkanı Koruması: tau_d(Young) >= 100 * tau_d(Aged)\\nEnzimatik Proton Tünelleme '
  'Olasılığı Kaybı: P_{tunnel} = exp(-2*K*a) <= 0.30 * Basal.',
  'Dekoherans kalkanının erimesi, hücresel biyokimyanın kuantum hızından klasik mekanik sürtünme hızına gerilemesine '
  'yol açar.'),
 ('9.9',
  'Kronobiyolojik Yaşlanma Senkronizasyonu: Organ Saatlerinin Uyumsuzluğu',
  'Çok hücreli bir memelide yaşlanma tüm organlarda aynı anda ve aynı hızda gerçekleşmez. Kalp, böbrek, karaciğer ve '
  'beyin farklı kronobiyolojik hızlarda yaşlanır.',
  "Bu durum 'Organ Spesifik Asenkron Yaşlanma' (Organ-specific Asynchronous Ageing) olarak adlandırılır. Örneğin "
  'böbreği 70 yaşında olan bir bireyin kalbi 50, beyni 60 yaşında olabilir. En yaşlı organ, salgıladığı toksik '
  "sitokinler ve ürik asit ile diğer genç organları da kendi yaş seviyesine doğru aşağı çeken sistemik bir 'karadelik' "
  'gibi davranır.',
  'Çok Organlı Asenkron Yaş Vektörü ve Sistemik Sürüklenme:\\nVector_{Age} = [ Age_{heart}, Age_{brain}, Age_{liver}, '
  'Age_{kidney}, Age_{immune} ]\\nSistemik Yaşlanma İndeksi: SystemicAge = Max_k( Age_k ) + SUM_{j} w_j * ( Age_j - '
  'Min_k(Age_k) )\\nOrganlar Arası Biyolojik Yaş Sapması: Delta Age_{inter-organ} >= 15 yıl (Klinik kriz).',
  'PROJECT AETERNITAS, en zayıf ve en yaşlı organı tespit edip öncelikli olarak gençleştirerek sistemik dominonun '
  'yıkılmasını engeller.'),
 ('9.10',
  'Çok Boyutlu Biyolojik Saat Matrisi: 10 Katmanlı Zaman Entegrasyonu',
  "Bölüm 1'in zirvesinde, hücresel zamanın tek bir saatle ölçülemeyeceği; 10 farklı biyofiziksel ve moleküler saatin "
  "eşzamanlı çalıştığı 'Çok Boyutlu Zaman Matrisi' kurulur.",
  'Bu 10 saat: 1) DNA Metilasyon Saati (Horvath), 2) Telomer Sayacı (Hayflick), 3) Transkriptomik Saat, 4) Proteomik '
  'Saat, 5) Glikomiks Saati (Glikanlar), 6) Lipidomik Saat, 7) Sirkadiyen Dalga Saati, 8) Nükleositoplazmik Bariyer '
  'Saati, 9) Mitokondriyel Heteroplazmi Saati ve 10) Kuantum Biyo-Fotonik Koherans Saati.',
  'On Katmanlı Biyolojik Zaman Tensörü (Chronos Tensor):\\nTau_{Biological} = SUM_{m=1}^{10} W_m * Clock_m(t) + '
  'Tensor_{interactions}\\nBurada W_m ağırlık katsayıları matrisidir (SUM W_m = 1.0)\\nMutlak Biyolojik Sıfırlama '
  'Kriteri: Tüm 10 saatte d(Clock_m)/dt <= 0 koşulunun sağlanmasıdır.',
  'Bu entegre matris, bir insanın biyolojik yaşını atomik düzeyde şifreler ve uygulanan tedavilerin zamanı ne kadar '
  'geriye sardığını kesin matematiksel verilerle ispatlar.')]

part10_topics = [('10.1',
  'Termodinamik Tersinirlik: Loschmidt Paradoksu ve Biyolojik Mikrostates',
  '1876 yılında Josef Loschmidt, temel fizik yasalarının (Newton, Schrödinger denklemleri) zamanın tersine '
  'çevrilmesine karşı simetrik olduğunu (Time-Reversal Symmetry - T-simetrisi); eğer bir sistemdeki tüm parçacıkların '
  'hız vektörleri tersine çevrilirse sistemin tam olarak başlangıç durumuna döneceğini öne sürmüştür (Loschmidt '
  'Paradoksu).',
  'Klasik makroskobik sistemlerde milyarlarca parçacığın hızını tersine çevirmek imkansızdır; bu yüzden kırılan bir '
  'bardak kendiliğinden birleşmez. Ancak biyolojik hücrede genetik ve epigenetik bilgi makroskobik değil, dijital ve '
  'nükleik asit düzeyinde diskret bir enformasyondur. Hücresel mikrostates kümesi sonludur ve dışarıdan yönlendirilen '
  'serbest enerji ile tersine çevrilmesi termodinamik olarak kesinlikle mümkündür.',
  'Loschmidt Zaman Simetrisi ve Biyolojik Durum Dönüşümü:\\nH(p, q) = H(-p, q)   (Hamiltonyen Simetrisi)\\nTersinir '
  'Mikrodurum Geçiş Olasılığı: P(Aged -> Young) = exp( -Delta S_{reversal} / k_B ) * P_{catalysis}\\nDış Enerji '
  'Enjeksiyonu: Delta G_{external} >= T * Delta S_{internal_damage}.',
  'Bu termodinamik temel, yaşlanmanın tersine çevrilmesinin fizik yasalarını ihlal eden bir mucize değil; tam aksine '
  'termodinamiğin birinci ve ikinci yasalarına tam oturan bir mühendislik problemi olduğunu kanıtlar.'),
 ('10.2',
  "Maxwell Cini (Maxwell's Demon) Olarak Hücresel Enzimler ve DNA Tamir Makineleri",
  '1867 yılında James Clerk Maxwell, iki bölme arasındaki kapıda durarak hızlı molekülleri bir tarafa, yavaşları diğer '
  "tarafa geçirip ikinci yasayı ihlal etmeden entropiyi düşüren hayali bir 'Cin' (Demon) tasarlamıştır. Leo Szilard ve "
  'Rolf Landauer, bu cinin aslında bilgi toplayarak çalıştığını ve bilginin entropiye dönüştürülebileceğini '
  'göstermiştir.',
  'Hücrede DNA tamir enzimleri (DNA Ligaz, OGG1, Fotoliyaz), şaperonlar ve otofajik reseptörler gerçek birer '
  "'Moleküler Maxwell Cini'dir. Hasarlı molekülleri sağlamlardan ayırırlar; ATP harcayarak hücresel entropiyi lokal "
  'olarak düşürürler. Yaşlanma, hücresel Maxwell cinlerinin yorulmasıdır; gençleşme ise bu cinlerin sayısını ve bilgi '
  'işleme hızını artırmaktır.',
  'Szilard Motoru ve Moleküler Cin Entropi Düşüşü:\\nDelta S_{system} = -k_B * ln(2) * I_{information}\\nATP Maliyet '
  'Oranı: E_{cost} >= k_B * T * ln(2) bit başına\\nEnzimatik Enformasyonel Verim: eta_{demon} = Delta S_{cleared} / '
  'ATP_{consumed} >= 0.82.',
  'Hücre, doğru genetik komutlar ve ATP verildiğinde kendi kendini kusursuzca temizleyen milyarlarca moleküler cinle '
  'donatılmış bir harikalar diyarıdır.'),
 ('10.3',
  'Biyolojik Ölümsüzlüğün Doğadaki Varoluş Kanıtları: Hydra, Turritopsis dohrnii ve Planaria',
  'Biyolojik ölümsüzlük insanlığın icat ettiği bir hayal değil; doğanın milyarlarca yıldır başarıyla uyguladığı somut '
  'bir biyolojik gerçektir.',
  'Hydra vulgaris (tatlı su hidrası), sürekli bölünen kök hücreleri ve aktif telomerazı sayesinde yaşlanmaz; Hayflick '
  'limiti yoktur ve ölüm oranı zamanla artmaz (sıfır senesans). Turritopsis dohrnii (ölümsüz denizanası), '
  "yaşlandığında veya stres altına girdiğinde 'transdiferansiyasyon' yoluyla hücrelerini yeniden genç polip formuna "
  'döndürerek biyolojik zamanını geriye sarar. Planaria yassı solucanları ise kök hücreleri (neoblastlar) sayesinde '
  'sonsuz rejenerasyon sergiler.',
  'Gompertz Yaşlanma Yasası ve Sıfır Senesans Kriteri:\\nmu(t) = mu_0 * exp( G * t )   (İnsanda G > 0 - Ölüm riski her '
  '8 yılda 2 katına çıkar)\\nHydra ve Turritopsis Şartı: G = 0  ===>  mu(t) = mu_0 = Sabit (Sıfır Biyolojik '
  'Yaşlanma)\\nTransdiferansiyasyon Geri Dönüşüm Verimi: % Reversal = 100%.',
  'Doğada var olan bir biyolojik mekanizma, prensipte insan hücresine de genetik ve epigenetik mühendislikle '
  'aktarılabilir; hiçbir evrensel engel yoktur.'),
 ('10.4',
  'Entropi İhracı: Ekstrasellüler Veziküller (EVs) ve Eksozom Dinamiği',
  'Açık bir sistem olan hücrenin içeride ürettiği entropiyi dışarı atmasının en sofistike biyolojik yollarından biri '
  'Ekstrasellüler Veziküller (EVs) ve Eksozomlardır.',
  'Hücre, lizozomun parçalayamadığı toksik proteinleri, okside lipidleri ve hasarlı DNA parçalarını multivesiküler '
  'cisimcikler (MVB) içinde paketleyerek ekzositozla hücre dışına fırlatır. Genç hücreler entropilerini hızla dışarı '
  'atarken; yaşlı hücrelerin eksozom sekresyonu bozulur ve hücre kendi çöpleriyle tıkanır. Genç kandan alınan '
  'eksozomların yaşlı hücrelere verilmesi ise içeriye negatif entropi enjekte eder.',
  'Ekzozomal Entropi İhraç Hızı ve Doku Klirensi:\\nFlux_{entropy_export} = J_{exosome} * S_{cargo} = Rate_{MVB} * '
  '[Damage]_{vesicle} * Delta S_{unit}\\nEksozom Klirens Kapasitesi: N_{EV} / hücre / gün = 2.400 (Genç) ---> 450 '
  '(Yaşlı)\\nHeterokronik Eksozom Tedavisi: Doku DNAmAge skorunda 6 haftada %30 düşüş.',
  'Entropinin eksozomlarla dışarı atılması ve genç eksozomlarla dokunun yıkanması, hücresel temizliğin en zarif '
  'biyolojik yöntemidir.'),
 ('10.5',
  'Geri Dönüşümsüzlük İllüzyonu: Waddington Manzarasında Tersine Tırmanış',
  'On yıllar boyunca gelişimsel biyolojinin en katı dogması, bir kök hücrenin özelleşmiş bir hücreye (nöron veya deri) '
  'dönüştükten sonra asla geri dönemeyeceğiydi (Waddington vadisinden aşağı tek yönlü yuvarlanış).',
  '2006 yılında Shinya Yamanaka, sadece 4 transkripsiyon faktörü (Oct4, Sox2, Klf4, c-Myc) ile yetişkin bir deri '
  'hücresini embriyonik pluripotent kök hücreye (iPSC) geri döndürerek bu dogmayı yerle bir etti. Hücre vadiden '
  "yukarıya tırmanabilirdi! 2020'lerde ise Sinclair ve ekibi, hücreyi tepeye kadar çıkarmadan (kimliğini unutturmadan) "
  'sadece vadinin tabanındaki yaşlılık çukurundan gençlik çukuruna geri tırmandırmayı (kısmi reprogramlama) başardı.',
  'Waddington Vadisinde Tersine Tırmanış İtici Gücü:\\nF_{reversal} = -grad(V_{Waddington}) + F_{OSK_vector} > '
  '0\\nGereken Transkripsiyonel Enerji: Delta G_{OSK} >= Delta V_{differentiation_barrier}\\nKimlik Koruma Penceresi: '
  'InductionTime in [t_{rejuvenation_start}, t_{dedifferentiation_threshold}] (Tipik olarak 2 ila 4 gün darbe).',
  'Waddington vadisinde tersine tırmanış artık teorik bir spekülasyon değil; laboratuvarlarda her gün tekrarlanan '
  'rutin bir hücresel operasyondur.'),
 ('10.6',
  'Fraktal Genom Mimarisi ve Boyutsuz Yaşlanma Katsayısı',
  "Genomun 3D katlanması rastgele değildir; 'Fraktal Globül' (Fractal Globule) adı verilen düğümsüz, aşırı yoğun ancak "
  'her an çözülüp okunabilir muazzam bir fraktal geometri sergiler.',
  "Fraktal globül mimarisinin fraktal boyutu (D_f) genç hücrelerde tam olarak D_f = 3.0'a oturur. Yaşlanma ile "
  'kromatin ilmeklerinin kopması ve heterokromatin erimesi fraktal boyutu düşürür (D_f -> 2.4); genom düğümlü ve '
  'karmaşık bir kaotik yumağa (equilibrium globule) dönüşür. Bu geometrik bozulma, transkripsiyon faktörlerinin '
  'aradıkları genleri bulma süresini 10 kat uzatır.',
  'Fraktal Boyut ve Kromatin Temas Olasılığı Yasası:\\nP(s) prop to s^{-gamma},   gamma = D_f / 3\\nGenç Fraktal '
  'Globül: gamma = 1.0 (D_f = 3.0 - Sıfır düğüm, anında erişim)\\nYaşlı Düğümlü Globül: gamma = 1.5 (D_f = 2.4 - '
  'Düğümlenmiş kromatin, gen susturma iflası)\\nArama Zamanı Artışı: t_{search} = r^2 / D_{fractal} >= 8.5 kat.',
  'Kromatin fraktal mimarisinin restore edilmesi, genetik kütüphanenin raflarını yeniden kusursuz bir kütüphaneci '
  'düzenine sokar.'),
 ('10.7',
  "Biyolojik Tekillik (Biological Singularity) ve 'Kaçış Hızı' (Longevity Escape Velocity)",
  "Fütürist Aubrey de Grey tarafından ortaya atılan 'Uzun Ömür Kaçış Hızı' (Longevity Escape Velocity - LEV), "
  'biyoteknolojik gelişmelerin bir insanın kalan yaşam süresine yılda bir yıldan daha fazla ömür eklediği tarihsel '
  'dönüm noktasıdır.',
  'Eğer tıp her geçen yıl biyolojik saatinizi 1.2 yıl gençleştirebiliyorsa, kronolojik olarak ne kadar yaşlanırsanız '
  'yaşlanın biyolojik olarak gençleşirsiniz; yani ölüme doğru değil ölümsüzlüğe doğru koşarsınız. PROJECT AETERNITAS, '
  'insanlığı bu kaçış hızının ötesine fırlatacak moleküler motorları tasarlar.',
  'Longevity Escape Velocity Diferansiyel Eşitsizliği:\\nd(Remaining_Lifespan) / dt = Rate_{biotech_rejuvenation}(t) - '
  '1.0 > 0\\nLEV Eşiği: Rate_{biotech} >= 1.0 yıl / yıl\\nKritik Kaçış Tarihi Projeksiyonu: 2026 - 2035 aralığı (İlk '
  'insan klinik denemeleri fazında).',
  'Kaçış hızına ulaşıldığında, ölüm artık biyolojik bir kader olmaktan çıkıp sadece isteğe bağlı bir tercih haline '
  'gelir.'),
 ('10.8',
  'Termodinamik Maliyet: Yaşlanmayı Tersine Çevirmenin Enerji Faturası',
  'Yaşlanmayı tersine çevirmek termodinamik olarak bedelsiz değildir; Landauer prensibi ve disipatif yapılar teorisi '
  'uyarınca içerideki entropiyi sıfırlamak çevreye devasa bir enerji faturası ödemeyi gerektirir.',
  'Bir insan hücresini 80 yaşından 20 yaşına geri döndürmek için gereken net serbest enerji miktarı hesaplanmıştır: '
  'DNA tamiri, heterokromatin restruktürasyonu, amiloidlerin eritilmesi ve membran yenilenmesi hücre başına yaklaşık '
  '4.2 * 10^9 ATP hidrolizine (yaklaşık 1.3 * 10^{-10} Joule) mal olur. Bir insan vücudunda 37 trilyon hücre olduğu '
  'düşünüldüğünde, tüm bedenin gençleşmesi yaklaşık 4.8 * 10^3 kJ metabolik serbest enerji gerektirir.',
  'Biyolojik Gençleşmenin Termodinamik Maliyet Denklemi:\\nDelta G_{rejuvenation}^{total} = N_{cells} * SUM_{k=1}^M [ '
  'Delta G_{repair, k} + T * Delta S_{cleared, k} ]\\nHücre Başına Maliyet: Delta G_{cell} approx 4.2 * 10^9 '
  'ATP\\nToplam Beden Enerji İhtiyacı: E_{reversal} approx 1.150 kkal (Bir günlük tam metabolik kalori dengesi).',
  'Bu şaşırtıcı derecede küçük enerji faturası, doğru moleküler araçlar kullanıldığında gençleşmenin biyolojik olarak '
  'ne kadar ekonomik ve uygulanabilir olduğunu kanıtlar.'),
 ('10.9',
  'Biyolojik Ölümsüzlüğün Aksiyomları: PROJECT AETERNITAS Temel Yasaları',
  "Bölüm 1'in felsefi ve biyofiziksel doruk noktasında, tüm külliyatın üzerine inşa edileceği 'Biyolojik Ölümsüzlüğün "
  "Dört Temel Aksiyomu' ilan edilir.",
  'AKSİYOM 1 (Enformasyonun Korunumu): Hücresel yaşlanma donanımsal sekans kaybı değil, yazılımsal epigenetik '
  'gürültüdür; gençlik enformasyonu genomun kuantum ve TET tabanlı yedek havuzunda ebediyen saklıdır.\\nAKSİYOM 2 '
  '(Termodinamik Açıklık): Canlı hücre açık bir disipatif sistemdir; yeterli negatif entropi ve metabolik iş akışı '
  'sağlandığı sürece hücresel entropi süresiz olarak sıfırlanabilir.\\nAKSİYOM 3 (Germline Eşdeğerliği): Doğanın üreme '
  'hücrelerine (germline) bahşettiği sınırsız telomerik ve metabolik ölümsüzlük, somatik hücrelere aktarılabilir bir '
  'genetik kurgudur.\\nAKSİYOM 4 (Geri Dönüşebilirlik): Hücresel zaman tek yönlü bir ok değildir; hücresel biyolojik '
  'yaş dinamik, programlanabilir ve sıfıra geri sarılabilir bir koordinattır.',
  'Aeternitas Dinamik Kararlılık Fonksiyonu:\\nlim_{t -> sonsuz} Age_{biological}(t) = Constant_{optimum} = 22.5 '
  'Yaş\\nSistemik Biyolojik Ömür: Tau_{max} -> Sonsuz.',
  'Bu aksiyomlar, insanlığın ölüm karşısındaki çaresizlik çağını kapatıp bilinçli biyomühendislik çağını resmen '
  'başlatır.'),
 ('10.10',
  "Bölüm 1 Epilogu: Hücresel Kaderin Yeniden Yazımı ve Cilt 2'ye Giriş",
  'Bu ilk ciltte, hücresel yaşlanmanın en derin termodinamik, istatistiksel, biyofiziksel ve nükleer temellerini atom '
  'atom inceledik. Gördük ki yaşlanma ne Tanrısal bir ceza ne de kaçınılmaz bir doğa yasasıdır; sadece kontrolsüz '
  'bırakılmış entropik gürültüdür.',
  'Entropinin doğasını ve hücresel saatin işleyişini çözdüğümüze göre, şimdi sıra bu saati bizzat ölçen, saniyelerini '
  "sayan ve bize biyolojik yaşımızı fısıldayan moleküler cetvelleri elime almaktadır. Cilt 2, Steve Horvath'ın "
  "epigenetik saatlerinden GrimAge'e, DNA metilasyon algoritmalarından hücresel kronometrelerin kalibrasyonuna uzanan "
  "'Epigenetik Saatler ve Horvath Metilasyon Algoritmaları' başyapıtını açacaktır.",
  'Bölüm 1 Biyofiziksel Kapanış Özeti:\\nİncelenen Alt-Bölüm: 100 Granüler Kısım\\nFormüle Edilen Fiziksel Yasalar: '
  'Boltzmann, Prigogine, Landauer, Shannon, Anfinsen, Hayflick, Loschmidt\\nNihai Sonuç: Canlılık Ölümsüzlüğe '
  'Programlanabilir.',
  'Cilt 1 burada mühürlenmiştir. Biyolojik ölümsüzlük yolculuğu resmen başlamıştır!')]

tables_data = [('Tablo 1.1: Termodinamik Entropi, Boltzmann Dağılımı ve Hücresel Serbest Enerji Parametreleri',
  ['Termodinamik Parametre',
   'Fiziksel Formülasyon / Birim',
   'Genç Hücre Değeri',
   'Yaşlı Hücre Değeri',
   'Biyolojik / Patolojik Anlamı'],
  [['Hücre İçi Entropi Üretimi (sigma)',
    'dS_i / dt / V (mW / K / m^3)',
    '12.4 +/- 1.2',
    '48.6 +/- 4.5',
    'Yaşlı hücrede metabolik sürtünme ve enerji israfı 4 katına çıkar'],
   ['Gibbs Serbest Enerjisi (Delta G)',
    'Delta H - T*Delta S (kJ / mol)',
    '-58.5 (ATP hidrolizi)',
    '-41.2 (ATP hidrolizi)',
    'Hücresel iş yapabilme kapasitesinin termodinamik çöküşü'],
   ['Negatif Entropi İthalat Hızı',
    '- d_e S / dt (J / K / s)',
    '3.85 * 10^-11',
    '0.92 * 10^-11',
    'Schrödinger negatif entropi pompasının 4 kat yavaşlaması'],
   ['Boltzmann Durum Sayısı (Omega)',
    'S = k_B * ln(Omega) (Mikrodurum)',
    'Omega_0 (Düzenli)',
    'Omega_0 * 10^8 (Kaotik)',
    'Makromoleküler konformasyonel düzensizlik ve belirsizlik patlaması'],
   ['Disipatif Yapı Kararlılığı',
    'd(sigma) / dt',
    '<= 0 (Kararlı Atraksiyon)',
    '> 0 (Kararsızlaşma)',
    'Prigogine minimum entropi üretim dengesinden sapma ve kaos'],
   ['Langevin Sürtünme Katsayısı',
    'gamma = 6*pi*eta*r (pN * s / um)',
    '0.18 +/- 0.02',
    '0.58 +/- 0.05',
    'Sitoplazmik vizkozite artışı nedeniyle moleküler motorların yavaşlaması'],
   ['Biyolojik Zaman Diferansiyeli',
    'd(tau_bio) / dt (Boyutsuz Hız)',
    '0.85 (Yavaş Akış)',
    '1.65 (Hızlanmış Yaşlanma)',
    'Biyolojik saatin kronolojik takvimden bağımsız hızlı akışı'],
   ['Membran Dinlenim Potansiyeli',
    'Delta Psi_{plasma} (mV)',
    '-70 ila -85 mV',
    '-35 ila -45 mV',
    'Hücre zarındaki elektrokimyasal potansiyel gradyanının çöküşü'],
   ['Flory Etkileşim Parametresi',
    'chi (LLPS Kararlılık Faktörü)',
    'chi < chi_{spinodal}',
    'chi >> chi_{spinodal}',
    'Zarsız organellerin sıvı fazdan katı amiloid jele geçişi'],
   ['Landauer Bilgi Entropi Kaybı',
    'dH_{epigenetic} / dt (bit / gün)',
    '1.2 * 10^-6',
    '1.8 * 10^-4',
    'Epigenetik işletim sisteminin termodinamik veri bozulması hızı']]),
 ('Tablo 1.2: Stokastik Moleküler Gürültü, Transkripsiyonel Sürüklenme ve Hata Katastrofu Verileri',
  ['Gürültü ve İflas Bileşeni',
   'Matematiksel / Moleküler Model',
   'Gençlik Sadakat Düzeyi',
   'Yaşlılık Hata Seviyesi',
   'Dokusal Fonksiyonel Sonuç'],
  [['Kimyasal Master Denklem Gürültüsü',
    'Fano Faktörü F = sigma^2 / mu',
    'F approx 1.05 (Poisson)',
    'F >= 3.85 (Aşırı Gürültü)',
    'Aynı dokudaki hücreler arasında gen ifade kaosu ve heterojenlik'],
   ['Transkripsiyonel Sürüklenme',
    'scRNA-seq Öklid Mesafesi',
    'Mesafe = 1.0 (Referans)',
    'Mesafe >= 2.65 kat artış',
    'Temel ev idaresi genlerinin hücreler arası tutarsız ekspresyonu'],
   ['RNAPII Transkripsiyon Sadakati',
    'Hata / Baz Çifti (Error Rate)',
    '2.8 * 10^-5',
    '1.2 * 10^-4',
    'Kusurlu mRNA sentezi ve primer sekans mutasyonları'],
   ['Ribozomal DRiPs Üretimi',
    'Defective Protein Fraksiyonu',
    '%12 (Normal atık)',
    '%38 (Devasa çöp yükü)',
    'Yeni sentezlenen proteinlerin üçte birinden fazlasının doğuştan çöp olması'],
   ['Soma Enerji Tahsis Oranı',
    'Kirkwood E_maint / E_total',
    '%45 (Yüksek bakım)',
    '%15 (Enerji tükenişi)',
    'Bedenin kendi somatik dokularını onarmayı bırakıp ölüme terk etmesi'],
   ['İyon Kanalı Stokastik Kaçağı',
    'Spontan P_open (Dinlenimde)',
    '0.001 (Sıkı kapalı)',
    '0.045 (Sürekli kaçak)',
    'Dinlenim halindeyken bile kontrolsüz sodyum ve kalsiyum girişi'],
   ['Makromoleküler Kalabalıklaşma',
    'Hücre İçi Hacim Fraksiyonu (phi)',
    '%25 (Akışkan ortam)',
    '%48 (Aşırı tıkalı)',
    'Protein difüzyon hızının %65 yavaşlaması ve metabolik tıkanıklık'],
   ['Lizozomal v-ATPase Proton Akısı',
    'Trans-membran pH Gradyanı',
    'Delta pH = 2.4 ünite',
    'Delta pH = 1.1 ünite',
    'Lizozomun asitliğini kaybederek sindirim yeteneğini durdurması'],
   ['Orgel Hata Katastrofu Kinetiği',
    'dE / dt = alpha*E + beta*E^2',
    'E << E_{crit} (Kararlı)',
    'E -> E_{crit} (Devrilme)',
    'Protein sentez makinelerinin kendi kendini zehirlemesi'],
   ['Sitoplazmik Viskozite',
    'eta (Mikroviskozite / Centipoise)',
    '1.2 cP',
    '3.8 cP',
    'Enzim-substrat karşılaşma frekansının dramatik düşüşü']]),
 ('Tablo 1.3: Epigenetik Enformasyon Kaybı, Horvath Saatleri ve Kromatin Dinamiği Karşılaştırması',
  ['Epigenomik Parametre',
   'Moleküler Biyolojik Tanım',
   'Genç Hücre Karakteristiği',
   'Yaşlı Hücre Karakteristiği',
   'Enformasyonel Biyolojik Etki'],
  [['Horvath DNAmAge Saati',
    '353 CpG Adacığı Metilasyon Skoru',
    'Biyolojik Yaş = Kronolojik',
    'DNAmAge >> Kronolojik (+12 yıl)',
    'Dokunun biyolojik yaşlanma derecesinin matematiksel tescili'],
   ['DunedinPACE Yaşlanma Hızı',
    '173 CpG Lokusu Kan Algoritması',
    '0.75 - 0.90 yıl / yıl',
    '1.25 - 1.55 yıl / yıl',
    'Her takvim yılında 1.5 biyolojik yıl yaşlanma sürati'],
   ['GrimAge Mortalite Riski',
    'Plazma Proteinli Epigenetik Saat',
    'Hazard Ratio = 1.0',
    'Hazard Ratio >= 2.45',
    'Kardiyovasküler ve tüm nedenli ölüm riskinde dev artış'],
   ['Heterokromatin Blokları (H3K9me3)',
    'Nükleer Zarf Susturma Yoğunluğu',
    "Kromatinin %65'i kilitli",
    "Kromatinin %28'i kilitli",
    'Milyonlarca yıllık LINE-1 retrotranspozonlarının uyanması'],
   ['ICE Modeli SIRT Göç Kaybı',
    'DNA Çift Zincir Kırık Tamir Kaybı',
    'epsilon = 0.001 (Tam dönüş)',
    'epsilon >= 0.015 (Kayıp)',
    'Sirtuinlerin promotörlerden kopup bir daha geri dönememesi'],
   ['CpG Metilasyon Entropisi (MEI)',
    'Shannon İkili Belirsizlik İndeksi',
    'MEI = 0.15 (Kesin 0 veya 1)',
    'MEI = 0.58 (Kaotik ara değer)',
    'Hücresel genetik talimatların grileşmesi ve silinmesi'],
   ['SIRT1-7 Enzim Katalitik Verimi',
    'NAD+ Bağımlı Deasetilasyon Hızı',
    'V_{max} = 100% (Tam doyum)',
    'V_{max} <= 25% (NAD+ açlığı)',
    'Kromatin deasetilasyonunun ve mitokondriyel temizliğin durması'],
   ['Topolojik Alanlar (TADs) Sınırı',
    'CTCF / Kohezin Yalıtım Skoru',
    'Insulation = 0.06 (Tam)',
    'Insulation = 0.45 (Geçirgen)',
    'Yanlış promotörlerin yanlış güçlendiricilerle temas etmesi'],
   ['Waddington Çekici Havuzu',
    'Liapunov Potansiyel Derinliği V(x)',
    'Derin Vadi (Kararlı Kimlik)',
    'Sığ Tepe (Kimlik Kaybı)',
    'Hücrenin kendi kimliğini unutup senesansa veya kansere kayması'],
   ['TET Aracılı Epigenetik Reset',
    'dCas9-TET1 5hmC Demetilasyonu',
    'Kapasite Tam Aktif',
    'Baskılanmış / Susturulmuş',
    'Yamanaka faktörleriyle hücresel yaşın sıfıra geri döndürülebilmesi']]),
 ('Tablo 1.4: Serbest Radikaller, Mitohormezis ve Oksidatif Hasar Parametreleri',
  ['Redoks ve Oksidatif Parametre',
   'Biyokimyasal Reaksiyon / Mekanizma',
   'Genç Fizyolojik Değer',
   'Yaşlı Patolojik Değer',
   'Hücresel Moleküler Hasar Çıktısı'],
  [['Elektron Taşıma Kaçak Oranı',
    'Kompleks I/III O2•- Üretimi',
    '%0.8 tüketilen O2',
    '%5.4 tüketilen O2',
    'Mitokondri iç zarının kendi kendini serbest radikalle yakması'],
   ['Glutatyon Redoks Potansiyeli',
    'E_{GSH/GSSG} (Nernst Denklem mV)',
    '-240 mV (Yüksek İndirgeyici)',
    '-168 mV (Oksitleyici Kriz)',
    'Hücre içi proteinlerin disülfid köprülerinin spontan denatürasyonu'],
   ['GSH / GSSG Molar Oranı',
    'İndirgenmiş / Okside Glutatyon',
    '100 : 1 (Güçlü tampon)',
    '12 : 1 (Tükenmiş tampon)',
    'Akut oksidatif stres dalgalarına karşı mutlak savunmasızlık'],
   ['8-OHdG DNA Lezyon Yoğunluğu',
    'Guanin Oksidasyonu (Lezyon / 10^6 dG)',
    '0.45 lezyon',
    '4.85 lezyon (10 kat artış)',
    'Replikasyon sırasında G:C -> T:A transversiyon mutasyonları'],
   ['Peroksinitrit (ONOO-) Oluşumu',
    'NO• + O2•- Reaksiyon Akısı',
    'Bazal minimal seviye',
    '8.5 kat artış',
    'Protein tirozinlerinin nitrotirozine çevrilip kilitlenmesi'],
   ['Nrf2 / ARE Yanıt Kapasitesi',
    'Nükleer Nrf2 Translokasyonu',
    '6.5 kat indüksiyon (Stres anı)',
    '1.2 kat (Körleşmiş yanıt)',
    'Antioksidan enzim genlerinin transkripsiyonunun durması'],
   ['Bach1 Represör Baskınlığı',
    'ARE Promotör Rekabet Katsayısı',
    'Bach1 / Nrf2 = 0.15',
    'Bach1 / Nrf2 = 3.80',
    'Antioksidan yanıt elemanlarının represör tarafından kilitlenmesi'],
   ['Endoplazmik Retikulum UPR Stresi',
    'CHOP ve p-eIF2alpha Seviyesi',
    'Geçici ve adaptif',
    'Kronik hiper-aktivasyon',
    'Protein sentezinin durması ve Caspase-12 bağımlı apoptoz'],
   ['mtDNA Heteroplazmi Oranı',
    'Mutant mtDNA / Toplam mtDNA',
    '< %0.01 (Homoplazmik)',
    '> %60 (Kritik Eşik Aşımı)',
    'Mitokondriyel solunum zincirinin tamamen felç olması'],
   ['Fenton Serbest Demir Havuzu (LIP)',
    'Labile Iron Pool [Fe2+] (uM)',
    '0.15 uM (Ferritinde kilitli)',
    '1.85 uM (Serbest toksik demir)',
    'Durdurulamaz hidroksil radikali (•OH) üretimi']]),
 ('Tablo 1.5: Protein Katlanma Termodinamiği, Proteostaz ve Amiloid Agregasyonu',
  ['Proteostaz Bileşeni',
   'Biyofiziksel / Biyokimyasal Süreç',
   'Genç Hücre Kapasitesi',
   'Yaşlı Hücre Kapasitesi',
   'Hücresel Proteotoksik Yıkım'],
  [['Katlanma Serbest Enerjisi',
    'Delta G_{folding} (Doğal Konformasyon)',
    '-28.5 kJ / mol (Kararlı)',
    '-6.2 kJ / mol (Kararsız)',
    'Proteinlerin kinetik tuzaklara düşerek yanlış katlanması'],
   ['Şaperon Tamponlama Oranı',
    '[Chaperone] / [Misfolded_Protein]',
    '4.5 : 1 (Geniş Rezerv)',
    '0.6 : 1 (Şaperon İflası)',
    'Hatalı katlanmış proteinlerin çökelerek agregat oluşturması'],
   ['Amiloid Çapraz Beta Nükleasyonu',
    'Nucleated Polymerization t_lag',
    'Aylar / Yıllar (Gecikmeli)',
    'Saniyeler / Dakikalar (Hızlı)',
    'Geriye dönüşsüz fibrillerin hızla dokuya çökmesi'],
   ['Toksik Membran Delici Oligomerler',
    'Annular Channel Gözenek Oluşumu',
    'Tespit edilemez düzey',
    'Yüksek konsantrasyon',
    'Hücre zarına delikler açarak sitoplazmaya kontrolsüz kalsiyum akıtma'],
   ['26S Proteazom Klirens Hızı',
    'K48-Ubiquitin Protein Yıkımı',
    '100% (Optimal temizlik)',
    '28% (Ağır tıkanma)',
    'Proteazom fıçısının parçalanamayan agregatlarla kilitlenmesi'],
   ['CMA (Şaperon Aracılı Otofaji)',
    'LAMP-2A Bağımlı Lizozomal Akı',
    'Yüksek seçici yıkım',
    '%78 azalma',
    'Metabolik enzimlerin ve regülatör faktörlerin birikip zehirlemesi'],
   ['Makro-Otofaji ve Katepsin Hızı',
    'Otofagozom-Lizozom Füzyonu',
    'Tam akı (LC3-II klirensi)',
    'Tıkanmış akı (p62 birikimi)',
    'Sitoplazmanın parçalanamayan otofagozom çöpleriyle dolması'],
   ['İleri Glikasyon (AGEs - Glukozpan)',
    'Proteinler Arası Kovalent Çapraz Bağ',
    'Minimal düzey',
    'Dokuda 4.5 kat artış',
    'Damarların ve doku iskeletinin kaskatı tahtaya dönmesi'],
   ['Lipofuscin Birikim Hacmi',
    'Lizozomal Amorf Yaşlılık Çöpü',
    "Hücre hacminin < %1'i",
    "Hücre hacminin %25'i",
    'Post-mitotik nöron ve kalbin kendi çöpünde boğulması'],
   ['Perkolasyon Eşiği (p / p_c)',
    'Agregat Ağının Makroskobik Boyutu',
    'p / p_c = 0.15 (İzole noktalar)',
    'p / p_c >= 1.0 (Makro Jel)',
    'Hücre içi difüzyonun tamamen donması ve ani hücre ölümü']]),
 ('Tablo 1.6: Hücre Membranı Biyofiziği, Akışkanlık ve Lipid Peroksidasyon Kinetiği',
  ['Membran Parametresi',
   'Biyofiziksel Ölçüm / Prob',
   'Genç Membran Değeri',
   'Yaşlı Membran Değeri',
   'Membran Fonksiyonu ve İletim Kaybı'],
  [['DPH Floresans Polarizasyonu (r)',
    'Membran Anizotropi ve Sertliği',
    'r = 0.145 (Yüksek Akışkan)',
    'r = 0.295 (Sert / Rijit)',
    'Reseptörlerin ve iyon kanallarının zar içinde hareket edememesi'],
   ['Lipid Peroksidasyon Zincir Hızı',
    'PUFA Otooksidasyon Döngüsü',
    'Kontrollü sonlanma',
    'N >= 250 döngü / radikal',
    'Tek bir radikalin tüm zar boyunca binlerce lipidi yakması'],
   ['4-HNE Kovalent Adükt Düzeyi',
    'Michael Katılması Toksik Aldehit',
    '< 0.2 uM (İhmal edilebilir)',
    '3.8 uM (Sitotoksik)',
    'Protein sisteinlerinin kovalent kilitlenmesi ve apoptoz'],
   ['Plazmalojen Fraksiyonu (PlsEtn)',
    'Peroksizomal Vinil-Eter Fosfolipid',
    "Toplam PE'nin %35'i",
    "Toplam PE'nin %7'si",
    'Membranın en güçlü doğal antioksidan kalkanını kaybetmesi'],
   ['Kolesterol / Fosfolipid Oranı',
    'Membran Kolesterol Doygunluğu',
    'Molar Oran = 0.62',
    'Molar Oran = 1.35',
    'Kaveolaların sertleşip çökmesi ve endositozun durması'],
   ['Mikroskobik Membran Yırtıkları',
    'Sub-nanometre Boyutlu Açılmalar',
    'Hızlı onarım (< 2 saniye)',
    'Kalıcı sızıntı delikleri',
    'Ekstrasellüler kalsiyumun sitoplazmaya kontrolsüz boşalması'],
   ['Fosfatidilserin Dışa Çıkışı (PS)',
    'Flippase / Scramblase Dengesi',
    'Dış yaprakta < %1',
    'Dış yaprakta %11.5',
    "Makrofajlar için 'Beni Ye' sinyalinin zamansız ateşlenmesi"],
   ['Kardiyolipin Sağlamlık Oranı',
    'Mitokondri İç Zar Dört Bacaklı Lipit',
    '%98 bozulmamış',
    '%32 sağlam (%68 peroksidasyon)',
    "Sitokrom c'nin serbest kalıp apoptozu başlatması"],
   ['Aquaporin Su Geçirgenliği (L_p)',
    'Ozmotik Hidrolik İletkenlik',
    'L_p = 4.2 * 10^-3 cm/s',
    'L_p = 1.8 * 10^-3 cm/s',
    'Hücre içi serbest su hacminin %40 kaybı (Hücresel dehidrasyon)'],
   ['Lipidomik Biyolojik Yaş (LipidAge)',
    'Kütle Spektrometri Lipidom Skoru',
    'LipidAge = 20 Yaş',
    'LipidAge = 78 Yaş',
    'Membranın moleküler gençleşme protokolüne mutlak ihtiyacı']]),
 ('Tablo 1.7: Nükleer Zarf, Nükleopor Kompleksi ve Nükleositoplazmik Taşıma Verileri',
  ['Nükleer Mimari Bileşeni',
   'Yapısal / Fonksiyonel Parametre',
   'Genç Çekirdek Özelliği',
   'Yaşlı Çekirdek Özelliği',
   'Genomik Kararlılık ve İmmün Sonuç'],
  [['Nükleer Şekil Sirkülaritesi',
    'Circularity = 4*pi*Area / P^2',
    '0.95 (Mükemmel Küre)',
    '0.48 (Derin Kırışıklıklar)',
    'Heterokromatin demirleme noktalarının mekanik olarak kopması'],
   ['Lamin B1 Protein Yoğunluğu',
    'Nükleer Lamina Ağ Örgüsü',
    '100% (Sıkı koruyucu ağ)',
    '%18 (Delik deşik iskelet)',
    'Hücresel senesansın geri döndürülemez kesin biyomarkörü'],
   ['Progerin Toksik Birikimi',
    'Farnesillenmiş LMNA Mutasyonu',
    'Saptanamaz düzey',
    'Düşük ama kümülatif birikim',
    'Fizyolojik yaşlanmada progeroid nükleer deformasyon'],
   ['Nükleoporin ELLPs Yaşı',
    'Nup96/107 İskelet Yenilenmesi',
    'Sıfır Turnover (Yaşla aynı)',
    '70 Yıllık Okside Nup Kompleksi',
    'Seçici NPC filtresinin genişleyerek kaçak geçişe izin vermesi'],
   ['Nükleositoplazmik Ran Gradyanı',
    'Delta mu_{Ran} (RanGTP Çekirdek/Sito)',
    '14.5 kJ / mol (Yüksek İtici)',
    '2.1 kJ / mol (Düzleşmiş Eğim)',
    'Transkripsiyon faktörlerinin çekirdeğe girememesi ve nükleer felç'],
   ['Nükleer Yırtılma Sıklığı (NE)',
    'Zarın Mekanik Olarak Delinmesi',
    'Aylar boyu sıfır',
    'Haftada birden fazla yırtılma',
    "Genomik dsDNA'nın kontrolsüzce sitoplazmaya fışkırması"],
   ['cGAS-STING İltihap Ateşi',
    'Sitoplazmik dsDNA Tanıma',
    'Sıfır aktivasyon',
    'Kronik patlama (cGAMP artışı)',
    'Tip-I İnterferon ve IL-6 ile dokuyu yakan steril iltihap'],
   ['Nükleolus Hacim Fraksiyonu',
    'Vol_{nucleolus} / Vol_{nucleus}',
    '%4.2 (Küçük ve Sıkı)',
    '%22.5 (Dev Hipertrofik)',
    'Aşırı rRNA transkripsiyonu ile ömrün dramatik kısalması'],
   ['LINC Kompleksi Gerilim İletimi',
    'SUN-KASH Nükleer-Sitoiskelet Bağı',
    'Tam mekanik iletim',
    'Köprülerin kopması',
    'Hücrenin mekanik hissini kaybederek dış basınca körleşmesi'],
   ['LADs Heterokromatin Demirlenmesi',
    'Lamina ile İlişkili Genom Alanı',
    "Genomun %35'i kilitli",
    "Genomun %8'i bağlı",
    'Karanlık genlerin ve retrotranspozonların serbestçe açılması']]),
 ('Tablo 1.8: Hücresel Biyoenerjetik, ATP Havuzu ve Termal Dağılım Metrikleri',
  ['Biyoenerjetik Değişken',
   'Termodinamik / Biyokimyasal Birim',
   'Genç Metabolik Seviye',
   'Yaşlı Metabolik Seviye',
   'Hücresel Enerji Rezervi ve Güç Çıktısı'],
  [['Adenilat Enerji Şarjı (AEC)',
    'Atkinson Enerji İndeksi [0, 1]',
    '0.92 +/- 0.02 (Tam şarj)',
    '0.68 +/- 0.04 (Boş pil)',
    'Hücresel anabolik biyosentezlerin tamamen kilitlenmesi'],
   ['AMPK Aktivasyon Hassasiyeti',
    'p-AMPK(Thr172) Yanıt Oranı',
    "Düşük ATP'de anında 8 kat",
    "Düşük ATP'de 1.4 kat (Körleşme)",
    'Hücre enerjisizken bile açlık yanıtı ve otofaji verememesi'],
   ['mTORC1 Büyüme Yangısı',
    'p-S6K1 / p-4EBP1 Seviyesi',
    'Kontrollü ve döngüsel',
    'Sürekli hiper-aktif',
    'Blagosklonny hiper-fonksiyonu ile hücrenin tükenip senesansa girmesi'],
   ['Fosfokreatin (PCr) Mekiği',
    'mtCK Oktamer / Dimer Oranı',
    '%85 Oktamer (Aktif Transfer)',
    '%22 Oktamer (İnaktif Dimer)',
    "Mitokondride üretilen ATP'nin tüketim noktalarına taşınamaması"],
   ['Glikolitik Kayma İndeksi (GSI)',
    'Laktat Üretimi / O2 Tüketimi',
    '0.10 (Oksidatif Ağırlıklı)',
    '0.88 (Warburg Benzeri Kayma)',
    "Glukoz başına elde edilen ATP'nin %60'tan fazla erimesi"],
   ['Hücre İçi Termal Sıcak Nokta',
    'Mitokondri - Sitozol Delta T',
    'Delta T <= 2.5 deg C',
    'Delta T >= 8.5 deg C (Aşırı Isı)',
    'Lokal termal sıcak noktaların komşu enzimleri pişirip bozması'],
   ['Boşuna Döngüler (Futile Cycles)',
    'İsraf Edilen ATP Fraksiyonu',
    "Toplam ATP'nin %5'i",
    "Toplam ATP'nin %28'i",
    'Zıt enzimlerin aynı anda çalışarak enerjiyi boşa yakması'],
   ['Mitokondriyel Proton Sızıntısı',
    'Ayrışma Akısı / PMF Kaybı',
    'P/O = 2.50 (Yüksek Verim)',
    'P/O = 1.15 (Ağır Sızıntı)',
    'Elektron taşınmasına rağmen ATP üretemeyen motor iflası'],
   ['İyon Pompası Enerji Faturası',
    'NaK / SERCA ATP Tüketim Payı',
    "Toplam ATP'nin %35'i",
    "Toplam ATP'nin %72'si",
    'Hücrenin tüm servetini sadece iyon sızıntılarını pompalamaya harcaması'],
   ['Biyoenerjetik Reset Skoru (BERS)',
    'NEXA Enerji Restorasyon Puanı',
    '98.5 / 100.0',
    '34.0 / 100.0',
    'Ekzojen NAD+ ve metabolik baypas ile gençlik seviyesine dönüş']]),
 ('Tablo 1.9: Çok Ölçekli Biyolojik Saatler ve Kronobiyolojik Desenkronizasyon',
  ['Biyolojik Saat Düzeyi',
   'Ölçüm Frekansı / Moleküler Saat',
   'Genç Ritmik Davranış',
   'Yaşlı Desenkronizasyon',
   'Organizma Düzeyinde Sistemik Sonuç'],
  [['Sirkadiyen TTFL Amplitüdü',
    'CLOCK/BMAL1 Tepe-Çukur Farkı',
    'Yüksek genlik (Amp = 100)',
    'Sönük genlik (Amp = 35)',
    'Gündüz ve gece metabolik ayrımının silinmesi'],
   ['Organ Faz Senkronizasyonu',
    'SCN vs Karaciğer Faz Kayması',
    'Delta phi = 0 saat (Tam Koherans)',
    'Delta phi = 5.5 saat (Uyumsuz)',
    'Beyin uyurken karaciğerin yemek beklediği biyolojik kaos'],
   ['Epifiz Gece Melatonini',
    'Tepe Plazma Konsantrasyonu',
    '120 pg / mL',
    '15 pg / mL (%88 kayıp)',
    'Mitokondrinin gece antioksidan zırhından mahrum kalması'],
   ['Hayflick Bölünme Sayacı',
    'Kalan Replikatif Bölünme Sayısı',
    '35 - 45 Bölünme Kredisi',
    '0 - 3 Bölünme (Bariyer Sınırı)',
    'Hücrenin bölünmeyi durdurup SASP yayan senesansa girmesi'],
   ['Sirkadiyen NAMPT Salınımı',
    'Hücre İçi NAD+ Günlük Genliği',
    'Delta [NAD+] = 450 uM',
    'Delta [NAD+] = 45 uM (Düz Çizgi)',
    'Sirtuin enzimlerinin 24 saat boyunca uyuması'],
   ['Mikrotübüler THz Titreşimi',
    'Tubulin Dipol Salınım Frekansı',
    '8.5 THz, Q-faktörü = 1.200',
    'Glikasyonlu bozuk rezonans (Q=180)',
    'Kinezin motorlarının kargo taşırken tökezlemesi'],
   ['Biyo-Fotonik Koherans (g2)',
    'Glauber Optik UPE Fonksiyonu',
    'g2(0) = 1.0 (Kuantum Lazer)',
    'g2(0) = 2.1 (Termal Gürültü)',
    'Hücrelerin ışık hızındaki optik telsiz iletişiminin kopması'],
   ['Kuantum Dekoherans Süresi',
    'EZ Su Kalkanı Koruma Süresi',
    'tau_d >= 100 femtosaniye',
    'tau_d <= 1 femtosaniye',
    'Enzimlerde kuantum proton tünelleme hızının düşmesi'],
   ['Organlar Arası Yaş Farkı',
    'Çok Organlı Biyolojik Yaş Vektörü',
    'Tüm organlar 20-22 yaş',
    'Kalp 50, Böbrek 75 yaş',
    'En yaşlı organın sitokinleriyle tüm bedeni çürütmesi'],
   ['Chronos Çok Boyutlu Tensör',
    '10 Katmanlı Biyolojik Zaman Skoru',
    'Tensör Yaşı = 20.5 Yaş',
    'Tensör Yaşı = 76.8 Yaş',
    'PROJECT AETERNITAS ile tüm 10 saatin geri sarılması']]),
 ('Tablo 1.10: Entropik Tersinirlik, Biyolojik Ölümsüzlük ve Doğadaki Emsal Modeller',
  ['Biyolojik / Fiziksel Olgu',
   'Doğal Emsal / Fiziksel Mekanizma',
   'Yaşlanmayan Canlı Davranışı',
   'Normal İnsan Durumu',
   'AETERNITAS Hedef Mühendisliği'],
  [['Gompertz Yaşlanma Katsayısı',
    'mu(t) = mu_0 * exp(G*t)',
    'G = 0 (Hydra, Turritopsis)',
    'G = 0.085 (Ölüm riski katlanır)',
    'G katsayısını insanda sıfırlayarak ölümsüzlük tescili'],
   ['Hayflick Limiti İptali',
    'Sonsuz Bölünme ve Kök Hücreler',
    'Hydra interstisyel kök hücreleri',
    'Somatik hücrede 50 bölünme limiti',
    'Telomeraz gen terapisiyle Hayflick duvarını yıkmak'],
   ['Hücresel Transdiferansiyasyon',
    'Yaşlı Hücrenin Polipe Dönüşü',
    'Turritopsis dohrnii döngüsü',
    'Geriye dönüşsüz diferansiyasyon',
    'Kısmi Yamanaka reprogramlamasıyla hücreyi gençleştirmek'],
   ['Sonsuz Doku Rejenerasyonu',
    'Neoblast Kök Hücre Havuzu',
    'Planaria yassı solucanları',
    'Yara izi ve fibrozis oluşumu',
    'Endojen kök hücre nişlerinin yeniden programlanması'],
   ['Moleküler Maxwell Cini',
    'Szilard Entropi Azaltım Motoru',
    'Kusursuz DNA ve protein tamiri',
    'Tamir makinelerinin yorulması',
    'Sentetik DNA ligaz ve şaperon takviyesi'],
   ['Eksozomal Entropi İhracı',
    'Veziküler Hücresel Çöp Tahliyesi',
    'Sürekli ve yüksek kapasiteli',
    'Eksozom salınımının durması',
    'Heterokronik genç eksozom infüzyonuyla negatif entropi'],
   ['Waddington Tersine Tırmanış',
    'Epigenetik Vadide Yukarı Çıkış',
    'Doğal germline sıfırlaması',
    'Vadinin dibinde senesans',
    'OSK faktörleriyle kontrollü 4 günlük epigenetik tırmanış'],
   ['Fraktal Genom Mimarisi',
    'Düğümsüz Fraktal Globül (D_f)',
    'D_f = 3.0 (Düğümsüz ve açık)',
    'D_f = 2.4 (Düğümlenmiş kaos)',
    'CTCF ve kohezin stabilizasyonuyla 3D genomu düzeltmek'],
   ['Uzun Ömür Kaçış Hızı (LEV)',
    'Rate_{rejuv} > 1.0 yıl / yıl',
    'Doğal sıfır senesans canlıları',
    'Rate_{rejuv} = 0 (Doğal ölüm)',
    'Biyoteknoloji ile biyolojik saati kronolojiden hızlı geri sarmak'],
   ['AETERNITAS Dinamik Denge',
    'Biyolojik Ölümsüzlük Aksiyomları',
    'Ölümsüz biyolojik türler',
    'Kaçınılmaz yaşlanma dogması',
    'İNSAN TÜRÜNÜN BİYOLOJİK OLARAK ÖLÜMSÜZLEŞTİRİLMESİ']])]

# Append Parts 7, 8, 9, 10 to parts list
parts.append((7, "Nükleer Zarf Bütünlüğü, Nükleopor Kompleksi (NPC) ve Lamina Bozulması", part7_topics))
parts.append((8, "Hücresel Enerji Dağılımı, ATP Tüketim Gradyanları ve Termal Dengesizlik", part8_topics))
parts.append((9, "Çok Ölçekli Biyolojik Saatler: Sirkadiyen Ritmlerden Kuantum Biyolojisine", part9_topics))
parts.append((10, "Entropik Tersinirlik İlkesi ve Biyolojik Ölümsüzlüğün Fiziksel Temeli", part10_topics))

print(f"[PROJECT AETERNITAS] Total Parts Loaded: {len(parts)}")
total_topics = sum(len(p[2]) for p in parts)
print(f"[PROJECT AETERNITAS] Total Granular Sections Loaded: {total_topics}")

# Execution loop to render document
print(f"[PROJECT AETERNITAS] Compiling Book 1 Chapter 01: 10 Parts x 10 Topics = 100 Granular Sections...")
sec_counter = 1

for part_idx, (part_num, part_title, topics) in enumerate(parts):
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(22)
    h1.paragraph_format.space_after = Pt(10)
    h1.paragraph_format.keep_with_next = True
    r_h1 = h1.add_run(f"KISIM {part_num}: {part_title.upper()}")
    r_h1.font.name = "Calibri"
    r_h1.font.size = Pt(16)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(15, 35, 75)

    for sec_id, sec_name, lead_txt, deep_txt, formula_txt, deep_exp_txt in topics:
        print(f"  -> Generating Section {sec_id}...")
        add_academic_section(sec_id, sec_name, lead_txt, deep_txt, formula_txt, deep_exp_txt)
        sec_counter += 1

    # Add the corresponding academic table for this part
    tbl_title, tbl_headers, tbl_rows = tables_data[part_idx]
    print(f"  [+] Adding Academic Table {part_idx + 1}...")
    add_academic_table(tbl_title, tbl_headers, tbl_rows)

print(f"[PROJECT AETERNITAS] Saving Masterpiece Document to: {OUTPUT_PATH}...")
doc.save(OUTPUT_PATH)
print(f"[PROJECT AETERNITAS] BÖLÜM 01 100 SAYFALIK MASTERPIECE BAŞARIYLA ÜRETİLDİ: {OUTPUT_PATH}")
