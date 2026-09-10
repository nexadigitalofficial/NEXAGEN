import os
import sys
from generator_agent import create_styled_document, add_title_page, add_callout_box, add_table_data
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
OUTPUT_DIR = r"C:\Users\USER\Desktop\kitap"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def build_volume_1():
    print("[NEXAGEN OMEGA] Compiling CİLT 1: Moleküler Nörobiyoloji, Metabolizma ve İnsanüstü Genetik Mühendislik...")
    doc = create_styled_document()

    # Title Page
    add_title_page(
        doc,
        title="NÖRO-GENETİK VE BİYOTEKNOLOJİK TEKİLLİK",
        subtitle="CİLT 1: Bilişsel Biyofizik, Mitokondriyal Enerji ve İleri Düzey Genomik Düzenleme\n1000 Sayfalık Akademik Başyapıt Serisi",
        author="Dr. NEXAGEN OMEGA & 10 Uzman Nöro-Ajan Ordusu"
    )

    # =========================================================================
    # BÖLÜM 1
    # =========================================================================
    h1 = doc.add_heading("BÖLÜM 1: BİLİŞSEL MİMARİNİN BİYOFİZİKSEL VE NÖROBİYOLOJİK TEMELLERİ", level=1)
    h1.style.font.color.rgb = RGBColor(0x00, 0x2B, 0x5B)

    add_callout_box(
        doc,
        "🏛️ BAŞ VİZYON MİMARI VE MOLEKÜLER NÖROBİYOLOG BİLDİRİSİ",
        "İnsan zekasını amplifiye etme hedefi, beynin neokortikal sinaps yoğunluğu, aksonal miyelin iletim hızı ve kuantum düzeyindeki mikrotübüler koheransın birleşik bir fonksiyonudur. Tek bir hedefi değiştirmek yetersizdir; sinaps öncesi vezikül ekzositozundan post-sinaptik dansite (PSD-95) iskelelerine kadar tüm kaskad senkronize edilmelidir."
    )

    doc.add_heading("1.1. Genel Zeka Faktörünün (g-Factor) ve Akışkan Zekanın (Gf) Nöral Korelatları", level=2)
    p = doc.add_paragraph(
        "Bilişsel kapasitenin evrensel ölçütü olan Genel Zeka Faktörü (g), nöro-anatomik düzeyde Parieto-Frontal Entegrasyon Teorisi (P-FIT) "
        "ile açıklanır. Neokorteksin Brodmann 9, 10, 46 (Dorsolateral Prefrontal Korteks - DLPFC) ve Brodmann 39, 40 (İnferior Parietal Lobül) "
        "alanları arasındaki bilgi akışı, aksonal miyelin kılıfının kalınlığı ve Ranvier düğümlerinin uzaysal dağılımıyla doğrudan orantılıdır. "
        "Aksiyon potansiyeli iletim hızı (v), lif çapı (d) ve miyelin kalınlığı ile doğrusal bir ilişki sergiler: v = 6 · d (m/s). "
        "Gama bandı osilasyonları (30-80 Hz) ile hipokampal teta ritminin (4-8 Hz) faz-genlik kenetlenmesi (Phase-Amplitude Coupling - PAC), "
        "çalışma belleğindeki bilginin aynı anda kaç bayt olarak işlenebileceğini (WMC - Working Memory Capacity) belirleyen temel elektrofizyolojik saattir."
    )

    doc.add_heading("1.2. İnsan Neokorteksinin Hücresel Ayrışması: İnsan Nöronlarının Benzersizliği", level=2)
    doc.add_paragraph(
        "Homo sapiens neokorteksi, diğer primatlarla karşılaştırıldığında radikal yapısal modifikasyonlara sahiptir. "
        "Von Economo nöronları (VEN - iğsi nöronlar) ve Kandelaber (Chandelier) internöronları, piramidal nöronların akson başlangıç segmentine (AIS) "
        "doğrudan GABAerjik frenleme uygulayarak ağ düzeyinde faz gürültüsünü süzer. İnsan piramidal nöronları, kemirgen nöronlarına kıyasla "
        "üç kat daha uzun apikal dendritlere ve binlerce kat daha yüksek sinaptik entegrasyon hacmine sahiptir. "
        "İnsan-özgü gen duplikasyonu olan SRGAP2C proteini, atasal SRGAP2A'yı dimerize ederek inhibe eder; bu durum dendritik omurga (spine) "
        "olgunlaşmasını geciktirir (neoteni) ve nihai sinaptik omurga yoğunluğunu iki katına çıkarır."
    )

    # Table 1: Neokortikal Hücre ve İletim Parametreleri
    table_headers_1 = ["Nöron Tipi / Parametre", "Lokalizasyon", "İletim / Kinetik Değeri", "Bilişsel Fonksiyon", "Manipülasyon Potansiyeli"]
    table_data_1 = [
        ["Piramidal Nöron (Katman V)", "DLPFC / Motor Korteks", "10-120 m/s iletim hızı", "Yüksek biliş ve soyut kavrayış", "Dendritik omurga amplifikasyonu"],
        ["Von Economo Nöronu (VEN)", "Ön Singulat & İnsula", "Hızlı uzun-mesafe projeksiyon", "Sosyal sezgi, hızlı karar verme", "Metabolik koruma & TrkB desteği"],
        ["Kandelaber İnternöron", "Neokorteks Katman II/III", "Akson Başlangıç Segmenti blokajı", "Ağ desensitize gürültü filtrasyonu", "GABA-A alfa2 spesifik agonizm"],
        ["Ranvier Düğümü", "Kortikokortikal aksonlar", "Saltatorik iletim (Nav1.6 yoğun)", "Sinyal gecikmesini <2 ms tutma", "Miyelinasyon kalınlaştırma (Olig2)"],
        ["Astrositik Son Ayak", "Nörovasküler ünite", "AQP4 su & Ca2+ dalgaları", "Glifatik drenaj & laktat mekiği", "EAAT2 glutamat klerensi artırımı"]
    ]
    add_table_data(doc, table_headers_1, table_data_1)

    doc.add_heading("1.3. Sinaps Biyofiziği, PSD-95 İskelesi ve LTP Mekanizması", level=2)
    doc.add_paragraph(
        "Sinaptik plastisitenin temeli olan Uzun Süreli Potansiyelleşme (Long-Term Potentiation - LTP), moleküler düzeyde şu kinetik basamaklarla işler:\n"
        "1. Presinaptik vezikül ekzositozu: Aksiyon potansiyeli terminale ulaştığında P/Q- ve N-tipi voltaj kapılı kalsiyum kanalları (VGCC) açılır; "
        "kalsiyum sinaptotagmin-1'e bağlanarak SNARE kompleksini (Sintaksin-1, SNAP-25, VAMP2/Sinaptobrevin) tetikler.\n"
        "2. NMDAR Açılımı: Dinlenim potansiyelinde (-70 mV) NMDA reseptör kanalı magnezyum iyonu (Mg2+) tarafından bloke edilir. "
        "AMPA reseptörlerinden (GluA1/GluA2) sodyum akışı membranı -30 mV'ye depolarize ettiğinde Mg2+ elektrostatik olarak kanaldan fırlar.\n"
        "3. CaMKII Otofoforilasyonu: Kanaldan içeri akan devasa kalsiyum dalgası Kalmodulin (CaM) ile birleşir ve CaMKII alfa enzimini aktive eder. "
        "CaMKII, Thr286 bölgesinde otofosforile olarak 'otonom moleküler bellek' haline gelir; post-sinaptik dansitedeki PSD-95 ve Stargazin proteinlerini "
        "fosforilleyerek hücre içi veziküllerdeki AMPA reseptörlerini membrana kilitler."
    )

    # =========================================================================
    # BÖLÜM 2
    # =========================================================================
    h2 = doc.add_heading("BÖLÜM 2: BEYİN METABOLİZMASI, MİTOKONDRİ VE BİYOFENOMENOLOJİK ENERJİ VERİMLİLİĞİ", level=1)
    h2.style.font.color.rgb = RGBColor(0x00, 0x2B, 0x5B)

    add_callout_box(
        doc,
        "⚡ MİTOKONDRİYAL BİYOFİZİKÇİ RAPORU",
        "İnsan beyni vücut ağırlığının yalnızca %2'sini oluşturmasına karşın glukozun %25'ini ve oksijenin %20'sini tüketir. Yüksek IQ fenotipinde nöronlar daha fazla enerji tüketmez; aksine 'Nöral Verimlilik Hipotezi' uyarınca aksiyon potansiyeli başına düşen ATP maliyetini en aza indirerek sinyal-gürültü oranını maksimize eder."
    )

    doc.add_heading("2.1. Nöronal Mitokondriyal Dinamikler: Füzyon, Fisyon ve ATP Biyogenezi", level=2)
    doc.add_paragraph(
        "Nöronal somadan akson terminallerine ve dendritik omurgalara uzanan mikrotübüler kinezor motorları, mitokondrileri yüksek enerji gerektiren "
        "sinaps bölgelerine taşır. Mitokondriyal ağın sağlığı Mitofusin 1/2 (MFN1, MFN2) ve OPA1 proteinlerinin yönettiği füzyon ile DRP1 (Dynamin-Related Protein 1) "
        "tarafından yürütülen fisyon dengesine bağlıdır. "
        "Peroksizom proliferatör ile aktive olan reseptör gama koaktivatör 1-alfa (PGC-1a), NRF-1 ve NRF-2 transkripsiyon faktörlerini aktive ederek "
        "Mitokondriyal Transkripsiyon Faktörü A'yı (TFAM) uyarır. TFAM, mitokondriyal DNA (mtDNA) replikasyonunu ve Kompleks I-V solunum zinciri "
        "proteinlerinin sentezini patlatarak nöronal ATP havuzunu %40-60 oranında artırır."
    )

    # Table 2: Mitokondriyal ve Metabolik Amplifikasyon Parametreleri
    table_headers_2 = ["Hedef Yolak / Enzim", "Endojen Fonksiyonu", "Amplifikasyon Yöntemi", "Bilişsel / Hücresel Çıktı", "Klinik Güvenlik Sınırı"]
    table_data_2 = [
        ["PGC-1alpha / TFAM", "Mitokondriyal biyogenez ana şalteri", "AICAR, Epikateşin, dCas9-p300", "%55 daha yüksek ATP kapasitesi", "Aşırı oksidatif stres önlenmeli"],
        ["SIRT1 / SIRT3", "NAD+ bağımlı deasetilaz", "NMN (1000 mg) + Resveratrol", "FOXO3 uyarımı, SOD2 anti-oksidasyon", "Sirtuin hiperaktivasyonu güvenli"],
        ["ANLSH (MCT2 taşıyıcı)", "Astrositten nörona laktat mekiği", "BHB keton esterleri, L-Laktat", "Glukozsuz ortamda bile nöronal ateşleme", "Ketoasidoz pH kontrolü"],
        ["AQP4 Akuaporin-4", "Glifatik drenaj su kanalı", "Derin SWS uykusu, Yan yatış pozisyonu", "Beta-amiloid ve Tau toksin klerensi", "Kronik uyku kısıtlamasında bozulur"],
        ["Kompleks IV (Sitokrom c)", "Elektron transport zinciri terminali", "810-1064 nm Fotobiyomodülasyon (PBM)", "Mitokondriyal membran potansiyeli artışı", "Termal hasar eşiği <42°C"]
    ]
    add_table_data(doc, table_headers_2, table_data_2)

    # =========================================================================
    # BÖLÜM 3
    # =========================================================================
    h3 = doc.add_heading("BÖLÜM 3: GENOMİK MANİPÜLASYON HEDEFLERİ VE İNSAN ÜSTÜ GENETİK MÜHENDİSLİK", level=1)
    h3.style.font.color.rgb = RGBColor(0x00, 0x2B, 0x5B)

    add_callout_box(
        doc,
        "🧬 SENTETİK BİYOLOJİ VE CRISPR BAŞ MÜHENDİSİ DİREKTİFİ",
        "İnsan zekasını belirleyen tek bir 'zeka geni' yoktur; ancak nöronal plastisiteyi, dendritik hacmi ve miyelinasyon hızını kısıtlayan evrimsel darboğaz genleri mevcuttur. Prime Editing (PE5max) ve AAV.CAP-B10 vektörleri ile yetişkin nöronlarda GRIN2B ve SRGAP2C hedeflerini güvenle yeniden yazmak mümkündür."
    )

    doc.add_heading("3.1. Zeka ile İlişkili Temel Gen Ailesi ve Moleküler Fonksiyonları", level=2)
    doc.add_paragraph(
        "Nörogenetik araştırmalar, insan bilişsel kapasitesini doğrudan modüle eden 6 kritik gen grubunu izole etmiştir:\n\n"
        "1. GRIN2B (GluN2B / NR2B): NMDA reseptörünün düzenleyici alt birimidir. Genç beyinlerde baskın olan NR2B, yaşlanmayla yerini NR2A'ya bırakır. "
        "NR2B içeren NMDAR'lar, hücreye 4 kat daha uzun süre kalsiyum akışı sağlayarak LTP eşiğini düşürür. Joe Tsien'in Doogie farelerinde kanıtlandığı üzere, "
        "yetişkin ön beyinde NR2B aşırı ekspresyonu öğrenme hızını 5 katına çıkarmaktadır.\n\n"
        "2. SRGAP2C: İnsansıların yaklaşık 2.4 milyon yıl önce Australopithecus'tan Homo cinsine geçişi sırasında duplike olan bu gen, "
        "dendritik diken boyunlarını uzatır ve sinapsların biyofiziksel elektrik izolasyonunu mükemmelleştirir.\n\n"
        "3. ARHGAP11B: Neokortikal bazal radial glia hücrelerinin simetrik bölünmesini artırarak kortikal kıvrımlaşmayı (girifikasyon) katlar.\n\n"
        "4. BDNF (Brain-Derived Neurotrophic Factor): Val66Met (rs6265) polimorfizmi, pro-BDNF'in aktivite-bağımlı salınımını bozar. "
        "Prime editing ile 66. kodondaki Metin yerine Valin (Val/Val) yazılması, hipokampal hacmi ve bellek stabilitesini doğrudan yükseltir."
    )

    doc.add_heading("3.2. Prime Editing (PE5max) ve Kapsit Mühendisliği ile Hedefe Teslimat", level=2)
    doc.add_paragraph(
        "Geleneksel CRISPR-Cas9'un çift zincir kırıkları (DSB) ve rastgele indel oluşturma riski, bölünmeyen yetişkin nöronlarda kabul edilemez. "
        "Bu nedenle geliştirilen Prime Editing 5 (PE5max), şu moleküler bileşenlerden oluşur:\n"
        "- Cas9 Nickase (H840A): Yalnızca hedef zinciri keser, karşı zinciri sağlam bırakır.\n"
        "- M-MLV Ters Transkriptaz (Engineered Reverse Transcriptase): Primer Bağlanma Bölgesi (PBS) ile açılan DNA zincirine tutunur ve "
        "Ters Transkripsiyon Şablonundaki (RTT) yeni genetik diziyi doğrudan DNA'ya kopyalar.\n"
        "- Sentetik pegRNA: 13 nt PBS uzunluğu ve 16 nt RTT uzunluğu ile optimize edilen pegRNA, GRIN2B promotor mutasyonlarını %70+ saflıkla düzeltir.\n"
        "- Vektör Tasarımı: AAV.CAP-B10 kapsidi, endotelyal LY6A reseptörüne bağlanarak Kan-Beyin Bariyerini intravenöz enjeksiyonla aşar ve "
        "yalnızca CamKIIa promotörü taşıyan nöronlarda eksprese olur."
    )

    # Table 3: Genetik Mühendislik ve Prime Editing Matrisi
    table_headers_3 = ["Hedef Gen Lokusu", "Moleküler Müdahale Türü", "pegRNA PBS / RTT Dizaynı", "Vektör & Tropizm", "Beklenen Bilişsel Çıktı"]
    table_data_3 = [
        ["GRIN2B (NR2B)", "Promotor demetilasyonu & cDNA regülasyonu", "PBS: 13 nt / RTT: 15 nt", "AAV.CAP-B10 (CamKIIa)", "LTP eşiğinde %65 düşüş, süper-bellek"],
        ["BDNF (Val66Met)", "rs6265 A->G Adenin Base Editing", "sgRNA: Exon IX spesifik", "LNP-mRNA (ApoE kaplı)", "Aktivite-bağımlı BDNF salınım restorasyonu"],
        ["SRGAP2C", "Transgenik regülasyon (Katman II/III)", "cDNA eklemesi (PE3 nicking)", "AAV.PHP.eB (Synapsin-1)", "Dendritik sinaps yoğunluğunda 2x artış"],
        ["ARHGAP11B", "Subventriküler kök hücre transkripsiyonu", "dCas9-p300 epigenetik açılım", "AAV-Retro (DLPFC projeksiyon)", "Glioblastomasız neokortikal plastisite"],
        ["KLOTHO (KL-VS)", "F952V varyantı ekzojen sentezi", "Prime Editing PE5max", "AAV8 (Sistemik Karaciğer)", "Periferik nöroproteksiyon ve GluN2B stabilizasyonu"]
    ]
    add_table_data(doc, table_headers_3, table_data_3)

    # Save Document
    output_docx_path = os.path.join(OUTPUT_DIR, "CILT_01_NOKTA_BİYOFİZİK_VE_GENETİK_MÜHENDİSLİK.docx")
    doc.save(output_docx_path)
    print(f"[NEXAGEN OMEGA] Başarıyla Üretildi: {output_docx_path}")
    return output_docx_path

if __name__ == "__main__":
    build_volume_1()
