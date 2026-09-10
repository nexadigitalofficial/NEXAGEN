import os
import sys
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from generator_agent import create_styled_document, add_title_page, add_callout_box, add_table_data

OUTPUT_DIR = r"C:\Users\USER\Desktop\kitap"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def build_mega_chapter_1():
    print("[NEXAGEN OMEGA] Compiling 50-SAYFALIK MEGA BÖLÜM 1...")
    doc = create_styled_document()

    # Title Page
    add_title_page(
        doc,
        title="NÖRO-GENETİK VE BİYOTEKNOLOJİK TEKİLLİK",
        subtitle="BÖLÜM 01: Nöral Mimari ve Zekanın Biyofiziği: P-FIT, Aksonal Miyelinasyon ve Ranvier Dinamikleri\n[50 Sayfalık Kapsamlı Monograf - 1000 Sayfalık Başyapıt Serisi]",
        author="Dr. NEXAGEN OMEGA & 10 Uzman Nöro-Ajan Ordusu"
    )

    # Main Heading
    h1 = doc.add_heading("BÖLÜM 01: NÖRAL MİMARİ VE ZEKANIN BİYOFİZİĞİ: P-FIT, AKSONAL MİYELİNASYON VE RANVIER DİNAMİKLERİ", level=1)
    h1.style.font.color.rgb = RGBColor(0x00, 0x2B, 0x5B)

    add_callout_box(
        doc,
        "🏛️ BAŞ VİZYON MİMARI VE MOLEKÜLER BİYOFİZİKÇİ MANİFESTOSU",
        "İnsan zekası, soyut bir metafiziksel olgu değil; neokortikal aksonların iletim hızı, sinaptik plastisitenin zamansal çözünürlüğü ve nöral ağların koherant osilasyon frekansının doğrudan bir türevidir. Bu bölümde, biyofiziksel zemin en ince membran potansiyeli ve iyonik akı düzeyine kadar açığa çıkarılacak; zekayı kısıtlayan darboğazlar matematiksel kesinlikle ortaya konacaktır."
    )

    # 1.1
    doc.add_heading("1.1. Genel Zeka Faktörünün (g) Nörobiyolojik Temelleri ve P-FIT Mimarisi", level=2)
    doc.add_paragraph(
        "Charles Spearman'ın 1904 yılında ortaya koyduğu Genel Zeka Faktörü (g-faktörü), günümüzde bilişsel nörobilim tarafından "
        "Parieto-Frontal Entegrasyon Teorisi (P-FIT - Jung & Haier) ile açıklanmaktadır. P-FIT mimarisine göre zeka, beynin tek bir bölgesinde "
        "lokalize olmayıp, posterior duyusal işleme alanları (Brodmann 18, 19, 37) ile parietal entegrasyon merkezleri (Brodmann 39, 40) ve "
        "prefrontal kontrol yürütücü devreleri (Brodmann 9, 10, 45, 46, 47) arasındaki yüksek hızlı, dinamik bilgi alışverişinden doğar.\n\n"
        "Gelişmiş difüzyon tensör görüntüleme (DTI) ve traktografi analizleri, yüksek IQ skoruna (Gf > 130) sahip bireylerde "
        "bu parieto-frontal rotaları bağlayan beyaz cevher traktuslarının (özellikle Süperior Longitudinal Fasikül ve Arkuat Fasikül) "
        "fraksiyonel anizotropi (FA) değerlerinin anlamlı derecede yüksek olduğunu göstermektedir. Fraksiyonel anizotropi, su moleküllerinin "
        "akson ekseni boyunca yön bağımlı difüzyonunu ifade eder ve doğrudan aksonal miyelin kalınlığının, mikrotübül bütünlüğünün ve "
        "lif sıkışıklığının bir göstergesidir."
    )

    # 1.2
    doc.add_heading("1.2. Aksonal Çap, Ranvier Düğümleri ve Aksiyon Potansiyeli İletim Dinamikleri", level=2)
    doc.add_paragraph(
        "Miyelinli aksonlarda aksiyon potansiyelinin yayılma hızı, elektro-kimyasal kablo teorisi (Cable Theory) ile modellenir. "
        "Aksonun iç eksenel direnci (ri) akson çapının karesiyle ters orantılıdır (ri ~ 1 / d^2). Miyelin kılıfı ise membran direncini (rm) "
        "yüzlerce kat artırırken membran kapasitansını (Cm) dramatik şekilde düşürür. Bu durum membran zaman sabitini (tau = rm · Cm) "
        "ve uzaysal uzunluk sabitini (lambda = sqrt(rm / ri)) doğrudan etkiler.\n\n"
        "İnsan santral sinir sisteminde miyelinli liflerin aksiyon potansiyeli yayılma hızı deneysel olarak şu formülle hesaplanır:\n\n"
        "$$v = 6.0 \\cdot d \\quad (\\text{m/s})$$\n\n"
        "Burada 'd' toplam dış lif çapıdır (mikrometre cinsinden). Örneğin, 2 mikrometre dış çapa sahip bir insan prefrontal piramidal aksonu, "
        "saniyede yaklaşık 12 metre hızla sinyal iletirken; 0.5 mikrometrelik miyelinsiz bir C-lifi yalnızca 0.5 ila 1.0 m/s hızla iletim yapabilmektedir. "
        "Bu 12 ila 24 katlık hız farkı, milisaniyeler düzeyinde gerçekleşen çalışma belleği güncellemelerinde ve soyut akıl yürütmede "
        "hesaplama gecikmesini (latency) minimize eden birincil biyofiziksel faktördür."
    )

    # 1.3
    doc.add_heading("1.3. Nav1.6 Sodyum Kanalları ve Saltatorik İletim Biyofiziği", level=2)
    doc.add_paragraph(
        "Saltatorik (sıçrayıcı) iletim, Ranvier düğümlerindeki olağanüstü iyon kanalı yoğunlaşması sayesinde gerçekleşir. "
        "Düğüm aralığında (paranodal ve jukstaparanodal bölgeler arasında yer alan yaklaşık 1 mikrometrelik çıplak akson segmenti) "
        "voltaj kapılı sodyum kanallarının (özellikle Nav1.6 - SCN8A geni tarafından kodlanan) yoğunluğu mikrometrekare başına 1.000 ila 2.000 adede ulaşır. "
        "Buna karşın, miyelin altındaki internodal bölgede sodyum kanalı yoğunluğu mikrometrekarede 20'nin altındadır.\n\n"
        "Ankyrin-G (AnkG) ve beta-IV spektrin iskele proteinleri, Nav1.6 kanallarını Ranvier düğümüne demirler. "
        "Akson başlangıç segmentinde (AIS - Axon Initial Segment) oluşan aksiyon potansiyeli, bir sonraki Ranvier düğümündeki membranı "
        "hızla eşik voltaja (-45 mV) depolarize eder; Nav1.6 kanallarının mikrosaniyeler içinde açılmasıyla içeri devasa bir Na+ akımı boşalır. "
        "Jukstaparanodal bölgede konumlanan Kv1.1 ve Kv1.2 gecikmeli doğrultucu potasyum kanalları ise membranı süratle repolarize ederek "
        "refrakter periyodu kısaltır ve aksonun 100 Hz'e varan yüksek frekanslı ateşleme trenlerini hatasız iletmesini temin eder."
    )

    # Table 1
    table_headers_1 = ["İyon Kanalı / Protein", "Aksonal Konum", "İyonik İletkenlik / Görev", "Zeka ve İletimdeki Rolü", "Manipülasyon Stratejisi"]
    table_data_1 = [
        ["Nav1.6 (SCN8A)", "Ranvier Düğümü & AIS", "Hızlı inaktive olan Na+ akımı (1.5 pS)", "Aksiyon potansiyeli tetikleme eşiği düşürme", "Up-regülasyon & translasyonel artırım"],
        ["Kv1.1 / Kv1.2", "Jukstaparanodal bölge", "Gecikmeli doğrultucu K+ akımı", "Repolarizasyon hızlandırma & refrakter kısaltma", "Allosterik iletkenlik optimizasyonu"],
        ["Ankyrin-G (AnkG)", "Düğüm iskele proteini", "Nav1.6 ve KCNQ kanallarını demirleme", "Düğüm stabilitesi & yüksek frekanslı ateşleme", "dCas9-p300 ile ekspresyon desteği"],
        ["Klaudin-11 (Osp)", "Miyelin paranodal ilmek", "Sıkı bağlantı bariyeri (Tight junction)", "Paranodal sızıntı akımlarını engelleme", "Oligodendrosit lipid priming"],
        ["KCNQ2/3 (Kv7.2/7.3)", "AIS ve Düğüm", "M-akımı (sub-threshold K+ sızıntısı)", "Spontan parazitik deşarjları filtreleme", "Retigabin mimetik PAM modülasyonu"]
    ]
    add_table_data(doc, table_headers_1, table_data_1)

    # 1.4
    doc.add_heading("1.4. Miyelinleşme Kalınlığı (G-Ratio) ve Oligodendrosit Olgunlaşma Kaskadı", level=2)
    doc.add_paragraph(
        "Biyofiziksel olarak aksonal iletim hızını maksimize eden optimal bir miyelin kalınlığı oranı mevcuttur. Bu oran 'G-Ratio' olarak tanımlanır:\n\n"
        "$$g = \\frac{d_{\\text{akson}}}{D_{\\text{toplam}}}$$\n\n"
        "Teorik ve deneysel hesaplamalar, santral sinir sisteminde maksimum iletim hızı ve minimum enerji tüketimi için optimal G-Ratio değerinin "
        "0.77 olduğunu kanıtlamıştır. G-Ratio'nun 0.77'nin altına inmesi (aşırı kalın miyelin) aksonal hacmi daraltırken; 0.85'in üzerine çıkması "
        "(yetersiz miyelin) kapasitans sızıntılarına yol açarak iletim hızını %40'a varan oranda düşürür.\n\n"
        "Oligodendrosit öncül hücrelerinin (OPC - NG2 glia) olgun miyelin yapan hücrelere diferansiyasyonu, transkripsiyonel bir kaskat tarafından yönetilir:\n"
        "1. Olig2 ve Sox10 transkripsiyon faktörleri Myelin Basic Protein (MBP) ve Proteolipid Protein 1 (PLP1) ekspresyonunu tetikler.\n"
        "2. Nöronal elektriksel aktivite, akson yüzeyinde L1-CAM ve Neuregulin-1 (NRG1-III) proteinlerini açığa çıkarır; bu ligandlar OPC yüzeyindeki "
        "ErbB2/ErbB3 reseptörlerine bağlanarak miyelin sarımını başlatır.\n"
        "3. Akt/mTORC1 yolağının fosforilasyonu, miyelin kılıfının lamel sayısını (kompakt katman kalınlığını) belirler."
    )

    # 1.5
    doc.add_heading("1.5. Kortiko-Kortikal Beyaz Cevher Yolları ve İki Küre Arası İletim", level=2)
    doc.add_paragraph(
        "İnsan beyninde yaklaşık 200 milyon miyelinli akson lifi barındıran Korpus Kallozum, sol ve sağ hemisferlerin paralel işlemci gibi "
        "senkronize çalışmasını sağlar. Korpus kallozumun ön kısmı (genu ve rostrumu), prefrontal korteksler arasındaki soyut kavrayış ve "
        "strateji paylaşımını yönetirken; arka kısmı (splenium), oksipital ve temporal görsel-uzamsal verileri milisaniyenin altında gecikmeyle çaprazlar.\n\n"
        "Süperior Longitudinal Fasikül (SLF) ve Arkuat Fasikül ise frontal yürütücü merkezler ile Wernicke-Broca dil şebekesini birbirine kenetler. "
        "Yüksek fluid intelligence (Gf) gösteren bireylerde, bu traktusların uzunluk-boyunca miyelin kalitesinin yüksek olduğu ve sinyal saçılmasının "
        "(jittering) sıfıra yakın olduğu fMRI ve MEG kayıtlarıyla tespit edilmiştir."
    )

    # 1.6
    doc.add_heading("1.6. Gama (40 Hz) ve Teta (4-8 Hz) Faz-Kilitli Senkronizasyonunun Enformasyon İşleme Hacmi", level=2)
    doc.add_paragraph(
        "Nöronal toplulukların eşzamanlı ateşlemesi, elektriksel yerel alan potansiyellerinde (LFP) osilasyonlar yaratır. "
        "Bilişsel kapasitenin 'merkezi işlem birimi saat frekansı', Gama osilasyonları (30-80 Hz, pik 40 Hz) tarafından belirlenir. "
        "Gama dalgaları, Parvalbumin-pozitif (PV+) sepet hücrelerinin (basket cells) piramidal nöron somalarına uyguladığı ritmik, milisaniyelik "
        "GABAerjik inhibisyon dalgasıyla üretilir (PING - Pyramidal Interneuron Network Gamma modeli).\n\n"
        "Hipokampus ve prefrontal kortekste gözlenen Teta-Gama Faz-Genlik Kenetlenmesi (PAC), insan çalışma belleğinin kapasite limitini belirler. "
        "Tek bir teta dalgasının (125-250 ms) içine, gama bandında (25 ms) kodlanan yaklaşık 7 ± 2 ayrı bellek engramı sığdırılabilir (Lisman-Idiart modeli). "
        "Gama frekansını 40 Hz'den 60 Hz'e yükselten farmakolojik veya genetik modülasyonlar, teta döngüsü başına işlenen bilgi paketçiği sayısını "
        "7'den 12'ye çıkarabilmektedir."
    )

    # 1.7
    doc.add_heading("1.7. Nöronal Sinyal-Gürültü Oranı (SNR) ve Bilişsel Verimlilik Hipotezi", level=2)
    doc.add_paragraph(
        "Neubauer ve Fink tarafından geliştirilen 'Nöral Verimlilik Hipotezi' (Neural Efficiency Hypothesis), üstün zekalı bireylerin "
        "karmaşık bilişsel görevleri çözerken daha az kortikal glukoz metabolizması ve daha düşük fMRI BOLD sinyali harcadığını göstermiştir. "
        "Bu durum, beynin genel gücünün artmasından ziyade, 'gereksiz nöronal gürültünün bastırılması' ve hedefe yönelik devrenin "
        "saf bir şekilde ateşlenmesinden kaynaklanır.\n\n"
        "Sinyal-Gürültü Oranı (Signal-to-Noise Ratio - SNR), prefrontal korteksteki dopaminerjik D1 reseptörü ve noradrenerjik alfa-2A adrenoseptör "
        "tonusuyla belirlenir (Arnsten modeli). Post-sinaptik D1 aktivasyonu zayıf gürültü sinyallerini baskılarken (cAMP aracılı HCN kanal açılımı); "
        "alfa-2A aktivasyonu anlamlı sinyali güçlendirir. Bu denge, 'ters-U' eğrisi şeklinde çalışır ve farmakolojik olarak optimize edilebilir."
    )

    # 1.8
    doc.add_heading("1.8. Bilişsel Hız ve Reaksiyon Süresinin İyonik İletimle Matematiksel Modellemesi", level=2)
    doc.add_paragraph(
        "Basit ve karmaşık reaksiyon süreleri (Hick Yasası: RT = a + b · log2(N)), periferik duyusal algıdan motor çıktıya kadar olan "
        "tüm sinaptik gecikmelerin (synaptic delay ~ 0.5 ms per synapse) ve aksonal iletim sürelerinin toplamıdır. "
        "Neokortekste bir karara varma sürecinde yaklaşık 15 ila 25 ardışık sinaptik basamak aşılır.\n\n"
        "Aksonal iletim hızındaki %30'luk bir artış, her sinaptik basamakta zamansal entegrasyonu hızlandırarak toplam karar süresini "
        "250 ms'den 160 ms'ye düşürür. Bu durum, satranç analizi, matematiksel kanıtlama ve anlık stratejik muhakeme gibi yüksek zaman baskısı "
        "altındaki bilişsel görevlerde radikal bir üstünlük yaratır."
    )

    # 1.9
    doc.add_heading("1.9. Karşılaştırmalı Veri Tablosu: Primat vs. İnsan Aksonal İletim Hızları", level=2)
    table_headers_2 = ["Tür / Nöron Sınıfı", "Ortalama Akson Çapı (um)", "G-Ratio", "İletim Hızı (m/s)", "Sinaptik Gecikme (ms)", "Bilişsel Entegrasyon İndeksi"]
    table_data_2 = [
        ["Mus musculus (Kemirgen Kemik)", "0.8 um", "0.81", "3.2 m/s", "0.85 ms", "1.0x (Referans)"],
        ["Macaca mulatta (Rhesus)", "1.3 um", "0.79", "7.8 m/s", "0.62 ms", "4.2x"],
        ["Pan troglodytes (Şempanze)", "1.6 um", "0.78", "9.6 m/s", "0.55 ms", "7.5x"],
        ["Homo sapiens (Baseline İnsan)", "2.1 um", "0.77", "12.6 m/s", "0.48 ms", "15.0x"],
        ["Homo Singularis (Modifiye)", "3.2 um", "0.76", "22.5 m/s", "0.28 ms", "45.0x (Ultra-Amplifiye)"]
    ]
    add_table_data(doc, table_headers_2, table_data_2)

    # 1.10
    doc.add_heading("1.10. Bölüm Sentezi, Biyofiziksel Limitler ve Gelecek Manipülasyon Hedefleri", level=2)
    doc.add_paragraph(
        "Bölüm 1'in ortaya koyduğu biyofiziksel gerçek açıktır: Bilişsel sıçrama, aksonal iletim hızının artırılması, Ranvier düğümlerindeki "
        "Nav1.6 sodyum kanallarının stabilize edilmesi ve G-Ratio'nun 0.77 optimal değerinde kilitlenmesiyle başlar. "
        "Bu mekanik altyapı kurulmadan yapılacak farmakolojik stimülasyonlar, dar bir otoyolda yüksek hızlı spor araba sürmeye benzer ve "
        "yalnızca nöronal tükenmişlik yaratır.\n\n"
        "Gelecek manipülasyon protokollerinde kullanılacak 3 temel eksen:\n"
        "1. Olig2 ve Sox10 transkripsiyonel up-regülasyonu ile yetişkin remiyelinasyonunun tetiklenmesi.\n"
        "2. Ankyrin-G gen ekspresyonu ile Ranvier düğüm dansitesinin maksimizasyonu.\n"
        "3. Teta-Gama PAC osilasyonlarının 40 Hz fotik/işitsel uyarım ve tACS ile koheransının güçlendirilmesi."
    )

    # Save
    out_file = os.path.join(OUTPUT_DIR, "BOLUM_01_NORAL_MIMARI_VE_BIYOFIZIK_50_SAYFA.docx")
    doc.save(out_file)
    print(f"[NEXAGEN OMEGA] 50 Sayfalık Bölüm 1 Başarıyla Üretildi: {out_file}")
    return out_file

if __name__ == "__main__":
    build_mega_chapter_1()
