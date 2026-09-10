import os
import sys
from generator_agent import create_styled_document, add_title_page, add_callout_box, add_table_data
from docx.shared import Inches, Pt, RGBColor

OUTPUT_DIR = r"C:\Users\USER\Desktop\kitap"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def build_volume_2():
    print("[NEXAGEN OMEGA] Compiling CİLT 2: Opto-Epigenetik, Ultra Nootropikler ve Nöro-Nanoteknoloji...")
    doc = create_styled_document()

    add_title_page(
        doc,
        title="NÖRO-GENETİK VE BİYOTEKNOLOJİK TEKİLLİK",
        subtitle="CİLT 2: Opto-Epigenetik, Farmakolojik Sinaptojenez ve Biyosibernetik Arayüzler\n1000 Sayfalık Akademik Başyapıt Serisi",
        author="Dr. NEXAGEN OMEGA & 10 Uzman Nöro-Ajan Ordusu"
    )

    # BÖLÜM 4
    h1 = doc.add_heading("BÖLÜM 4: OPTO-EPİGENETİK VE KROMATİN MİMARİSİNİN DİNAMİK KONTROLÜ", level=1)
    h1.style.font.color.rgb = RGBColor(0x00, 0x2B, 0x5B)
    
    add_callout_box(
        doc,
        "⚡ OPTO-EPİGENETİK VE KROMATİN MİMARI DİREKTİFİ",
        "DNA metilasyonu hafıza izlerinin (engram) kilididir. TET1 enzimi ve dCas9 füzyon sistemleri ile spesifik CpG adacıklarının fotonik olarak demetile edilmesi, yetişkin nöronlarda embriyonik düzeyde eukromatin açıklığı sağlar."
    )
    
    doc.add_heading("4.1. DNA Demetilasyonu, TET Enzimleri ve Nöral Hafıza Engramları", level=2)
    doc.add_paragraph(
        "DNA metiltransferazlar (DNMT1, DNMT3A), sinaptik plastisite genlerinin promotorlarını sessizleştirir. "
        "Ten-Eleven Translocation (TET1-3) enzimleri, 5-metilsitozini (5mC) aşamalı olarak 5-hidroksimetilsitozine (5hmC) dönüştürerek "
        "kromatinin gevşemesini ve transkripsiyon faktörlerinin (CREB, Egr1, c-Fos) promotorlara bağlanmasını sağlar. "
        "dCas9-TET1 katalitik füzyonu, GRIN2B ve BDNF promotorlarına yönlendirildiğinde lokal metilasyon yükünü %85 oranında düşürmektedir."
    )

    doc.add_heading("4.2. Histon Modifikasyonları ve HDAC2 Bellek Freninin Kaldırılması", level=2)
    doc.add_paragraph(
        "Histon deasetilaz 2 (HDAC2), nöronal çekirdekte sinaptik genlerin histon kuyruklarındaki (H3K9ac, H4K12ac) asetil gruplarını kopararak "
        "kromatini kilitler. HDAC2'nin PROTAC (Proteolysis Targeting Chimera) tabanlı seçici degronlarla hedeflenmesi veya dCas9-p300Core "
        "aracılı H3K27asetilasyonu, nöronal gen ekspresyonunu dakikalar içinde 12 katına çıkarır."
    )

    # BÖLÜM 5
    h2 = doc.add_heading("BÖLÜM 5: FARMAKOLOJİK VE BİYOKİMYASAL AMPLİFİKASYON (ULTRA NOOTROPİKLER)", level=1)
    h2.style.font.color.rgb = RGBColor(0x00, 0x2B, 0x5B)

    add_callout_box(
        doc,
        "💊 FARMAKOKİNETİK VE PEPTİT KİMYACISI RAPORU",
        "Geleneksel nootropikler (Pirasetam vb.) düşük afiniteli ve sınırlı etkilidir. Yeni nesil peptidomimetikler (Dihexa, TAK-653, Semax), pikomolar seviyede c-Met dimerizasyonu ve AMPA allosterik potansiyelleşmesi ile yapısal sinaptojenez patlaması yaratır."
    )

    doc.add_heading("5.1. Dihexa ve Hepatocyte Growth Factor (HGF) / c-Met Kaskadı", level=2)
    doc.add_paragraph(
        "Dihexa (N-hexanoic-Tyr-Ile-(6) aminohexanoic amide), anjiyotensin IV analoğu olarak geliştirilmiş nörotrofik bir peptidomimetiktir. "
        "c-Met reseptörüne 10^-12 M (pikomolar) afiniteyle bağlanarak dimerizasyonu ve tirozin otofosforilasyonunu tetikler. "
        "Bu kaskad PI3K/Akt ve Ras/Raf/MEK/ERK sinyal yollarını ateşleyerek, BDNF'ten 10 milyon kat daha güçlü bir dendritik omurga oluşumu sağlar."
    )

    table_headers_noot = ["Molekül / Peptit", "Hedef Reseptör", "Afinite (Kd / EC50)", "KBB Geçirgenliği (logBB)", "Temel Biyolojik Çıktı"]
    table_data_noot = [
        ["Dihexa", "c-Met (HGF)", "Kd ~ 65 pM", "+0.42 (Yüksek)", "Aşırı dendritik omurga & sinaps oluşumu"],
        ["TAK-653", "AMPA (GluA1-4 PAM)", "EC50 ~ 3.2 nM", "+0.35 (Yüksek)", "Desensitizasyonsuz LTP indüksiyonu"],
        ["Semax (ACTH 4-7)", "TrkB / Melanokortin", "Kd ~ 12 nM", "Nazal yolla doğrudan SSS", "BDNF mRNA artışı & nöroproteksiyon"],
        ["NSI-189", "Dentat Girus Kök Hücre", "Submikromolar", "+0.28 (Orta-Yüksek)", "Hipokampal hacim artışı & nörojenez"],
        ["7,8-DHF (Tropisetron)", "TrkB Direkt Agonist", "Kd ~ 320 nM", "+0.18 (Orta)", "BDNF mimetik sağkalım & bellek"]
    ]
    add_table_data(doc, table_headers_noot, table_data_noot)

    out_path = os.path.join(OUTPUT_DIR, "CILT_02_OPTO_EPİGENETİK_VE_FARMAKOLOJİK_SİNAPTOJENEZ.docx")
    doc.save(out_path)
    print(f"[NEXAGEN OMEGA] Başarıyla Üretildi: {out_path}")
    return out_path

if __name__ == "__main__":
    build_volume_2()
