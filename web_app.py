import streamlit as st
from pptx import Presentation
import io

# Sayfa Tasarımı
st.set_page_config(page_title="Sunum Otomasyonu", page_icon="📐")
st.title("📐 Expo Art Design Sunum Otomasyonu")
st.write("Fuar bilgilerini içeren şablonunuzu ve stant render'larını yükleyin, sunumunuz saniyeler içinde hazır olsun.")

# Dosya Yükleme Alanları
sablon_dosya = st.file_uploader("Şablon Dosyasını (sablon.pptx) Yükleyin", type=["pptx"])
yuklenen_resimler = st.file_uploader("Stant Görsellerini Yükleyin", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

if st.button("Sunumu Oluştur", type="primary"):
    if sablon_dosya is not None and len(yuklenen_resimler) > 0:
        with st.spinner("Sunum hazırlanıyor, lütfen bekleyin..."):
            prs = Presentation(sablon_dosya)
            eski_slaytlar = list(prs.slides._sldIdLst)
            
            # Resimleri ismine göre sırala
            yuklenen_resimler = sorted(yuklenen_resimler, key=lambda x: x.name)

            # İlk slaytı ekle (Kapak / Bilgi)
            ilk_duzen = prs.slide_layouts[0]
            slayt_1 = prs.slides.add_slide(ilk_duzen)
            for shape in slayt_1.placeholders:
                if shape.placeholder_format.type == 18:
                    shape.insert_picture(yuklenen_resimler[0])
                    break

            # Orta Slaytları Ekle
            orta_duzen = prs.slide_layouts[1]
            for resim in yuklenen_resimler[1:]:
                slayt = prs.slides.add_slide(orta_duzen)
                for shape in slayt.placeholders:
                    if shape.placeholder_format.type == 18:
                        shape.insert_picture(resim)
                        break

            # Kapanış Slaytı Ekle
            bitis_duzeni = prs.slide_layouts[2] 
            prs.slides.add_slide(bitis_duzeni)

            # Baştaki boş slaytı temizle
            for slide_id in eski_slaytlar:
                prs.slides._sldIdLst.remove(slide_id)

            # Dosyayı hafızaya kaydet (indirmek için)
            pptx_stream = io.BytesIO()
            prs.save(pptx_stream)
            pptx_stream.seek(0)
            
            st.success("Tebrikler! Sunum başarıyla oluşturuldu.")
            
            # İndirme Butonu
            st.download_button(
                label="📥 Hazır Sunumu İndir",
                data=pptx_stream,
                file_name="ASIS_Innotrans_Sunum.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    else:
        st.warning("Lütfen önce şablonu ve resimleri yüklediğinizden emin olun.")