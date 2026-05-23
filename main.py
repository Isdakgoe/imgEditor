import streamlit as st
import os
import io
import zipfile
from PIL import Image
import aspose.slides as slides

st.markdown("## Version-1 — 20260523 100153 更新 - PPTXのアップロード・画像化・左90度回転保存機能を追加。")

st.title("PPTX to Rotated Images")
st.write("アップロードされたPPTXの各スライドを画像化し、左に90度回転して保存します。")

uploaded_file = st.file_uploader("PPTXファイルを選択", type=["pptx"])

if uploaded_file is not None:
    filename_with_ext = uploaded_file.name
    pptx_filename = os.path.splitext(filename_with_ext)[0]
    
    output_dir = pptx_filename
    os.makedirs(output_dir, exist_ok=True)
    
    temp_pptx_path = f"temp_{filename_with_ext}"
    with open(temp_pptx_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    if st.button("変換開始"):
        with st.spinner("スライドをレンダリングし、回転処理を行っています..."):
            try:
                pres = slides.Presentation(temp_pptx_path)
                total_slides = len(pres.slides)
                progress_bar = st.progress(0)
                
                for i in range(total_slides):
                    slide = pres.slides[i]
                    
                    # スライドを画像（サムネイル）として取得 (スケール1.0)
                    bmp = slide.get_thumbnail(1.0, 1.0)
                    
                    # メモリ上でPIL画像に変換
                    img_stream = io.BytesIO()
                    bmp.save(img_stream, slides.export.ImageFormat.JPEG)
                    img_stream.seek(0)
                    pil_img = Image.open(img_stream)
                    
                    # 左に90度回転（expand=Trueで見切れを防止）
                    rotated_img = pil_img.rotate(90, expand=True)
                    
                    # 指定された命名規則で保存 (1-indexed)
                    slide_num = i + 1
                    save_path = os.path.join(output_dir, f"{pptx_filename}-{slide_num}.jpg")
                    rotated_img.save(save_path, "JPEG")
                    
                    progress_bar.progress(slide_num / total_slides)
                    
                st.success(f"完了: `{output_dir}` フォルダ内に {total_slides} 枚の画像を保存しました。")
                
                # デプロイ環境でも取得できるようZIPダウンロードボタンを提供
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for i in range(total_slides):
                        slide_num = i + 1
                        img_path = os.path.join(output_dir, f"{pptx_filename}-{slide_num}.jpg")
                        zf.write(img_path, arcname=f"{pptx_filename}/{pptx_filename}-{slide_num}.jpg")
                
                zip_buffer.seek(0)
                st.download_button(
                    label="保存フォルダをZIPでダウンロード",
                    data=zip_buffer,
                    file_name=f"{pptx_filename}.zip",
                    mime="application/zip"
                )

            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
            finally:
                if os.path.exists(temp_pptx_path):
                    os.remove(temp_pptx_path)
