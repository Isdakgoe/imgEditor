import streamlit as st
import os
import io
import zipfile
import subprocess
import shutil
import platform
from pdf2image import convert_from_path

st.markdown("## Version-2 — 20260523 110429 更新 - WEB版Streamlit対応。LibreOfficeを用いた透かしなし変換・左90度回転機能を追加。")

st.title("PPTX to Rotated Images")
st.write("アップロードされたPPTXの各スライドを画像化し、左に90度回転して保存します。")

uploaded_file = st.file_uploader("PPTXファイルを選択", type=["pptx"])

if uploaded_file is not None:
    filename_with_ext = uploaded_file.name
    pptx_filename = os.path.splitext(filename_with_ext)[0]
    
    if st.button("変換開始"):
        with st.spinner("スライドを変換しています（数十秒かかる場合があります）..."):
            temp_dir = "temp_processing"
            os.makedirs(temp_dir, exist_ok=True)
            temp_pptx_path = os.path.join(temp_dir, filename_with_ext)
            
            try:
                # 1. アップロードされたファイルを一時保存
                with open(temp_pptx_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # 2. OSに応じたLibreOfficeコマンドの判定
                # Windowsローカルテスト時は 'soffice'、Streamlit Cloud(Linux)環境では 'libreoffice'
                libreoffice_cmd = "soffice" if platform.system() == "Windows" else "libreoffice"
                
                # 3. PPTX -> PDF 変換 (LibreOffice Headless)
                subprocess.run(
                    [libreoffice_cmd, "--headless", "--nologo", "--nofirststartwizard", "--convert-to", "pdf", temp_pptx_path, "--outdir", temp_dir],
                    check=True,
                    capture_output=True
                )
                
                temp_pdf_path = os.path.join(temp_dir, f"{pptx_filename}.pdf")
                
                # 4. PDF -> 画像化 (pdf2image)
                images = convert_from_path(temp_pdf_path)
                
                # 5. 回転とZIP圧縮
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for i, img in enumerate(images):
                        slide_num = i + 1
                        
                        # 左に90度回転（見切れ防止）
                        rotated_img = img.rotate(90, expand=True)
                        
                        # メモリ上でJPEG化
                        img_byte_arr = io.BytesIO()
                        rotated_img.save(img_byte_arr, format="JPEG")
                        img_byte_arr.seek(0)
                        
                        # フォルダ構成を維持してZIPに書き込み
                        zip_path = f"{pptx_filename}/{pptx_filename}-{slide_num}.jpg"
                        zf.writestr(zip_path, img_byte_arr.read())
                        
                zip_buffer.seek(0)
                st.success(f"完了: {len(images)} 枚の画像を変換しました。")
                
                st.download_button(
                    label="保存フォルダをZIPでダウンロード",
                    data=zip_buffer,
                    file_name=f"{pptx_filename}.zip",
                    mime="application/zip"
                )

            except Exception as e:
                st.error(f"変換エラーが発生しました: {e}")
            finally:
                # 一時ディレクトリのクリーンアップ
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
