import streamlit as st
import streamlit.components.v1 as components
import os
import io
import base64
import json
import subprocess
import shutil
import platform
from pdf2image import convert_from_path

st.markdown("## Version-4 — 20260523 162000 更新 - クラウド対応のままZIPを廃止。JavaScript経由で複数画像をローカルへ直接一括ダウンロードする方式に変更")

st.title("PPTX to Rotated Images")
st.write("アップロードされたPPTXの各スライドを画像化し、左に90度回転してローカルへ直接ダウンロードします。")

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
                # 1. アップロードファイルを一時保存
                with open(temp_pptx_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # 2. OS判定（Windowsローカルテスト時は soffice、クラウドLinux環境では libreoffice）
                libreoffice_cmd = "soffice" if platform.system() == "Windows" else "libreoffice"
                
                # 3. PPTX → PDF 変換（LibreOffice Headless）
                subprocess.run(
                    [libreoffice_cmd, "--headless", "--nologo", "--nofirststartwizard",
                     "--convert-to", "pdf", temp_pptx_path, "--outdir", temp_dir],
                    check=True,
                    capture_output=True
                )
                
                temp_pdf_path = os.path.join(temp_dir, f"{pptx_filename}.pdf")
                
                # 4. PDF → 画像化（pdf2image）
                images = convert_from_path(temp_pdf_path)
                
                # 5. 各画像を90度回転 → base64エンコードしてリスト化
                image_data_list = []
                for i, img in enumerate(images):
                    slide_num = i + 1
                    rotated_img = img.rotate(90, expand=True)  # 左90度回転（見切れ防止）
                    
                    img_byte_arr = io.BytesIO()
                    rotated_img.save(img_byte_arr, format="JPEG")
                    img_byte_arr.seek(0)
                    
                    b64_str = base64.b64encode(img_byte_arr.read()).decode("utf-8")
                    image_data_list.append({
                        "name": f"{pptx_filename}-{slide_num}.jpg",
                        "data": b64_str
                    })
                
                st.success(f"変換完了: {len(images)} 枚の画像を生成しました。下のボタンを押すと一括ダウンロードが始まります。")
                
                # 6. JavaScript一括ダウンロードボタンを描画
                images_json = json.dumps(image_data_list)
                download_html = f"""
                <div style="font-family: sans-serif;">
                    <button id="dl-btn" style="
                        padding: 12px 24px; 
                        font-size: 16px; 
                        background-color: #FF4B4B; 
                        color: white; 
                        border: none; 
                        border-radius: 6px; 
                        cursor: pointer;
                        font-weight: bold;
                    ">📥 全画像を一括ダウンロード</button>
                    <p id="status" style="margin-top: 12px; font-size: 14px;"></p>
                </div>
                <script>
                const images = {images_json};
                const btn = document.getElementById("dl-btn");
                const status = document.getElementById("status");
                
                btn.addEventListener("click", async () => {{
                    btn.disabled = true;
                    btn.style.backgroundColor = "#888";
                    for (let i = 0; i < images.length; i++) {{
                        const img = images[i];
                        const a = document.createElement("a");
                        a.href = "data:image/jpeg;base64," + img.data;
                        a.download = img.name;
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        status.textContent = `ダウンロード中: ${{i + 1}} / ${{images.length}} 枚`;
                        // ブラウザの連続DL制御のため少しウェイトを入れる
                        await new Promise(r => setTimeout(r, 350));
                    }}
                    status.textContent = `✅ ${{images.length}} 枚すべてのダウンロードが完了しました。`;
                }});
                </script>
                """
                components.html(download_html, height=140)

            except Exception as e:
                st.error(f"変換エラーが発生しました: {e}")
            finally:
                # 一時ディレクトリのクリーンアップ
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
