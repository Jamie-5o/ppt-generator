from flask import Flask, request, send_file, render_template
import os
import zipfile
from pptx import Presentation
from pptx.util import Inches
import io
import json
import comtypes.client

app = Flask(__name__, template_folder='templates')

UPLOAD_FOLDER = 'uploaded_images'
OUTPUT_FOLDER = 'outputs'
TEMPLATE_PATH = os.path.join('templates', 'default_template.pptx')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def convert_ppt_to_pdf(ppt_path, pdf_path):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
    deck.SaveAs(pdf_path, 32)  # 32 = PDF
    deck.Close()
    powerpoint.Quit()

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    uploaded_files = request.files.getlist('images')
    image_map = {}
    for file in uploaded_files:
        original_name = file.filename
        filename = os.path.basename(original_name)
        path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(path)
        image_map[original_name] = path

    project_data = json.loads(request.form.get('project_data', '[]'))

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        for project in project_data:
            *img_filenames, title = project
            prs = Presentation(TEMPLATE_PATH)

            # 제목 수정 (폰트 유지)
            for shape in prs.slides[0].shapes:
                if shape.has_text_frame and "광고 상품 소개서" in shape.text:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "광고 상품 소개서" in run.text:
                                run.text = f"{title} 광고 상품 소개서"

            # 이미지 추가
            for img_name in img_filenames:
                img_path = image_map.get(img_name)
                if not img_path:
                    continue
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_picture(
                    img_path, Inches(0), Inches(0),
                    width=prs.slide_width, height=prs.slide_height
                )

            # 저장 경로
            output_pptx = f"{title}.pptx"
            output_pdf = f"{title}.pdf"
            output_pptx_path = os.path.join(OUTPUT_FOLDER, output_pptx)
            output_pdf_path = os.path.join(OUTPUT_FOLDER, output_pdf)

            # 저장 및 변환
            prs.save(output_pptx_path)
            convert_ppt_to_pdf(os.path.abspath(output_pptx_path), os.path.abspath(output_pdf_path))

            # zip에 추가
            zipf.write(output_pptx_path, arcname=output_pptx)
            zipf.write(output_pdf_path, arcname=output_pdf)

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name='광고소개서_모음.zip',
        mimetype='application/zip'
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
