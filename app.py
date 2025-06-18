from flask import Flask, request, send_file, render_template
import os
import zipfile
from pptx import Presentation
from pptx.util import Inches
import io
import json

app = Flask(__name__, template_folder='templates')

UPLOAD_FOLDER = 'uploaded_images'
OUTPUT_FOLDER = 'outputs'
TEMPLATE_PATH = os.path.join('templates', 'default_template.pptx')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    # 이미지 저장
    uploaded_files = request.files.getlist('images')
    image_map = {}
    for file in uploaded_files:
        original_name = file.filename
        filename = os.path.basename(original_name)
        path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(path)
        image_map[original_name] = path

    # 프로젝트 데이터 파싱
    project_data = json.loads(request.form.get('project_data', '[]'))

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        for project in project_data:
            *img_filenames, title = project
            prs = Presentation(TEMPLATE_PATH)

            # 제목 슬라이드 수정
            for shape in prs.slides[0].shapes:
                if shape.has_text_frame and "광고 상품 소개서" in shape.text:
                    shape.text = f"{title} 광고 상품 소개서"
                    break

            for img_name in img_filenames:
                img_path = image_map.get(img_name)
                if not img_path:
                    continue
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_picture(
                    img_path, Inches(0), Inches(0),
                    width=prs.slide_width,
                    height=prs.slide_height
                )

            output_pptx = f"{title}.pptx"
            output_path = os.path.join(OUTPUT_FOLDER, output_pptx)
            prs.save(output_path)
            zipf.write(output_path, arcname=output_pptx)

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
