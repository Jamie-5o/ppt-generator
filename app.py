from flask import Flask, request, send_file, jsonify, render_template
from werkzeug.utils import secure_filename
import os
from pptx import Presentation
from pptx.util import Inches

app = Flask(__name__)
UPLOAD_FOLDER = 'uploaded_images'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/upload', methods=['POST'])
def upload_files():
    uploaded_files = request.files.getlist('images')
    for file in uploaded_files:
        filename = secure_filename(file.filename)
        file.save(os.path.join(UPLOAD_FOLDER, filename))
    return jsonify({"status": "uploaded"}), 200

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    variable_images_list = data.get("projects", [])
    output_path = os.path.join(OUTPUT_FOLDER, "result.pptx")
    prs = Presentation()

    for image_group in variable_images_list:
        *img_filenames, title = image_group
        prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = title
        for img_name in img_filenames:
            img_path = os.path.join(UPLOAD_FOLDER, img_name)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(img_path, Inches(1), Inches(1), width=Inches(8))

    prs.save(output_path)
    return send_file(output_path, as_attachment=True)

@app.route('/')
def home():
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)

