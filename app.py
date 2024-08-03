from flask import Flask, render_template, request, flash, redirect, url_for, send_file
from flask_wtf.csrf import CSRFProtect
from forms import UploadForm
from converter import converter_pdf_para_excel
import os


from flask_frozen import Freezer

app = Flask(__name__)
app.config['SECRET_KEY'] = '04b49f7c740b5763f9a4b79ad07623e0'
csrf = CSRFProtect(app)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/upload", methods=['GET', 'POST'])
def upload():
    form = UploadForm()
    if form.validate_on_submit():
        file = form.file.data
        pdf_path = os.path.join('uploads', file.filename)
        file.save(pdf_path)

        # Definir o caminho do arquivo Excel de saída
        excel_path = os.path.splitext(pdf_path)[0] + '.xlsx'

        # Converter o PDF para Excel
        converter_pdf_para_excel(pdf_path, excel_path)

        # Verificar se o arquivo Excel foi criado
        if os.path.exists(excel_path):
            # Nome do arquivo Excel convertido
            excel_filename = os.path.basename(excel_path)

            flash('Arquivo convertido com sucesso para Excel', 'success')
            return render_template('upload.html', form=form, excel_filename=excel_filename)

    return render_template('upload.html', form=form)

@app.route("/download/<excel_filename>", methods=['GET'])
def download(excel_filename):
    excel_path = os.path.join('uploads', excel_filename)
    if os.path.exists(excel_path):
        return send_file(excel_path, as_attachment=True)
    else:
        flash('Arquivo nao encontrado', 'error')
        return redirect(url_for('upload'))

if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True, host='0.0.0.0')
