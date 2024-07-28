from flask_wtf import FlaskForm
from wtforms import FileField, SubmitField
from wtforms.validators import DataRequired

class UploadForm(FlaskForm):
    file = FileField('Upload PDF', validators=[DataRequired()])
    submit = SubmitField('Convert to Excel')
