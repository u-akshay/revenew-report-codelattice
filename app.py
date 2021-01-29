from flask import Flask,render_template,request,url_for, send_file

from werkzeug.utils import secure_filename
import os
import time
import final


app = Flask(__name__)


################################################################


@app.route('/', methods=['GET' , 'POST'])
def login():
    
    if request.method=="POST":
        a = request.files['fileField']
        b = request.files['fileField1']
        a.save(os.path.join("static", "pd.xlsx"))
        b.save(os.path.join("static", "et.xlsx"))

        final.generate()
        p = "Revenew.pdf"
        return send_file(p,as_attachment=True)
        # return render_template('index.html', message="upload now")
    
        

    return render_template('index.html')



if __name__ == '__main__':
    app.run(debug=True)
