from flask import Flask,render_template,send_from_directory,request,send_file
from group import process_excel

import pandas as pd



app=Flask(__name__)
@app.route("/downloads/<path:filename>")
def allow_download(filename):
    return send_file("downloads/"+filename,as_attachment=True)

@app.route('/', methods=["GET","POST"])


def open_index():
    if request.method=="POST":
        if 'file' not in request.files:
            return render_template('index.html',message="no file selected")
        excelfile=request.files['file']
        if excelfile.filename=="":

            return render_template('index.html',message="chuitya select file")
        if excelfile.filename.rsplit(".")[-1]!="xlsx":
            return render_template('index.html',message="please select excel file only")
        if request.form["coloumnnumber"]=="":
            return render_template('index.html',message="please put column number")
        try:
            input_file_pate="uploads/"+excelfile.filename
            output_file_pate='downloads/consolidate.xlsx'
            sumcoloumn=int(request.form['coloumnnumber'])
            excelfile.save(input_file_pate)
            process_excel(input_file_pate,output_file_pate,sumcoloumn)
            return render_template("index.html",message="file process succesfully",dowload_file=output_file_pate)

        except Exception as e:
            return render_template("index.html",message=e)    


    return render_template("index.html")




if __name__ == '__main__':
    app.run(debug=True,host='0.0.0.0', port=5000)
