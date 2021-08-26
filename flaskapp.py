
from flask import Flask, render_template, request, send_file, make_response,Response
import pandas as pd
from styleframe import StyleFrame
from pathlib import Path

app = Flask(__name__)

UPLOAD_FOLDER = str(Path.home() / "Downloads")
ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
var_list = []

@app.route("/",methods=['GET', 'POST'])
def hello():
    if request.method == 'POST':
        act.clear()
        actresult.clear()
        var_list.clear()
        Rfalse.clear()
        Rtrue.clear()
        persent.clear()

        #file = request.files['file']

        f = request.files['file']
        f2 = request.files['file2']
        f3 = request.form['file3']
        f.save(f.filename)
        f2.save(f2.filename)

        print(f3)

        assignment(f,f2,f3)

        print('done')
        name1 = request.files['file'].filename
        var_list.append(name1)





        return render_template('result.html',persent = persent,Rtrue = Rtrue,Rfalse = Rfalse)

    #return render_template('home.html')
    return render_template('index.html')


#def upload(f,f2):
   # master = pd.read_excel(f)
   # student =pd.read_excel(f2)
    #result = master['MasterVerdict'].str.lower() == student['Verdict'].str.lower()


    #pdList = [student['Case ID'],student['Verdict'], master['MasterVerdict'],result,student['Observations'], master['MasterObservation'] ]  # List of your dataframes
    #new_df = pd.concat(pdList, axis=1)
    #name1 = request.files['file'].filename
    #new_df.to_excel(name1)
    #return render_template('download.html')

act = []
actresult = []
persent = []
Rtrue = []
Rfalse = []
def assignment(f,f2,f3):

    master = pd.read_excel(f)
    student = pd.read_excel(f2)
    df = pd.concat([master, student], axis=1)
    master['Actualverdict'] = master['Actualverdict'].str.lower()
    student['Verdict'] = student['Verdict'].str.lower()
    for i in df.index:
        result = master['Actualverdict'][i] == student['Verdict'][i]
        if result == False:
            corobr = master['Actualobservation'][i]

        else:
            corobr = ''

        act.append(corobr)
        actresult.append(result)
       # print(corobr)

    master['three'] = act
    master['four'] = actresult

    finaldflist = [master['Case id'], master['Actualverdict'], student['Verdict'], master['four'], master['three']]

    new_df = pd.concat(finaldflist, axis=1)
    #new_df.to_excel('assign.xlsx', encoding="utf-8")
    true = new_df['four'].value_counts()[1]
    false = new_df['four'].value_counts()[0]
    total = true + false
    quotient = true / total

    percentage = quotient * 100
    persent.append(percentage)
    Rtrue.append(true)
    Rfalse.append(false)


    fit = "{}{}.xlsx".format(f3,persent)
    #path =  "C:\Users\sande\{}".format(fit)
    out_path = "C:\\Users\\sande\\{}".format(fit)
    ot =  r'C:\Users\sande\assignments\{}'.format(fit)



    print(fit)

    s = new_df.style.applymap(color_negative_red)
    writer = pd.ExcelWriter(ot, engine='xlsxwriter')
    s.to_excel(writer, sheet_name='Sheet1')
    #s.to_excel(fit, encoding="utf-8")
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Add some cell formats.
    format1 = workbook.add_format({'num_format': '#,##0.00'})
    format2 = workbook.add_format({'num_format': '0%'})
    worksheet.set_column('F:F', 22, format1)
    writer.save()


   # writer_orig.save()
    #new_df.to_excel('gbhn.xlsx', encoding="utf-8")

def color_negative_red(val):


    color = 'red' if val == False else 'black'
    return 'color: %s' % color

@app.route("/download",methods=['GET', 'POST'])
def download_file():
    print('globel')
    print(var_list[0])
    q = var_list[0]


    return send_file(q,as_attachment=True)

@app.route("/result",methods=['GET', 'POST'])
def result():
    return render_template('result.html')

@app.route("/index",methods=['GET', 'POST'])
def index():

    return render_template('index.html')




if __name__=="__main__":
    app.run(debug=True)
