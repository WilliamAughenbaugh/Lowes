from flask import Flask, render_template, request, redirect
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from werkzeug.utils import secure_filename
import panda as pd


app = Flask(__name__)
app.secret_key = "yeet"
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///test.sqlite3'
app.config['UPLOAD_FOLDER']
app.config['MAX_CONTENT_PATH'] = 1000000
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# excel global sheet vars
file = ""

class Todo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customerName = db.Column(db.String, primary_key=True)
    customerPhoneNumber = db.Column(db.String, primary_key=True)
    customerProductETA = db.Column(db.String, primary_key=True)
    date_created = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return '<Task %r>' % self.id


@app.route('/upload')
def upload_file():
    return render_template('upload.html')


@app.route('/uploader', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        f.save(secure_filename(f.filename))
        file = f.filename
        return 'file uploaded successfully'


@app.route('/', methods=['POST', 'GET'])
def index():
    if request.method == 'POST':
        task_content = request.form['content']
        new_task = Todo(content=task_content)
        try:
            db.session.add(new_task)
            db.session.commit()
            return redirect('/')
        except:
            return 'There was an issue adding your task'

    else:
        tasks = Todo.query.order_by(Todo.date_created).all()
        return render_template('index.html', tasks=tasks)

if __name__ == "__main__":
    db.create_all()
    app.run(debug=True)


def ExcelInfoGrabber():
    # Global Vars for excel parsing.
    todayDate = datetime.date.today()
    twoWeeksOuts = todayDate + datetime.timedelta(days=14)
    yearVar = todayDate + datetime.timedelta(days=28835)



    workbookGUI = pd.read_excel(file, sheet_name="Sheet1")
    workbook2GUI = pd.read_excel(file, sheet_name="Sheet1")

    customerName = workbookGUI["Customer Name"]
    customerPhoneNum = workbookGUI["Customer Name"]
    userName = workbookGUI["Assigned To"]

    sizeOfArray = len(userName)
    print(sizeOfArray)

    pd.to_datetime(workbook2GUI["ETA"], format="%b %d %Y")
    eta = workbook2GUI["ETA"]

    customerNameArr = [""]
    customerPhoneNumArr = [""]
    etaArr = [""]
    etaToCompare: datetime

    for x in range(sizeOfArray):
        customerNameArr.append(customerName.iloc[0 + x])
        customerPhoneNumArr.append(customerPhoneNum.iloc[0+x])
        for x in range(sizeOfArray):
            etaArr.append(eta.iloc[0 + x])
        for x in range(sizeOfArray):
            etaStr = etaArr[1 + x]
            if etaArr[1 + x] == "ERROR":
                etaArr[1 + x] = yearVar
