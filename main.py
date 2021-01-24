from flask import   Flask, \
                    render_template, \
                    request,flash, \
                    redirect, \
                    url_for, \
                    send_file, \
                    after_this_request
from flask_caching import Cache
from threading import Thread, Event
import pandas as pd
import mammoth
import docx2txt
import glob
import os
import time

config = {
    "DEBUG": False,             # some Flask specific configs
    "CACHE_TYPE": "simple",     # Flask-Caching related configs
    "CACHE_DEFAULT_TIMEOUT": 1
}

app = Flask(__name__)           #initialize app
app.config.from_mapping(config) #set app with configurations
cache = Cache(app)              #create cache instance for app

UPLOAD_FOLDER = './tempFiles'   #folder path for uploading documents
ALLOWED_EXTENSIONS = {'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def clearFile():
    df = pd.DataFrame()
    with pd.ExcelWriter('./tempFiles/output.xlsx') as writer:
        df.to_excel(writer,sheet_name = "null")
    return '';

class MyThread(Thread):
    def __init__(self,event):
        Thread.__init__(self)
        self.stopped = event

    def run(self):
        time.sleep(3)
        print("clearing files")
        clearFile()


@app.route('/')
def index():
    clearFile()
    return render_template('index.html')

@app.route('/instructions', methods=['GET', 'POST'])
def convert():
    if request.method == 'POST':
        
        files = request.files.getlist("file[]") #Read files

        print (files)
       
        DocxFiles = []
        for file in files:
            if file.filename == '':
                return redirect(request.url)
            if (file.filename).split('.')[1] == "docx": #if file isnt docx then dont parse
                DocxFiles.append(file)

        print(DocxFiles)
        
        headers = []
        htmlFiles = []
        
        for file in DocxFiles:
            fn = file.filename.split('.')[0]
            fe = file.filename.split('.')[1]
            fp = './tempFiles/'+file.filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], fn+'.'+fe))

            text = docx2txt.process(fp)
            L = text.split('\n')
            L2 = []
            for x in range(len(L)):
                if L[x] != '':
                    L2.append(L[x])

            keys =[['Stability Program #:',
                    'Product Code / Master #:'],
                    ['Lot #:',
                    'Theoretical Batch Size:'],
                    ['Storage Conditions:',
                    'Packaging Description:'],
                    ['Packaging Description:',
                    'Active Claims:']]

            HeaderInfo = []
            
            for x in keys:
                start = L2.index(x[0])
                end = L2.index(x[1])
                if (end - start) > 1:
                    HeaderInfo.append(x[0]+" "+L2[start +1])
                else:
                    HeaderInfo.append(x[0]+" N/A")

            
            with open(fp, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                html = result.value # The generated HTML
                messages = result.messages

                f = open('./tempFiles/'+fn+'.html',"w")
                f.write(html)
                f.close()
                
            headers.append(HeaderInfo)
            htmlFiles.append('./tempFiles/'+fn+'.html')

        sheets = []    
        for file in range(len(htmlFiles)):
            df = pd.DataFrame(headers[file])
            dflst = pd.read_html(htmlFiles[file])
            df2 = dflst[0]
            for x in range(len(dflst) - 1):
                    df2.append(dflst[x+1],ignore_index=True)

            result = pd.concat([df,df2],axis=1)

            sheets.append(result)
            
        #clean files not being used
        folder = './tempFiles'
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))

        #save to an xlsx with multiple sheets    
        with pd.ExcelWriter('./tempFiles/output.xlsx') as writer:
            x = 0
            for df in sheets:
                df.to_excel(writer,sheet_name=DocxFiles[x].filename)
                x += 1       
        
    return render_template('instructions.html')

@app.route('/download')
def downloadScreen():

    #clear cache to download updated file
    with app.app_context():
        cache.clear()
        
    stopFlag = Event()
    thread = MyThread(stopFlag)
    thread.start()
    
    #download file
    return send_file('./tempFiles/output.xlsx',as_attachment=True, cache_timeout=0)

from werkzeug.middleware.shared_data import SharedDataMiddleware
app.add_url_rule('/uploads/<filename>', 'uploaded_file',build_only=True)
app.wsgi_app = SharedDataMiddleware(app.wsgi_app, {'/uploads':  app.config['UPLOAD_FOLDER']})

if __name__ == "__main__":
    app.run()
