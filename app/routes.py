# -*- coding: utf-8 -*-
"""
Created on Thu Aug 23 13:33:38 2018

@author: denisse.perez
"""
from werkzeug.utils import secure_filename

#from app import app
import pandas as pd
from flask import render_template, request, Flask
import os

import xlrd

app = Flask(__name__)

#@app.route('/')
@app.route('/', methods = ['GET', 'POST'])
def index():

    try:
        try:
            file = request.files['file']
            #filename = secure_filename(file.filename)
            #print(filename)
            file.save(os.path.join(app.root_path,"../files/Input.xlsx"))
        except Exception as e:
            print(e)
        base = pd.read_excel(os.path.join(app.root_path,"../files/Input.xlsx"), sheet_name="Base", index_col=False)
        activity_type = base['Activity Type']
        activity_name = base['Activity Name']
        portfolio = base['Portfolio']
        base['MEDIA COST (000\'S)'] = round(base['MEDIA COST (000\'S)'], 2)

        # Gold Activities (First Matrice's row)

        GC = base.loc[(activity_type == 'GOLD') & (portfolio == 'Core') & (activity_name != 'AON') & (
                    activity_name != 'E-Commerce')]
        gold_cores = pd.DataFrame
        gold_cores = {'Activity': GC['Key'], 'iTO': GC['iTO Euros (000\'S)'], 'Media Cost': GC['MEDIA COST (000\'S)']}
        gold_cores = (pd.DataFrame(gold_cores)).reset_index()
        gold_cores.index = gold_cores.index + 1
        gold_cores = gold_cores.drop(gold_cores.columns[0], axis=1)

        GFC = base.loc[(activity_type == 'GOLD') & (portfolio == 'Future Core')]
        gold_FutureCores = pd.DataFrame
        gold_FutureCores = {'Activity': GFC['Key'], 'iTO': GFC['iTO Euros (000\'S)'],
                            'Media Cost': GFC['MEDIA COST (000\'S)']}
        gold_FutureCores = (pd.DataFrame(gold_FutureCores)).reset_index()
        gold_FutureCores.index = gold_FutureCores.index + 1
        gold_FutureCores = gold_FutureCores.drop(gold_FutureCores.columns[0], axis=1)

        GS = base.loc[(activity_type == 'GOLD') & (portfolio == 'Supporter')]
        gold_supporter = pd.DataFrame
        gold_supporter = {'Activity': GS['Key'], 'iTO': GS['iTO Euros (000\'S)'],
                          'Media Cost': GS['MEDIA COST (000\'S)']}
        gold_supporter = (pd.DataFrame(gold_supporter)).reset_index()
        gold_supporter.index = gold_supporter.index + 1
        gold_supporter = gold_supporter.drop(gold_supporter.columns[0], axis=1)

        # Silver Activities (Second Matrice's row)

        SC = base.loc[(activity_type == 'SILVER') & (portfolio == 'Core') & (activity_name != 'AON') & (
                    activity_name != 'E-Commerce')]
        silver_cores = pd.DataFrame
        silver_cores = {'Activity': SC['Key'], 'iTO': SC['iTO Euros (000\'S)'], 'Media Cost': SC['MEDIA COST (000\'S)']}
        silver_cores = (pd.DataFrame(silver_cores)).reset_index()
        silver_cores.index = silver_cores.index + 1
        silver_cores = silver_cores.drop(silver_cores.columns[0], axis=1)

        SFC = base.loc[(activity_type == 'SILVER') & (portfolio == 'Future Core')]
        silver_FutureCores = pd.DataFrame
        silver_FutureCores = {'Activity': SFC['Key'], 'iTO': SFC['iTO Euros (000\'S)'],
                              'Media Cost': SFC['MEDIA COST (000\'S)']}
        silver_FutureCores = (pd.DataFrame(silver_FutureCores)).reset_index()
        silver_FutureCores.index = silver_FutureCores.index + 1
        silver_FutureCores = silver_FutureCores.drop(silver_FutureCores.columns[0], axis=1)

        SS = base.loc[(activity_type == 'SILVER') & (portfolio == 'Supporter')]
        silver_supporter = pd.DataFrame
        silver_supporter = {'Activity': SS['Key'], 'iTO': SS['iTO Euros (000\'S)'],
                            'Media Cost': SS['MEDIA COST (000\'S)']}
        silver_supporter = (pd.DataFrame(silver_supporter)).reset_index()
        silver_supporter.index = silver_supporter.index + 1
        silver_supporter = silver_supporter.drop(silver_supporter.columns[0], axis=1)

        # Bronze Activities (Third Matrice's row)

        BC = base.loc[(activity_type == 'BRONZE') & (portfolio == 'Core') & (activity_name != 'AON') & (
                    activity_name != 'E-Commerce')]
        bronze_cores = pd.DataFrame
        bronze_cores = {'Activity': BC['Key'], 'iTO': BC['iTO Euros (000\'S)'], 'Media Cost': BC['MEDIA COST (000\'S)']}
        bronze_cores = (pd.DataFrame(bronze_cores)).reset_index()
        bronze_cores.index = bronze_cores.index + 1
        bronze_cores = bronze_cores.drop(bronze_cores.columns[0], axis=1)

        BFC = base.loc[(activity_type == 'BRONZE') & (portfolio == 'Future Core') & (activity_name != 'AON') & (
                    activity_name != 'E-Commerce')]
        bronze_FutureCores = pd.DataFrame
        bronze_FutureCores = {'Activity': BFC['Key'], 'iTO': BFC['iTO Euros (000\'S)'],
                              'Media Cost': BFC['MEDIA COST (000\'S)']}
        bronze_FutureCores = (pd.DataFrame(bronze_FutureCores)).reset_index()
        bronze_FutureCores.index = bronze_FutureCores.index + 1
        bronze_FutureCores = bronze_FutureCores.drop(bronze_FutureCores.columns[0], axis=1)

        BS = base.loc[(activity_type == 'BRONZE') & (portfolio == 'Supporter')]
        bronze_supporter = pd.DataFrame
        bronze_supporter = {'Activity': BS['Key'], 'iTO': BS['iTO Euros (000\'S)'],
                            'Media Cost': BS['MEDIA COST (000\'S)']}
        bronze_supporter = (pd.DataFrame(bronze_supporter)).reset_index()
        bronze_supporter.index = bronze_supporter.index + 1
        bronze_supporter = bronze_supporter.drop(bronze_supporter.columns[0], axis=1)
        r = render_template('index.html', title='Home', gold_cores=gold_cores.to_html(),
                    gold_FutureCores=gold_FutureCores.to_html(), gold_supporter=gold_supporter.to_html(),
                    silver_cores=silver_cores.to_html(), silver_FutureCores=silver_FutureCores.to_html(),
                    silver_supporter=silver_supporter.to_html(),
                    bronze_cores=bronze_cores.to_html(), bronze_FutureCores=bronze_FutureCores.to_html(),
                    bronze_supporter=bronze_supporter.to_html())
   # #print(r)
        return r
    except Exception as e:
        r = render_template('index.html', title='Home')







if __name__ == "__main__":
    app.run(debug=True)