from flask import Flask, render_template, request, redirect, url_for, make_response, Markup

import pandas
import json

fin_string = ""

app = Flask(__name__)

@app.route('/', methods = ['GET'])
def index():
    return render_template(("index.html"), out_str=fin_string)

@app.route('/', methods=['POST'])
def upload_file():
    uploaded_file = request.files['file']

    #print(uploaded_file.filename)

    if "xlsx" not in uploaded_file.filename.split("."):
        return redirect(url_for('index'))


    valid_client_keys = ["Adresa do zaměstnání",
                        "Adresa klienta",
                        "IČ?",
                        "Kient",
                        "Kontakt na klienta",
                        "Název firmy",
                        "PSČ",
                        "Pozice",
                        "Spr."
                    ]

    # excel -> pandas
    excel_data_df = pandas.read_excel(uploaded_file)

    # Convert excel to string 
    # (define orientation of document in this case from up to down)
    thisisjson = excel_data_df.to_json(orient='records')

    # Make the string into a list to be able to input in to a JSON-file
    thisisjson_dict = json.loads(thisisjson)

    # array of unwanted columns in excel table
    keys_to_delete = []

    #print(thisisjson_dict)

    # find non-valid columns ( via valid_client_keys[] )
    for client in thisisjson_dict:
        for item in client:
            if item not in valid_client_keys:
                keys_to_delete.append(item)

    # remove non-valid found in for-loop above
    for client, item in zip(thisisjson_dict, keys_to_delete):
        del client[item]


    def handle_one_client(client : dict) -> str :
        output_string = ""
        for item in client:
            output_string += f"{item} : {client[item]}<br>"
        output_string += "<br>"
        return output_string

    global fin_string
    fin_string = ""    

    for client in thisisjson_dict:
        fin_string += handle_one_client(client)


    if uploaded_file.filename != '':
        uploaded_file.save(uploaded_file.filename)
    return redirect(url_for('index', out_str=Markup(fin_string)))