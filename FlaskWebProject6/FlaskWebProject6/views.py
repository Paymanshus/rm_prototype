from pymongo import MongoClient
from datetime import datetime
from flask import Flask, render_template, request, url_for,redirect
import os 
from bson import ObjectId   
from flask_pymongo import PyMongo
from FlaskWebProject6 import app
#from flask_apscheduler import APScheduler 
from threading import Timer
from flask import send_file
from xlsxwriter import Workbook

app.config["MONGO_URI"] = "mongodb://WorkAmp%5FMVP:ampitup%40futurex@workamp-mvp-shard-00-00-buabm.mongodb.net:27017,workamp-mvp-shard-00-01-buabm.mongodb.net:27017,workamp-mvp-shard-00-02-buabm.mongodb.net:27017/test?ssl=true&replicaSet=WorkAmp-MVP-shard-0&authSource=admin&retryWrites=true&w=majority"
mongo = PyMongo(app)

client = MongoClient("mongodb://WorkAmp%5FMVP:ampitup%40futurex@workamp-mvp-shard-00-00-buabm.mongodb.net:27017,workamp-mvp-shard-00-01-buabm.mongodb.net:27017,workamp-mvp-shard-00-02-buabm.mongodb.net:27017/test?ssl=true&replicaSet=WorkAmp-MVP-shard-0&authSource=admin&retryWrites=true&w=majority")
#client.server_info()    
db = client.get_database('Demo01')
records = db['Daily Activities']
washroom_checklist = db['Washroom Checklist']
fridge_checklist = db['Fridge Checklist']
huddle_checklist = db['Huddle Room Checklist']
meeting_checklist = db['Meeting Room Checklist']
monthly_checklist = db["Monthly Checklist"]
weekly_checklist = db["Weekly Checklist"]
pantry = db["Pantry"]
fnb = db["fnb"]
office_supp = db["Office Supplies"]
housekeeping = db["Housekeeping"]
finance = db["Finance"]
snag = db["Snag"]


def redirect_url():
    return request.args.get('next') or \
           request.referrer or \
           url_for('admin')


@app.route('/')

@app.route('/login')
def login():
    return render_template('login.html')

@app.route('/admin_dash')
def admin_dash():
    wrem = washroom_checklist.find({"Status":"None"}).count()
    frem = fridge_checklist.find({"Status":"None"}).count()
    hrem = huddle_checklist.find({"Status":"None"}).count()
    mrem = meeting_checklist.find({"Status":"None"}).count()
    rem = wrem+frem+hrem+mrem

    pantryc = pantry.find({"Quantity":""}).count()
    fnbc = fnb.find({"Quantity":""}).count()
    office_suppc = office_supp.find({"Quantity":""}).count()
    housekeepingc = housekeeping.find({"Quantity":""}).count()
    totalc = pantryc+fnbc+office_suppc+housekeepingc


    monrem = monthly_checklist.find({"Status":"None"}).count()
    taskrem = records.find({"done":"no"}).count()
    return render_template('admin_dash.html', rem=rem,totalc = totalc, monrem = monrem,taskrem= taskrem)


@app.route('/rec_expense_dash')
def rec_expense_dash():
    wrem = washroom_checklist.find({"Status":"None"}).count()
    frem = fridge_checklist.find({"Status":"None"}).count()
    hrem = huddle_checklist.find({"Status":"None"}).count()
    mrem = meeting_checklist.find({"Status":"None"}).count()
    rem = wrem+frem+hrem+mrem

    pantryc = pantry.find({"Quantity":""}).count()
    fnbc = fnb.find({"Quantity":""}).count()
    office_suppc = office_supp.find({"Quantity":""}).count()
    housekeepingc = housekeeping.find({"Quantity":""}).count()
    totalc = pantryc+fnbc+office_suppc+housekeepingc


    monrem = monthly_checklist.find({"Status":"None"}).count()
    taskrem = records.find({"done":"no"}).count()

    return render_template('rec_expense_dash.html', rem=rem,totalc = totalc, monrem = monrem,taskrem= taskrem)

@app.route("/rec_expense_dashfunc", methods=['POST'])    
def rec_expense_dashfunc():
    finance.insert_one({"Title":request.form['Title'],"Amount":request.form['amount'],"Category":request.form.get('category'),"Subcategory":request.form.get('categorysub'),"Date of Payment":request.form['date'],"Invoice ID":request.form['invoice'],"GST no":request.form['gst'],"Payment Mode":request.form.get('payment')})

    redir=redirect_url()        
    return redirect(redir)  










@app.route('/checklist_dash')
def checklist_dash():
    wrem = washroom_checklist.find({"Status":"None"}).count()
    frem = fridge_checklist.find({"Status":"None"}).count()
    hrem = huddle_checklist.find({"Status":"None"}).count()
    mrem = meeting_checklist.find({"Status":"None"}).count()
    return render_template('checklist_dash.html',wrem = wrem, frem = frem, hrem = hrem, mrem = mrem)

@app.route("/checklist")
def checklist():
    washroom_list = washroom_checklist.find({"Status":"None"})
    washroom_done_list = washroom_checklist.find({"Status": {"$in": ["Okay","Not Okay"]}})
    return render_template('checklist.html',washroom_list = washroom_list, washroom_done_list = washroom_done_list)


@app.route("/okaydrop", methods=['POST'])    
def okaydrop():    


    flush = request.form.get("Flush")
    taps = request.form.get("Taps")
    cleanliness_water = request.form.get("Cleanliness of water from taps")
    door = request.form.get("Door mechanism of all cubicles")
    smell = request.form.get("Smell")
    mirror = request.form.get("Mirror")
    handshower = request.form.get("Handshower leakage and mechanism")

    if(str(flush)=="Okay"):
        washroom_checklist.update({"Item":"Flush"}, {"$set": {"Status":"Okay"}})
    elif(str(flush)=="Not Okay"):
        washroom_checklist.update({"Item":"Flush"}, {"$set": {"Status":"Not Okay"}})

    if(str(taps)=="Okay"):
        washroom_checklist.update({"Item":"Taps"}, {"$set": {"Status":"Okay"}})
    elif(str(taps)=="Not Okay"):
        washroom_checklist.update({"Item":"Taps"}, {"$set": {"Status":"Not Okay"}})

    if(str(cleanliness_water)=="Okay"):
        washroom_checklist.update({"Item":"Cleanliness of water from taps"}, {"$set": {"Status":"Okay"}})
    elif(str(cleanliness_water)=="Not Okay"):
        washroom_checklist.update({"Item":"Cleanliness of water from taps"}, {"$set": {"Status":"Not Okay"}})

    if(str(door)=="Okay"):
        washroom_checklist.update({"Item":"Door mechanism of all cubicles"}, {"$set": {"Status":"Okay"}})
    elif(str(door)=="Not Okay"):
        washroom_checklist.update({"Item":"Door mechanism of all cubicles"}, {"$set": {"Status":"Not Okay"}})

    if(str(smell)=="Okay"):
        washroom_checklist.update({"Item":"Smell"}, {"$set": {"Status":"Okay"}})
    elif(str(smell)=="Not Okay"):
        washroom_checklist.update({"Item":"Smell"}, {"$set": {"Status":"Not Okay"}})

    if(str(mirror)=="Okay"):
        washroom_checklist.update({"Item":"Mirror"}, {"$set": {"Status":"Okay"}})
    elif(str(mirror)=="Not Okay"):
        washroom_checklist.update({"Item":"Mirror"}, {"$set": {"Status":"Not Okay"}})

    if(str(handshower)=="Okay"):
        washroom_checklist.update({"Item":"Handshower leakage and mechanism"}, {"$set": {"Status":"Okay"}})
    elif(str(handshower)=="Not Okay"):
        washroom_checklist.update({"Item":"Handshower leakage and mechanism"}, {"$set": {"Status":"Not Okay"}})
    else:
        pass

    washroom_checklist.update_many({'Status': 'Okay'}, {"$set": { "done": "yes" }})
        
    if(washroom_checklist.find({"Item":"Flush"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Flush","done":"no","createdAt": datetime.now(),"Category":"Washroom"})
    if(washroom_checklist.find({"Item":"Taps"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Taps","done":"no","createdAt":datetime.now(),"Category":"Washroom"})
    if(washroom_checklist.find({"Item":"Cleanliness of water from taps"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Cleanliness of water from taps","done":"no","createdAt":datetime.now(),"Category":"Washroom"})
    if(washroom_checklist.find({"Item":"Door mechanism of all cubicles"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Door mechanism of all cubicles","done":"no","createdAt":datetime.now(),"Category":"Washroom"})
    if(washroom_checklist.find({"Item":"Smell"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Smell","done":"no","createdAt":datetime.now(),"Category":"Washroom"})
    if(washroom_checklist.find({"Item":"Mirror"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Mirror","done":"no","createdAt":datetime.now(),"Category":"Washroom"})
    if(washroom_checklist.find({"Item":"Handshower leakage and mechanism"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Handshower leakage and mechanism","done":"no","createdAt":datetime.now(),"Category":"Washroom"})
    else:
        pass

    return redirect(url_for('admin'))

@app.route("/checklist_fridge")
def checklist_fridge():
    fridge_list = fridge_checklist.find({"Status":"None"})
    fridge_done_list = fridge_checklist.find({"Status": {"$in": ["Okay","Not Okay"]}})

    return render_template('checklist_fridge.html',fridge_list = fridge_list,fridge_done_list = fridge_done_list)

@app.route("/okaydrop_fridge", methods=['POST'])    
def okaydrop_fridge():    

    door = request.form.get("Door mechanism")
    cool = request.form.get("Cooling")

    if(str(door)=="Okay"):
        fridge_checklist.update({"Item":"Door mechanism"}, {"$set": {"Status":"Okay"}})
    elif(str(door)=="Not Okay"):
        fridge_checklist.update({"Item":"Door mechanism"}, {"$set": {"Status":"Not Okay"}})

    if(str(cool)=="Okay"):
        fridge_checklist.update({"Item":"Cooling"}, {"$set": {"Status":"Okay"}})
    elif(str(cool)=="Not Okay"):
        fridge_checklist.update({"Item":"Cooling"}, {"$set": {"Status":"Not Okay"}})
    else:
        pass

    fridge_checklist.update_many({'Status': 'Okay'}, {"$set": { "done": "yes" }})
        
    if(fridge_checklist.find({"Item":"Door mechanism"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Door mechanism","done":"no","createdAt":datetime.now(),"Category":"Fridge"})
    if(fridge_checklist.find({"Item":"Cooling"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Cooling","done":"no","createdAt":datetime.now(),"Category":"Fridge"})
    else:
        pass

    return redirect(url_for('admin'))

@app.route("/checklist_huddleroom")
def checklist_huddleroom():
    huddle_list = huddle_checklist.find({"Status":"None"})
    huddle_done_list = huddle_checklist.find({"Status": {"$in": ["Okay","Not Okay"]}})
    return render_template('checklist_huddleroom.html',huddle_list = huddle_list, huddle_done_list = huddle_done_list)


@app.route("/okaydrop_huddleroom", methods=['POST'])    
def okaydrop_huddleroom():    

    hdmi = request.form.get("HDMI")
    damage = request.form.get("Chair and table damage")
    markers = request.form.get("No. of markers")
    wboard = request.form.get("Writing boards")
    remote = request.form.get("Remote")
    wear = request.form.get("Wear and tear")
    tv = request.form.get("TV")

    if(str(hdmi)=="Okay"):
        huddle_checklist.update({"Item":"HDMI"}, {"$set": {"Status":"Okay"}})
    elif(str(hdmi)=="Not Okay"):
        huddle_checklist.update({"Item":"HDMI"}, {"$set": {"Status":"Not Okay"}})

    if(str(damage)=="Okay"):
        huddle_checklist.update({"Item":"Chair and table damage"}, {"$set": {"Status":"Okay"}})
    elif(str(damage)=="Not Okay"):
        huddle_checklist.update({"Item":"Chair and table damage"}, {"$set": {"Status":"Not Okay"}})

    if(str(markers)=="Okay"):
        huddle_checklist.update({"Item":"No. of markers"}, {"$set": {"Status":"Okay"}})
    elif(str(markers)=="Not Okay"):
        huddle_checklist.update({"Item":"No. of markers"}, {"$set": {"Status":"Not Okay"}})

    if(str(wboard)=="Okay"):
        huddle_checklist.update({"Item":"Writing boards"}, {"$set": {"Status":"Okay"}})
    elif(str(wboard)=="Not Okay"):
        huddle_checklist.update({"Item":"Writing boards"}, {"$set": {"Status":"Not Okay"}})

    if(str(remote)=="Okay"):
        huddle_checklist.update({"Item":"Remote"}, {"$set": {"Status":"Okay"}})
    elif(str(remote)=="Not Okay"):
        huddle_checklist.update({"Item":"Remote"}, {"$set": {"Status":"Not Okay"}})

    if(str(wear)=="Okay"):
        huddle_checklist.update({"Item":"Wear and tear"}, {"$set": {"Status":"Okay"}})
    elif(str(wear)=="Not Okay"):
        huddle_checklist.update({"Item":"Wear and tear"}, {"$set": {"Status":"Not Okay"}})

    if(str(tv)=="Okay"):
        huddle_checklist.update({"Item":"TV"}, {"$set": {"Status":"Okay"}})
    elif(str(tv)=="Not Okay"):
        huddle_checklist.update({"Item":"TV"}, {"$set": {"Status":"Not Okay"}})
    else:
        pass

    huddle_checklist.update_many({'Status': 'Okay'}, {"$set": { "done": "yes" }})
        
    if(huddle_checklist.find({"Item":"HDMI"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"HDMI","done":"no","createdAt": datetime.now(),"Category":"Huddle Room"})
    if(huddle_checklist.find({"Item":"Chair and table damage"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Chair and table damage","done":"no","createdAt": datetime.now(),"Category":"Huddle Room"})
    if(huddle_checklist.find({"Item":"No. of markers"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"No. of markers","done":"no","createdAt": datetime.now(),"Category":"Huddle Room"})
    if(huddle_checklist.find({"Item":"Writing boards"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Writing boards","done":"no","createdAt": datetime.now(),"Category":"Huddle Room"})
    if(huddle_checklist.find({"Item":"Remote"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Remote","done":"no", "createdAt": datetime.now(),"Category":"Huddle Room"})
    if(huddle_checklist.find({"Item":"Wear and tear"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Wear and tear","done":"no", "createdAt": datetime.now(),"Category":"Huddle Room"})
    if(huddle_checklist.find({"Item":"TV"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"TV","done":"no", "createdAt": datetime.now(),"Category":"Huddle Room"})
    else:
        pass

    return redirect(url_for('admin'))

@app.route("/checklist_meeting")
def checklist_meeting():
    meeting_list = meeting_checklist.find({"Status":"None"})
    meeting_done_list = meeting_checklist.find({"Status": {"$in": ["Okay","Not Okay"]}})
    return render_template('checklist_meeting.html',meeting_list = meeting_list, meeting_done_list = meeting_done_list)


@app.route("/okaydrop_meeting", methods=['POST'])    
def okaydrop_meeting():    

    hdmi = request.form.get("HDMI")
    damage = request.form.get("Chair and table damage")
    markers = request.form.get("No. of markers")
    clean = request.form.get("Water glasses clean")
    remote = request.form.get("Remote")
    wear = request.form.get("Wear and tear")
    tv = request.form.get("TV")
    screen = request.form.get("Screen sharing")
    duster = request.form.get("Duster")
    plug = request.form.get("Plug points functional")
    broken = request.form.get("Water Jug broken")

    if(str(hdmi)=="Okay"):
        meeting_checklist.update({"Item":"HDMI"}, {"$set": {"Status":"Okay"}})
    elif(str(hdmi)=="Not Okay"):
        meeting_checklist.update({"Item":"HDMI"}, {"$set": {"Status":"Not Okay"}})

    if(str(damage)=="Okay"):
        meeting_checklist.update({"Item":"Chair and table damage"}, {"$set": {"Status":"Okay"}})
    elif(str(damage)=="Not Okay"):
        meeting_checklist.update({"Item":"Chair and table damage"}, {"$set": {"Status":"Not Okay"}})

    if(str(markers)=="Okay"):
        meeting_checklist.update({"Item":"No. of markers"}, {"$set": {"Status":"Okay"}})
    elif(str(markers)=="Not Okay"):
        meeting_checklist.update({"Item":"No. of markers"}, {"$set": {"Status":"Not Okay"}})

    if(str(clean)=="Okay"):
        meeting_checklist.update({"Item":"Water glasses clean"}, {"$set": {"Status":"Okay"}})
    elif(str(clean)=="Not Okay"):
        meeting_checklist.update({"Item":"Water glasses clean"}, {"$set": {"Status":"Not Okay"}})

    if(str(remote)=="Okay"):
        meeting_checklist.update({"Item":"Remote"}, {"$set": {"Status":"Okay"}})
    elif(str(remote)=="Not Okay"):
        meeting_checklist.update({"Item":"Remote"}, {"$set": {"Status":"Not Okay"}})

    if(str(wear)=="Okay"):
        meeting_checklist.update({"Item":"Wear and tear"}, {"$set": {"Status":"Okay"}})
    elif(str(wear)=="Not Okay"):
        meeting_checklist.update({"Item":"Wear and tear"}, {"$set": {"Status":"Not Okay"}})

    if(str(tv)=="Okay"):
        meeting_checklist.update({"Item":"TV"}, {"$set": {"Status":"Okay"}})
    elif(str(tv)=="Not Okay"):
        meeting_checklist.update({"Item":"TV"}, {"$set": {"Status":"Not Okay"}})


    if(str(screen)=="Okay"):
        meeting_checklist.update({"Item":"Screen sharing"}, {"$set": {"Status":"Okay"}})
    elif(str(screen)=="Not Okay"):
        meeting_checklist.update({"Item":"Screen sharing"}, {"$set": {"Status":"Not Okay"}})


    if(str(duster)=="Okay"):
        meeting_checklist.update({"Item":"Duster"}, {"$set": {"Status":"Okay"}})
    elif(str(duster)=="Not Okay"):
        meeting_checklist.update({"Item":"Duster"}, {"$set": {"Status":"Not Okay"}})


    if(str(plug)=="Okay"):
        meeting_checklist.update({"Item":"Plug points functional"}, {"$set": {"Status":"Okay"}})
    elif(str(plug)=="Not Okay"):
        meeting_checklist.update({"Item":"Plug points funtional"}, {"$set": {"Status":"Not Okay"}})


    if(str(broken)=="Okay"):
        meeting_checklist.update({"Item":"Water Jug broken"}, {"$set": {"Status":"Okay"}})
    elif(str(broken)=="Not Okay"):
        meeting_checklist.update({"Item":"Water Jug broken"}, {"$set": {"Status":"Not Okay"}})
    else:
        pass

    meeting_checklist.update_many({'Status': 'Okay'}, {"$set": { "done": "yes" }})
        
    if(meeting_checklist.find({"Item":"HDMI"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"HDMI","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"Chair and table damage"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Chair and table damage","done":"no","createdAt":datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"No. of markers"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"No. of markers","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"Water glasses clean"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Water glasses clean","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"Remote"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Remote","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"Wear and tear"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Wear and tear","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"TV"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"TV","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"Duster"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Duster","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"Screen sharing"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Screen sharing","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"Plug points functional"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Plug points functional","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    if(meeting_checklist.find({"Item":"Water Jug broken"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Water Jug broken","done":"no","createdAt": datetime.now(),"Category":"Meeting Room"})
    else:
        pass

    return redirect(url_for('admin'))


@app.route("/checklist_dash_weekly")
def checklist_dash_weekly():
    pantryc = pantry.find({"Quantity":""}).count()
    fnbc = fnb.find({"Quantity":""}).count()
    office_suppc = office_supp.find({"Quantity":""}).count()
    housekeepingc = housekeeping.find({"Quantity":""}).count()

    return render_template('checklist_dash_weekly.html',pantryc = pantryc, fnbc = fnbc, office_suppc = office_suppc, housekeepingc = housekeepingc)


@app.route("/checklist_weekly")
def checklist_weekly():
    pantry_list = pantry.find()
    return render_template('checklist_weekly.html',pantry_list = pantry_list)

@app.route("/inventory", methods = ['POST'])
def inventory():

    pantry.update({"Item":"Plate"}, {"$set": {"Quantity":request.form['Plate']}})
    pantry.update({"Item":"Spoon"}, {"$set": {"Quantity":request.form['Spoon']}})
    pantry.update({"Item":"Fork"}, {"$set": {"Quantity":request.form['Fork']}})
    pantry.update({"Item":"Bowl"}, {"$set": {"Quantity":request.form['Bowl']}})
    pantry.update({"Item":"Cup"}, {"$set": {"Quantity":request.form['Cup']}})
    pantry.update({"Item":"Mug"}, {"$set": {"Quantity":request.form['Mug']}})
    pantry.update({"Item":"Knife"}, {"$set": {"Quantity":request.form['Knife']}})
    pantry.update({"Item":"Glass"}, {"$set": {"Quantity":request.form['Glass']}})
    pantry.update({"Item":"Serving Tray"}, {"$set": {"Quantity":request.form['Serving Tray']}})
    pantry.update({"Item":"Water Jar"}, {"$set": {"Quantity":request.form['Water Jar']}})
    pantry.update({"Item":"Microwave"}, {"$set": {"Quantity":request.form['Microwave']}})
    pantry.update({"Item":"Induction"}, {"$set": {"Quantity":request.form['Induction']}})
    pantry.update({"Item":"Fridge"}, {"$set": {"Quantity":request.form['Fridge']}})
    pantry.update({"Item":"Infused Water Dispenser"}, {"$set": {"Quantity":request.form['Infused Water Dispenser']}})
    pantry.update({"Item":"Aquaguard"}, {"$set": {"Quantity":request.form['Aquaguard']}})
    pantry.update({"Item":"Coffee Machine"}, {"$set": {"Quantity":request.form['Coffee Machine']}})
    pantry.update({"Item":"Paper cup"}, {"$set": {"Quantity":request.form['Paper cup']}})
    pantry.update({"Item":"Stirrer"}, {"$set": {"Quantity":request.form['Stirrer']}})
    pantry.update({"Item":"Tissue Paper"}, {"$set": {"Quantity":request.form['Tissue Paper']}})
    pantry.update({"Item":"Disposable cutlery"}, {"$set": {"Quantity":request.form['Disposable cutlery']}})

    redir=redirect_url()        
    return redirect(redir)  


@app.route("/checklist_weekly_fnb")
def checklist_weekly_fnb():
    fnb_list = fnb.find()
    return render_template('checklist_weekly_fnb.html',fnb_list = fnb_list)

@app.route("/inventory_fnb", methods = ['POST'])
def inventory_fnb():

    fnb.update({"Item":"Coffee"}, {"$set": {"Quantity":request.form['Coffee']}})
    fnb.update({"Item":"Tea"}, {"$set": {"Quantity":request.form['Tea']}})
    fnb.update({"Item":"Sugar"}, {"$set": {"Quantity":request.form['Sugar']}})
    fnb.update({"Item":"Biscuit"}, {"$set": {"Quantity":request.form['Biscuit']}})
    fnb.update({"Item":"Cold Drink"}, {"$set": {"Quantity":request.form['Cold Drink']}})
    fnb.update({"Item":"Mineral Water"}, {"$set": {"Quantity":request.form['Mineral Water']}})
    fnb.update({"Item":"Milk"}, {"$set": {"Quantity":request.form['Milk']}})
    fnb.update({"Item":"Salt & Pepper"}, {"$set": {"Quantity":request.form['Salt & Pepper']}})
    fnb.update({"Item":"Chocolates"}, {"$set": {"Quantity":request.form['Chocolates']}})
   

    redir=redirect_url()        
    return redirect(redir)  

@app.route("/checklist_weekly_office_supp")
def checklist_weekly_office_supp():
    office_supp_list = office_supp.find()
    return render_template('checklist_weekly_office_supp.html',office_supp_list = office_supp_list)

@app.route("/inventory_office_supp", methods = ['POST'])
def inventory_office_supp():

    office_supp.update({"Item":"Pen"}, {"$set": {"Quantity":request.form['Pen']}})
    office_supp.update({"Item":"Pencil"}, {"$set": {"Quantity":request.form['Pencil']}})
    office_supp.update({"Item":"Notepad"}, {"$set": {"Quantity":request.form['Notepad']}})
    office_supp.update({"Item":"Cello Tape"}, {"$set": {"Quantity":request.form['Cello Tape']}})
    office_supp.update({"Item":"Printing paper"}, {"$set": {"Quantity":request.form['Printing paper']}})
    office_supp.update({"Item":"Stapler"}, {"$set": {"Quantity":request.form['Stapler']}})
    office_supp.update({"Item":"Stapler Pins"}, {"$set": {"Quantity":request.form['Stapler Pins']}})
    office_supp.update({"Item":"Flat file"}, {"$set": {"Quantity":request.form['Flat file']}})
    office_supp.update({"Item":"Box file"}, {"$set": {"Quantity":request.form['Box file']}})
    office_supp.update({"Item":"Plastic folder"}, {"$set": {"Quantity":request.form['Plastic folder']}})
    office_supp.update({"Item":"Envelope"}, {"$set": {"Quantity":request.form['Envelope']}})
    office_supp.update({"Item":"Sticky notes"}, {"$set": {"Quantity":request.form['Sticky notes']}})
    office_supp.update({"Item":"Diary"}, {"$set": {"Quantity":request.form['Diary']}})
    office_supp.update({"Item":"Push pin"}, {"$set": {"Quantity":request.form['Push pin']}})
    office_supp.update({"Item":"Paper pin"}, {"$set": {"Quantity":request.form['Paper pin']}})
    office_supp.update({"Item":"Stamp pad"}, {"$set": {"Quantity":request.form['Stamp pad']}})
    office_supp.update({"Item":"Stamp Ink"}, {"$set": {"Quantity":request.form['Stamp Ink']}})


    redir=redirect_url()        
    return redirect(redir) 


@app.route("/checklist_weekly_housekeeping")
def checklist_weekly_housekeeping():
    houskeeping_list = housekeeping.find()
    return render_template('checklist_weekly_housekeeping.html',housekeeping_list = houskeeping_list)

@app.route("/inventory_housekeeping", methods = ['POST'])
def inventory_housekeeping():

    housekeeping.update({"Item":"Wipe Cloth"}, {"$set": {"Quantity":request.form['Wipe Cloth']}})
    housekeeping.update({"Item":"Mask"}, {"$set": {"Quantity":request.form['Mask']}})
    housekeeping.update({"Item":"Caddy"}, {"$set": {"Quantity":request.form['Caddy']}})
    housekeeping.update({"Item":"Safety Goggle"}, {"$set": {"Quantity":request.form['Safety Goggle']}})
    housekeeping.update({"Item":"Safety Helmet"}, {"$set": {"Quantity":request.form['Safety Helmet']}})
    housekeeping.update({"Item":"Sign Board"}, {"$set": {"Quantity":request.form['Sign Board']}})
    housekeeping.update({"Item":"Bucket"}, {"$set": {"Quantity":request.form['Bucket']}})
    housekeeping.update({"Item":"Cleaning Chemical"}, {"$set": {"Quantity":request.form['Cleaning Chemical']}})
    housekeeping.update({"Item":"Brush"}, {"$set": {"Quantity":request.form['Brush']}})
    housekeeping.update({"Item":"Gloves"}, {"$set": {"Quantity":request.form['Gloves']}})
    housekeeping.update({"Item":"Sponge"}, {"$set": {"Quantity":request.form['Sponge']}})
    housekeeping.update({"Item":"Mop"}, {"$set": {"Quantity":request.form['Mop']}})
    housekeeping.update({"Item":"Garbage Bag"}, {"$set": {"Quantity":request.form['Garbage Bag']}})
    housekeeping.update({"Item":"Broom"}, {"$set": {"Quantity":request.form['Broom']}})
    housekeeping.update({"Item":"Sweeper"}, {"$set": {"Quantity":request.form['Sweeper']}})
    housekeeping.update({"Item":"Wiper"}, {"$set": {"Quantity":request.form['Wiper']}})
    housekeeping.update({"Item":"Spray Bottle With Trigger"}, {"$set": {"Quantity":request.form['Spray Bottle With Trigger']}})
    housekeeping.update({"Item":"Trolley"}, {"$set": {"Quantity":request.form['Trolley']}})
    housekeeping.update({"Item":"Ladder"}, {"$set": {"Quantity":request.form['Ladder']}})
    housekeeping.update({"Item":"Tissue Paper"}, {"$set": {"Quantity":request.form['Tissue Paper']}})
    housekeeping.update({"Item":"Room freshner"}, {"$set": {"Quantity":request.form['Room freshner']}})
    housekeeping.update({"Item":"Soap"}, {"$set": {"Quantity":request.form['Soap']}})
    housekeeping.update({"Item":"Urinal Screen"}, {"$set": {"Quantity":request.form['Urinal Screen']}}) 

    redir=redirect_url()        
    return redirect(redir)  



@app.route("/checklist_month")
def checklist_month():
    monthly_list = monthly_checklist.find()
    monthly_list = monthly_checklist.find({"Status":"None"})
    monthly_done_list = monthly_checklist.find({"Status": {"$in": ["Okay","Not Okay"]}})
    return render_template('checklist_month.html',monthly_list = monthly_list,monthly_done_list = monthly_done_list)


@app.route("/okaydrop_month", methods=['POST'])    
def okaydrop_month():    


    pest = request.form.get("Pest control")
    electrical_equip = request.form.get("Electrical equip")
    lift = request.form.get("Lift")
    plumbing = request.form.get("Plumbing")
    cp = request.form.get("Carpentary/Polishing")
    deepc = request.form.get("Deep Cleaning")
    

    if(str(pest)=="Okay"):
        monthly_checklist.update({"Title":"Pest control"}, {"$set": {"Status":"Okay"}})
    elif(str(pest)=="Not Okay"):
        monthly_checklist.update({"Title":"Pest control"}, {"$set": {"Status":"Not Okay"}})

    if(str(electrical_equip)=="Okay"):
        monthly_checklist.update({"Title":"Electrical Equip"}, {"$set": {"Status":"Okay"}})
    elif(str(electrical_equip)=="Not Okay"):
        monthly_checklist.update({"Title":"Electrical Equip"}, {"$set": {"Status":"Not Okay"}})

    if(str(lift)=="Okay"):
        monthly_checklist.update({"Title":"Lift"}, {"$set": {"Status":"Okay"}})
    elif(str(lift)=="Not Okay"):
        monthly_checklist.update({"Title":"Lift"}, {"$set": {"Status":"Not Okay"}})

    if(str(plumbing)=="Okay"):
        monthly_checklist.update({"Title":"Plumbing"}, {"$set": {"Status":"Okay"}})
    elif(str(plumbing)=="Not Okay"):
        monthly_checklist.update({"Title":"Plumbing"}, {"$set": {"Status":"Not Okay"}})

    if(str(cp)=="Okay"):
        monthly_checklist.update({"Title":"Carpentary/Polishing"}, {"$set": {"Status":"Okay"}})
    elif(str(cp)=="Not Okay"):
        monthly_checklist.update({"Title":"Carpentary/Polishing"}, {"$set": {"Status":"Not Okay"}})

    if(str(deepc)=="Okay"):
        monthly_checklist.update({"Title":"Deep Cleaning"}, {"$set": {"Status":"Okay"}})
    elif(str(deepc)=="Not Okay"):
        monthly_checklist.update({"Title":"Deep Cleaning"}, {"$set": {"Status":"Not Okay"}})
    else:
        pass
        
    if(monthly_checklist.find({"Title":"Pest Control"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Pest Control","done":"no","createdAt": datetime.now(),"Category":"Monthly"})
    if(monthly_checklist.find({"Title":"Electrical Equip"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Electrical Equip","done":"no","createdAt": datetime.now(),"Category":"Monthly"})
    if(monthly_checklist.find({"Title":"Lift"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Lift","done":"no","createdAt": datetime.now(),"Category":"Monthly"})
    if(monthly_checklist.find({"Title":"Plumbing"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Plumbing","done":"no","createdAt": datetime.now(),"Category":"Monthly"})
    if(monthly_checklist.find({"Title":"Carpentary/Polishing"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Carpentary/Polishing","done":"no","createdAt": datetime.now(),"Category":"Monthly"})
    if(monthly_checklist.find({"Title":"Deep Cleaning"})[0]["Status"]=="Not Okay"):
        records.insert_one({"Activity":"Deep Cleaning","done":"no","createdAt": datetime.now(),"Category":"Monthly"})
    else:
        pass

    return redirect(url_for('admin'))

@app.route('/admin_task')
def admin_task():
    recordsc = records.find({"done":"no"}).count()
    completesc = records.find({"done":"no"}).count()
    return render_template('admin_task.html',recordsc = recordsc,completesc = completesc)

@app.route('/admin')
def admin():
    
    records_list = records.find({"done":"no"})
    snag_list = snag.find({})
    return render_template('admin.html',records_list = records_list, snag_list=snag_list)

   
@app.route("/done")    
def done ():    
    #Done-or-not ICON    
    id=request.values.get("_id")    
    task=records.find({"_id":ObjectId(id)})    
    if(task[0]["done"]=="yes"):    
        records.update({"_id":ObjectId(id)}, {"$set": {"done":"no"}})
        records.update({"_id":ObjectId(id)}, {"$set": {"createdAt":datetime.now()}})
        records.update({"_id":ObjectId(id)}, {"$unset": {"doneAt":""}})


    else:    
        records.update({"_id":ObjectId(id)}, {"$set": {"done":"yes"}})    
        records.update({"_id":ObjectId(id)}, {"$set": {"doneAt":datetime.now()}})


    redir=redirect_url()        
    return redirect(redir)  

@app.route("/complete.html")
def complete():
    complete_list = records.find({"done":"yes"})
    return render_template("complete.html", complete_list = complete_list)

@app.route("/vendor_list.html")
def vendor_list():
    wrem = washroom_checklist.find({"Status":"None"}).count()
    frem = fridge_checklist.find({"Status":"None"}).count()
    hrem = huddle_checklist.find({"Status":"None"}).count()
    mrem = meeting_checklist.find({"Status":"None"}).count()
    rem = wrem+frem+hrem+mrem

    pantryc = pantry.find({"Quantity":""}).count()
    fnbc = fnb.find({"Quantity":""}).count()
    office_suppc = office_supp.find({"Quantity":""}).count()
    housekeepingc = housekeeping.find({"Quantity":""}).count()
    totalc = pantryc+fnbc+office_suppc+housekeepingc


    monrem = monthly_checklist.find({"Status":"None"}).count()
    taskrem = records.find({"done":"no"}).count()
    return render_template('vendor_list.html', taskrem=taskrem,monrem=monrem,rem=rem,totalc=totalc)

@app.route("/excel", methods=['GET'])
def excel():
    return send_file('F:/Meethi Folder/INTERNSHIPS/WorkAmp/trial01.xlsx')


@app.route("/check_inventory")
def check_inventory():
    pantry_dict = pantry.find({})
    pantry_dict=list(pantry_dict)
    fnb_dict = fnb.find({})
    fnb_dict=list(fnb_dict)
    office_supp_dict = office_supp.find({})
    office_supp_dict = list(office_supp_dict)
    housekeeping_dict = housekeeping.find({})
    housekeeping_dict = list(housekeeping_dict)

    concat = pantry_dict+fnb_dict+office_supp_dict+housekeeping_dict 
    return render_template('check_inventory.html', concat=concat)
   
@app.route("/lead_show")
def lead_show():
    timeframe = "01"
    finance.update_many({},[{"$set":{"month":{"$substr":["$Date of Payment",5,2]}}}])
    selected_list = list(finance.find({"month":timeframe}))
    lst=[]
    for i in selected_list:
        lst.append(int(i["Amount"]))
        
    total = sum(lst)
    avg = round(sum(lst)/len(lst))
    diff = avg - total
    if(diff<=0):
        word = "more"
    elif(diff>0):
        word = "less"
    
    diff = abs(diff)
        
    done = records.find({"done":"yes"}).count()
    total = records.find().count()
    per = round((done/total)*100)
    return render_template('lead.html', per = per, total = total, word = word, diff = diff)


@app.route("/lead", methods=['POST'])
def lead():
    timeframe = request.form.get('timeframe')
    if(timeframe==""):
        fromdate = request.form['from']
        frommonth = fromdate[5:7]
        fromday = int(fromdate[8:10])
        todate = request.form['to']
        tomonth=todate[5:7]
        today = int(todate[8:10])
        finance.update_many({},[{"$set":{"month":{"$substr":["$Date of Payment",5,2]},"date":{"$toInt":{"$substr":["$Date of Payment",8,2]}}}}])
        selected_list = finance.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}]})
        val = selected_list.count()
        lst=[]
        for i in range(0,val):
            lst.append(int(selected_list[i]["Amount"]))
        total = sum(lst)
        val = selected_list.count()
        avg = round(total/val)
        diff = avg - total
        if(diff<=0):
            word = "more"
        elif(diff>0):
            word = "less"
        diff = abs(diff)
        records.update_many({},[{"$set":{"month":{"$substr":["$Date of Task",3,2]},"date":{"$toInt":{"$substr":["$Date of Task",0,2]}}}}])
        recselect_list = records.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}]})
        total1 = recselect_list.count()
        done = records.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"done":"yes"}).count()
        per = (done/total1)*100
        return render_template('lead2.html', total = total, word = word, diff = diff,per=per,frommonth=frommonth,fromday=fromday, today=today,tomonth=tomonth)
    else:
        finance.update_many({},[{"$set":{"month":{"$substr":["$Date of Payment",5,2]}}}])
        selected_list = finance.find({"month":timeframe})
        val = selected_list.count()
        lst=[]
        for i in range(0,val):
            lst.append(int(selected_list[i]["Amount"]))
        total = sum(lst)
        val = selected_list.count()
        avg = round(total/val)
        diff = avg - total
        if(diff<0):
            word = "more"
        elif(diff>0):
            word = "less"
        diff = abs(diff)
        records.update_many({},[{"$set":{"month":{"$substr":["$Date of Task",3,2]}}}])
        select_reclist = records.find({"month":timeframe,"done":"yes"})
        done = select_reclist.count()
        total1 = records.find({"month":timeframe}).count()
        per = round((done/total1)*100)
        return render_template('lead.html', per = per, total = total, word = word, diff = diff, timeframe=timeframe)


@app.route('/expenditure')
def expenditure():
    timeframe = request.values.get("time")
    frommonth = request.values.get("frommonth")
    fromday = request.values.get("fromday")
    tomonth = request.values.get("tomonth")
    today = request.values.get("today")
    if(timeframe!=""):
        electricity = finance.find({"month":timeframe,"Subcategory":"Electricity"})
        vale = electricity.count()
        lstede=[]
        for i in range(0,vale):
            lstede.append(int(electricity[i]["Amount"]))
        totale = sum(lstede)
        vale = electricity.count()
        avge = totale/vale
        timeback = int(timeframe)
        timeback = timeback-1
        timebackstr = "0"+ str(timeback)
        electricity_old = finance.find({"month":timebackstr,"Subcategory":"Electricity"})
        lstedeold=[]
        valeold = electricity_old.count()
        for i in range(0,valeold):
            lstedeold.append(int(electricity_old[i]["Amount"]))
        totaleold = sum(lstedeold)
        salary = finance.find({"month":timeframe,"Subcategory":"Salary"})
        vals = salary.count()
        lsteds=[]
        for i in range(0,vals):
            lsteds.append(int(salary[i]["Amount"]))
        totals = sum(lsteds)
        vals = salary.count()
        avgs = totals/vals
        salary_old = finance.find({"month":timebackstr,"Subcategory":"Salary"})
        lstedsold=[]
        valsold = salary_old.count()
        for i in range(0,valsold):
            lstedsold.append(int(salary[i]["Amount"]))
        totalsold = sum(lstedsold)
        otheri = finance.find({"month":timeframe,"Subcategory":"Other","Category":"Inventory"})
        vali = otheri.count()
        lstedi=[]
        for i in range(0,vali):
            lstedi.append(int(otheri[i]["Amount"]))
        totali = sum(lstedi)
        vali = otheri.count()
        avgi = totali/vali
        otheri_old = finance.find({"month":timebackstr,"Subcategory":"Salary","Category":"Inventory"})
        lstediold=[]
        valiold = otheri_old.count()
        for i in range(0,valiold):
            lstediold.append(int(otheri[i]["Amount"]))
        totaliold = sum(lstediold)
        otherex = finance.find({"month":timeframe,"Subcategory":"Other","Category":"Expense"})
        valex = otherex.count()
        lstedex=[]
        for i in range(0,valex):
            lstedex.append(int(otherex[i]["Amount"]))
        totalex = sum(lstedex)
        valex = otherex.count()
        avgex = totalex/valex
        otherex_old = finance.find({"month":timebackstr,"Subcategory":"Salary","Category":"Expense"})
        lstedexold=[]
        valexold = otherex_old.count()
        for i in range(0,valexold):
            lstedexold.append(int(otherex[i]["Amount"]))
        totalexold = sum(lstedexold)
        total  = totali+totalex
        totalold = totaliold+totalexold
        avg = (avgs+avgex)/2
        inventory = finance.find({"Category":"Inventory"})
        valinv = inventory.count()
        lstedinv=[]
        for i in range(0,valinv):
            lstedinv.append(int(inventory[i]["Amount"]))
        total_inv = sum(lstedinv)
        water = finance.find({ "Subcategory":"Water"})
        valwater = water.count()
        lstedwater=[]
        for i in range(0,valwater):
            lstedwater.append(int(water[i]["Amount"]))
        total_water = sum(lstedwater)
        internet = finance.find({ "Subcategory":"Internet"})
        valint = internet.count()
        lstedint=[]
        for i in range(0,valint):
            lstedint.append(int(internet[i]["Amount"]))
        total_int = sum(lstedint)
        rent = finance.find({ "Subcategory":"Rent"})
        valrent = rent.count()
        lstedrent=[]
        for i in range(0,valrent):
            lstedrent.append(int(rent[i]["Amount"]))
        total_rent = sum(lstedrent)
        return render_template("expenditure.html", totale = totale,avge=avge , totaleold = totaleold, totals = totals, avgs = avgs, totalsold = totalsold,totali=totali, avgi=avgi, totaliold=totaliold, totalex=totalex, totalexold = totalexold, avgex = avgex, total=total, avg=avg, totalold=totalold, total_inv = total_inv, total_water = total_water,total_int = total_int, total_rent = total_rent)
    else:
        electricity = finance.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"Subcategory":"Electricity"})
        vale = electricity.count()
        lstede=[]
        for i in range(0,vale):
            lstede.append(int(electricity[i]["Amount"]))
        totale = sum(lstede)
        vale = electricity.count()
        if(vale==0):
            avge = 0
        else:
            avge = totale/vale
        timebackstr = "03"
        electricity_old = finance.find({"month":timebackstr,"Subcategory":"Electricity"})
        lstedeold=[]
        valeold = electricity_old.count()
        for i in range(0,valeold):
            lstedeold.append(int(electricity_old[i]["Amount"]))
        totaleold = sum(lstedeold)
        salary = finance.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"Subcategory":"Electricity"})
        vals = salary.count()
        lsteds=[]
        for i in range(0,vals):
            lsteds.append(int(salary[i]["Amount"]))
        totals = sum(lsteds)
        vals = salary.count()
        if(vals==0):
            avgs = 0
        else:
            avgs = totals/vals
        salary_old = finance.find({"month":timebackstr,"Subcategory":"Salary"})
        lstedsold=[]
        valsold = salary_old.count()
        for i in range(0,valsold):
            lstedsold.append(int(salary[i]["Amount"]))
        totalsold = sum(lstedsold)
        otheri = finance.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"Subcategory":"Electricity"})
        vali = otheri.count()
        lstedi=[]
        for i in range(0,vali):
            lstedi.append(int(otheri[i]["Amount"]))
        totali = sum(lstedi)
        vali = otheri.count()
        if(vali==0):
            avgi = 0
        else:
            avgi = totali/vali
        otheri_old = finance.find({"month":timebackstr,"Subcategory":"Salary","Category":"Inventory"})
        lstediold=[]
        valiold = otheri_old.count()
        for i in range(0,valiold):
            lstediold.append(int(otheri[i]["Amount"]))
        totaliold = sum(lstediold)
        otherex = finance.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"Subcategory":"Electricity"})
        valex = otherex.count()
        lstedex=[]
        for i in range(0,valex):
            lstedex.append(int(otherex[i]["Amount"]))
        totalex = sum(lstedex)
        valex = otherex.count()
        if(valex==0):
            avgex=0
        else:
            avgex = totalex/valex
        otherex_old = finance.find({"month":timebackstr,"Subcategory":"Salary","Category":"Expense"})
        lstedexold=[]
        valexold = otherex_old.count()
        for i in range(0,valexold):
            lstedexold.append(int(otherex[i]["Amount"]))
        totalexold = sum(lstedexold)
        total  = totali+totalex
        totalold = totaliold+totalexold
        avg = (avgs+avgex)/2
        inventory = finance.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"Category":"Inventory"})
        valinv = inventory.count()
        lstedinv=[]
        for i in range(0,valinv):
            lstedinv.append(int(inventory[i]["Amount"]))
        total_inv = sum(lstedinv)
        water = finance.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"Subcategory":"Water"})
        valwater = water.count()
        lstedwater=[]
        for i in range(0,valwater):
            lstedwater.append(int(water[i]["Amount"]))
        total_water = sum(lstedwater)
        internet = finance.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"Subcategory":"Internet"})
        valint = internet.count()
        lstedint=[]
        for i in range(0,valint):
            lstedint.append(int(internet[i]["Amount"]))
        total_int = sum(lstedint)
        rent = finance.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"Subcategory":"Rent"})
        valrent = rent.count()
        lstedrent=[]
        for i in range(0,valrent):
            lstedrent.append(int(rent[i]["Amount"]))
        total_rent = sum(lstedrent)
        return render_template("expenditure.html", totale = totale,avge=avge , totaleold = totaleold, totals = totals, avgs = avgs, totalsold = totalsold,totali=totali, avgi=avgi, totaliold=totaliold, totalex=totalex, totalexold = totalexold, avgex = avgex, total=total, avg=avg, totalold=totalold, total_inv = total_inv, total_water = total_water,total_int = total_int, total_rent = total_rent)
    
@app.route('/lead_task')
def lead_task():
    timeframe = request.values.get("time")
    if(timeframe!=""):
        done_task = records.find({"month":timeframe,"done":"yes"}).count()
        notdone_task = records.find({"month":timeframe,"done":"no"}).count()
        test = records.aggregate([{"$group":{"_id":"$Category","count":{"$sum":1}}},{"$sort":{"count":-1}},{"$limit":1}])
        max_category = list(test)
        maxtask = max_category[0]["_id"]
        maxcount = max_category[0]["count"]
        testtask = records.aggregate([{"$group":{"_id":"$Activity","count":{"$sum":1}}},{"$sort":{"count":-1}},{"$limit":1}])
        max_act = list(testtask)
        maxact = max_act[0]["_id"]
        maxactcount = max_act[0]["count"]
        return render_template("lead_task.html", done_task = done_task, notdone_task = notdone_task, maxtask=maxtask, maxcount=maxcount,maxact=maxact,maxactcount=maxactcount)
    else:
        frommonth = request.values.get("frommonth")
        fromday = request.values.get("fromday")
        tomonth = request.values.get("tomonth")
        today = request.values.get("today")
        done_task = records.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"done":"yes"}).count()
        notdone_task = records.find({"$or":[{"$and":[{"date":{"$lte":today}},{"month":tomonth}]},{"$and":[{"date":{"$gte":fromday}},{"month":frommonth}]}],"done":"no"}).count()
        test = records.aggregate([{"$group":{"_id":"$Category","count":{"$sum":1}}},{"$sort":{"count":-1}},{"$limit":1}])
        max_category = list(test)
        maxtask = max_category[0]["_id"]
        maxcount = max_category[0]["count"]
        testtask = records.aggregate([{"$group":{"_id":"$Activity","count":{"$sum":1}}},{"$sort":{"count":-1}},{"$limit":1}])
        max_act = list(testtask)
        maxact = max_act[0]["_id"]
        maxactcount = max_act[0]["count"]
        return render_template("lead_task.html", done_task = done_task, notdone_task = notdone_task, maxtask=maxtask, maxcount=maxcount,maxact=maxact,maxactcount=maxactcount)

@app.route('/signout')
def signout():
    return redirect(url_for('login'))
    

@app.route("/expense_inv")
def expense_inv():
    inv_list = finance.find({"Category":"Inventory"})
    return render_template('expense_inv.html', inv_list = inv_list)

@app.route("/expense_water")
def expense_water():
    water_list = finance.find({"Subcategory":"Water"})
    return render_template('expense_water.html', water_list = water_list)
@app.route("/expense_int")
def expense_int():
    int_list = finance.find({"Subcategory":"Internet"})
    return render_template('expense_int.html', int_list = int_list)
@app.route("/expense_rent")
def expense_rent():
    rent_list = finance.find({"Subcategory":"Rent"})
    return render_template('expense_rent.html', rent_list = rent_list)

@app.route('/add_task')
def add_task():
    recordsc = records.find({"done":"no"}).count()
    completesc = records.find({"done":"no"}).count()
    return render_template('add_task.html',recordsc = recordsc,completesc = completesc)

@app.route('/add_task_func',methods=['POST'])
def add_task_func():
    records.insert_one({"Activity":request.form['task'],"Category":request.form.get('category'),"Date of Task":request.form['date'],"done":"no"})
    redir=redirect_url()
    return redirect(redir)


if __name__=='__main__':

    app.run(host='0.0.0.0', port=8080, debug=True)


