from flask import Flask,render_template,request,abort
# import json to load json data to python dictionary
import json
# urllib.request to make a request to api
import urllib.request

from openpyxl import Workbook
wb = Workbook()

# grab the active worksheet
ws1 = wb.active
ws2 = wb.active

i=0
j=0

app = Flask(__name__)
def tocelcius(temp):
    return str(round(float(temp) - 273.16,2))

@app.route('/',methods=['POST','GET'])
def weather():
    if request.method == 'POST':
        city = request.form['city']
    else:
        #for default name mathura
        if j>=7:j=0;
        city = ['mathura','Washington','Franklin','Arlington','Centerville','Lebanon','Clinton','Springfield']
            
    try:
        source = urllib.request.urlopen('http://api.openweathermap.org/data/2.5/weather?q=' + city + '&appid=48a90ac42caa09f90dcaeee4096b9e53').read()
    except:
        return abort(404)
    # converting json data to dictionary

    list_of_data = json.loads(source)

    # data for variable list_of_data
    data = {
        "country_code": str(list_of_data['sys']['country']),
        "coordinate": str(list_of_data['coord']['lon']) + ' ' + str(list_of_data['coord']['lat']),
        "temp": str(list_of_data['main']['temp']) + 'k',
        "temp_unit": tocelcius(list_of_data['temperature']['unit']),
        "pressure": str(list_of_data['main']['pressure']),
        "humidity": str(list_of_data['main']['humidity']),
        "cityname":str(city[j]),
        "update":int(list_of_data['lastupdate']['value']),
    }
    if request.method == 'POST':
        if request.form['C'] == 'Do Something':
            return (input(float(list_of_data['main']['temp'])
        elif request.form['F'] == 'Do Something Else':
            return ((float(list_of_data['main']['temp']) * 9/5) + 32)
        elif request.form['Start/Stop']=='0'
            return (data.update=0)
        
    ws1.write(i, 0, data.cityname) 
    ws1.write(i, 1, str(list_of_data['main']['temp'])) 
    ws1.write(i, 2, tocelcius(list_of_data['temperature']['unit'])) 
    ws1.write(i, 3, data.update) 
    
    ws2.write(i, 2, str(list_of_data['sys']['country'])) 
    ws2.write(i, 3, IsModified(int(list_of_data['lastupdate']['value'])) 
    
    
    # Save the file
    wb.save("Project Sample.xlsx")
    i=i+1;
    return render_template('index.html',data=data)



if __name__ == '__main__':
    app.run(debug=True)