import requests

OWapikey="add your own key"
uin=input("select city: ")

owresp=requests.get(f"http://api.openweathermap.org/data/2.5/weather?q={uin}&units=imperial&appid={OWapikey}")

if owresp.json()['cod'] =='404':
    print("No City found")
else:
    jresp=owresp.json()
    weather=jresp['weather'][0]['main']
    temp=round(jresp['main']['temp'])
    print(f"Weather in {uin} is: {weather}")
    print(f"Temperature in {uin} is: {temp}F")
