import json
import requests
headers = {
    "Authorization":"Bearer ya29.a0AVvZVsqXb-ltLAtCXI5CdzSsoGQt3Vqig1hmFtYbMccVh3JoGWM6ILLZA3omQPedAZV5XNBw9id-5a3cyM9SktIj6Iqv9w-ADgCQB0rHJ9IzmU8utp3wV-RxWC1wcfIQuUYReFKPN7TwN6qE7SeVpG0sN9q_aCgYKAcoSARASFQGbdwaIYtvJsU2y_sacbDX5ZKG6vg0163"
}
 
para = {
    "name":"stocks_data.xlsx",
    "parents":["1uXwXat5YKMD71koxv8Ox6dPRkUT9hGun"]
}
 
files = {
    'data':('metadata',json.dumps(para),'application/json;charset=UTF-8'),
    'file':open('./stocks_data.xlsx','rb')
}
 
r = requests.post("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
    headers=headers,
    files=files
)