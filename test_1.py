from fyers_api import accessToken

from fyers_api import fyersModel

app_id = "V"

app_secret = ""

app_session = accessToken.SessionModel(app_id, app_secret)

response = app_session.auth()

print(response)
