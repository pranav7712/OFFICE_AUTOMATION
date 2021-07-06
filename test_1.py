from fyers_api import accessToken

from fyers_api import fyersModel

app_id = "1XVY7MKYZV"

app_secret = "L981CZ96V7"

app_session = accessToken.SessionModel(app_id, app_secret)

response = app_session.auth()

print(response)
