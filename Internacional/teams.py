from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

def read_teams():
  client_id = '83652916-8a3f-4624-a04c-f5b9054d530d'
  client_secret = 'Tribilin123.'
  url = 'https://agrosuper.sharepoint.com/sites/PlanesdeventaVVII'
  name_folder = 'Colaboración PV - Proyección'
  week = 'Semana 04'
  file_1 = 'Distribución Internacional - Terrestres.xlsx'
  relative_url = f'/Shared Documents/General/{name_folder}/{week}/{file_1}'
  print(url, relative_url)
  print(relative_url)
  path = 'https://agrosuper.sharepoint.com/:x:/r/sites/PlanesdeventaVVII/Shared%20Documents/General/Colaboraci%C3%B3n%20PV%20-%20Proyecci%C3%B3n/Semana%2004/Distribuci%C3%B3n%20Internacional%20-%20Terrestres.xlsx?d=wbdc399ae4be54a5a90c24786c133b081&csf=1&web=1&e=ovdeWp'

  ctx_auth = AuthenticationContext(url)
  if ctx_auth.acquire_token_for_app(client_id, client_secret):
    ctx = ClientContext(url, ctx_auth)
    with open(filename, 'wb') as output_file:
      response = File.open_binary(ctx, relative_url)
      output_file.write(response.content) 
  else:
    print(ctx_auth.get_last_error())