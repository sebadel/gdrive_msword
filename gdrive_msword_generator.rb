# Quote Generator - 2013
# Author: Sebastien Delneste (seb@bigup.be)
# Connects to a GDrive service, copies a template to a file, calls an Apps Script to
# alter contents of the file and finally exports it to MS Word format or PDF

require 'rubygems'
require 'google/api_client'
require 'launchy'

SERVICE_ACCOUNT_EMAIL = '1021070011307@developer.gserviceaccount.com'
SERVICE_ACCOUNT_PKCS12_PATH = './'
SERVICE_ACCOUNT_PKCS12_FILE = SERVICE_ACCOUNT_PKCS12_PATH + 'df61217b2778a409e42dcd18f0215a00b278e9ec-privatekey.p12'

# OAuth
CLIENT_ID = '532012347106-15sl2ci6v8esssr3g4bn7n7544qd4ina.apps.googleusercontent.com'
CLIENT_SECRET = '_BIG_SECRET_'
OAUTH_SCOPE = 'https://www.googleapis.com/auth/drive'
REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob'


QUOTE_TEMPLATE_FILE_ID = '1bY-kaebCvHC-X1Wuy7yAjSog-s9bQevQT0twUSQYovM'


def DriveServiceAccountLogin()
  ##
  # Build a Drive client instance authorized with the service account.
  #
  # @return [Google::APIClient]
  #   Client instance
  key = Google::APIClient::PKCS12.load_key(SERVICE_ACCOUNT_PKCS12_FILE, 'notasecret')
  asserter = Google::APIClient::JWTAsserter.new(SERVICE_ACCOUNT_EMAIL,
        'https://www.googleapis.com/auth/drive', key)
  client = Google::APIClient.new
  client.authorization = asserter.authorize()
  client
end

def DriveOAuthLogin()
  # Create a new API client & load the Google Drive API 
  client = Google::APIClient.new
  # Request authorization
  client.authorization.client_id = CLIENT_ID
  client.authorization.client_secret = CLIENT_SECRET
  client.authorization.scope = OAUTH_SCOPE
  client.authorization.redirect_uri = REDIRECT_URI
  uri = client.authorization.authorization_uri
  Launchy.open(uri)
  # Exchange authorization code for access token
  $stdout.write  "Enter authorization code: "
  client.authorization.code = gets.chomp
  client.authorization.fetch_access_token!
  client
end

def DriveListFiles(client, drive)
  result = Array.new
  page_token = nil
  begin
    parameters = {}
    if page_token.to_s != ''
      parameters['pageToken'] = page_token
    end
    api_result = client.execute(
      :api_method => drive.files.list,
      :parameters => parameters)
    if api_result.status == 200
      files = api_result.data
      result.concat(files.items)
      page_token = files.next_page_token
    else
      puts "An error occurred: #{result.data['error']['message']}"
      page_token = nil
    end
  end while page_token.to_s != ''
  # result
  result.each{|file| print "ID: #{file.id} - Title: #{file.title}"}
end

def DriveGetFileById(client, drive, file_id)
  result = client.execute(
    :api_method => drive.files.get,
    :parameters => { 'fileId' => file_id })
  
  if result.status == 200
    file = result.data
    return file
  else
    puts "An error occurred: #{result.data['error']['message']}"
  end
end

def DriveDownloadFile(client, file, local_filename, mime_type='application/pdf')
  print "Downloading file ... "
  download_url = file['exportLinks'][mime_type]	
  if download_url
    result = client.execute(:uri => download_url)
    if result.status == 200
      p result.body
      File.open(local_filename, 'w') {|f| f.write(result.body) }	
      return result.body
    else
      puts "An error occurred: #{result.data['error']['message']}"
      return nil
    end
  else
    puts "No file.download_url found: #{download_url}"
    # The file doesn't have any content stored on Drive.
    return nil
  end
end

def DriveInsertFile(drive, client, title, description, mime_type='text/plain')
  # Insert a file
  print "Inserting file ... "
  file = drive.files.insert.request_schema.new({
    'title' => title,
    'description' => description,
    'mimeType' => mime_type
  })
  media = Google::APIClient::UploadIO.new('document.txt', 'text/plain')
  result = client.execute(
	:api_method => drive.files.insert,
	:body_object => file,
	:media => media,
	:parameters => {
	  'uploadType' => 'multipart',
	  'alt' => 'json'})
  return file
end

def DriveCopyFile(drive, client, origin_file_id)
  copied_file = drive.files.copy.request_schema.new({
    'title' => copy_title
  })
  result = client.execute(
    :api_method => drive.files.copy,
    :body_object => copied_file,
    :parameters => { 'fileId' => origin_file_id })
  if result.status == 200
    return result.data
  else
    puts "An error occurred: #{result.data['error']['message']}"
  end
  return copied_file
end

def AdjustQuote(client, document_id, json_file)
  web_app_uri = 'https://script.google.com/macros/s/AKfycbym23u4nDoQuCbABZvca2RsY6HKBjh-wfhsOymM_s6nBa_cDmo/exec'
  web_app_parameters = {
    'document_id' => '1BAyDObdSnS7XjgepyQwhyirF9VZi-C5R1kr-x2mYj5k',
    'json_file' => 'file.json',
    'secret_token' => '_JuStBetterToKeepTHisPRIVateBECaUsETHEWEBAppMu$Tb3Publ!C'
  } 
  CallAppsScript(client, web_app_uri, web_app_parameters)
end

def CallAppsScript(client, script_uri, parameters)
  uri = script_uri + '?' + parameters.map{|k,v| "#{k}=#{v}"}.join('&')
  result = client.execute(:uri => URI.escape(uri))
  print "URI: #{uri}"
  print "Result: #{result}" 
  print "Result: #{result.body}" 
end


client = DriveServiceAccountLogin()
drive = client.discovered_api('drive', 'v2')
# Create JSON file
# file = DriveCopyFile(drive, client, QUOTE_TEMPLATE_FILE_ID)
AdjustQuote(client, 'document_id', 'json_file')

# ms_word_file = DriveDownloadFile(client, file, 'MyDoc.doc',
#   'application/vnd.openxmlformats-officedocument.wordprocessingml.document')


# file = DriveInsertFile(drive, client, 'My Doc', 'yop yop yop')
# DriveListFiles(client, drive)
# file = DriveGetFileById(client, drive, '1Iek6jDr3au1w1iXmNbWhB4LDkUo9gfcN34zmXVJOqWg')
    

