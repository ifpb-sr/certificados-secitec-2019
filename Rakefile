require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'fileutils'
require 'yaml'
require 'byebug'
require 'csv'
require 'set'
require 'uri'

OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'.freeze
APPLICATION_NAME = 'Google Sheets API Ruby Quickstart'.freeze
CREDENTIALS_PATH = 'credentials.json'.freeze
# The file token.yaml stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
TOKEN_PATH = 'token.yaml'.freeze
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS_READONLY

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize
  client_id = Google::Auth::ClientId.from_file(CREDENTIALS_PATH)
  token_store = Google::Auth::Stores::FileTokenStore.new(file: TOKEN_PATH)
  authorizer = Google::Auth::UserAuthorizer.new(client_id, SCOPE, token_store)
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil?
    url = authorizer.get_authorization_url(base_url: OOB_URI)
    puts 'Open the following URL in the browser and enter the ' \
         "resulting code after authorization:\n" + url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI
    )
  end
  credentials
end


# atividades = ['SEGUNDA1930', 'TERCA0830', 'TERCA1000-1', 'TERCA1000-2', 'TERCA1000-3', 'TERCA1000-4', 'TERCA1330-1', 'TERCA1330-2', 'TERCA1330-3', 'TERCA1330-4', 'QUARTA0830', 'QUARTA1000-1', 'QUARTA1000-2', 'QUARTA1000-3', 'QUARTA1000-4', 'QUARTA1330-1', 'QUARTA1330-2', 'QUARTA1330-3', 'QUARTA1330-4', 'QUINTA0830', 'QUINTA1000-1', 'QUINTA1000-2', 'QUINTA1000-3', 'QUINTA1000-4', 'QUINTA1330-1', 'QUINTA1330-2', 'QUINTA1330-3', 'QUINTA1330-4', 'SEXTA0800']


atividades = YAML.load_file('atividades.yml')

desc 'Gera os csv para mesclar com os certificados'
task 'csv' do
# Initialize the API
  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize

  # Prints the names and majors of students in a sample spreadsheet:
  # https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
  spreadsheet_id = '1v1OpBmgd-ws5EbSQ9PahMKsfXMtOJWrAubtu5m2oZvg'
  range = 'TERCA1330_1'
  response = service.get_spreadsheet_values(spreadsheet_id, range)
  puts 'No data found.' if response.values.empty?
  response.values.each do |row|
    # Print columns A and E, which correspond to indices 0 and 4.
    puts "#{row[0]},#{row[1]}"
  end

end




def gera_certificado(nome, modelo, titulo, data, chave)
  return if modelo == 'ApresentacaoOral'
  #byebug
  cvs_dir = "tmp/csv/#{nome}/"
  pdf_dir = "tmp/pdf/#{nome}/"
  FileUtils.mkdir_p(cvs_dir) unless File.directory?(cvs_dir)
  FileUtils.mkdir_p(pdf_dir) unless File.directory?(pdf_dir)
  csvfile = "tmp/csv/#{nome}/#{chave}.csv"
  pdffile = "tmp/pdf/#{nome}/#{chave}.pdf"
  modelofile = "modelos/digital/#{modelo}.svg"
  CSV.open(csvfile, "wb") do |csv|
    csv << ["nome", "titulo", "data"]
    csv << [nome, titulo, data]
  end

  sh "inkscape_merge -f #{modelofile} -d '#{csvfile}' -o '#{pdffile}'"

end

# Registra presença baseado no horário
# se já registrou naquele horário, não
# pode registrar em outro.
def registra_presenca(nome,chave,presenca)
  # assert presenca[nome] é um Set
  # o substring [0,9] mantém único os eventos
  # que ocorreram simultaneos
  # como estamos usando um set, conta apenas um.
  presenca[nome] << chave[0,9]
end

desc "Gera os certificados dos ouvintes a partir do google docs"
task :ouvintes do
  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize

  # Prints the names and majors of students in a sample spreadsheet:
  # https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
  spreadsheet_id = '1oo3zJuUMsoT8Mjde7ekgaGTTK2nyO2KBKkQWTmtQvg8'
  range = 'Página1'
  response = service.get_spreadsheet_values(spreadsheet_id, range)
  puts 'No data found.' if response.values.empty?
  presenca = {}
  response.values.each do |row|
    # Print columns A and E, which correspond to indices 0 and 4.
    puts "Gerando certificados do aluno: #{row[0]}"

    nome = row[0]
    presenca[nome] = Set[]

    atividades.each do |atividade|
      chave = atividade[0]
      titulo = atividade[1]['t']
      i = atividade[1]['i']
      modelo = atividade[1]['m']
      data = atividade[1]['d']
      presente = row[i]=="1"
      #puts "#{chave},#{presente}"


      if presente
        # gera_certificado(nome, modelo, titulo, data, chave)
        registra_presenca(nome,chave,presenca)
      end

    end

  end
  open('tmp/certificados.md', 'w') do |f|
    f << "# Certificados digitais da SECITEC 2018 - IFPB Santa Rita\n\n"
    f << "Clique no seu nome para baixar seus certificados. \n\n -  [Solicitar correção de certificado](https://goo.gl/forms/zG4fxGAuoiKGdmLD2)\n- [Justificar ausência](https://goo.gl/forms/hFkSZgS5rsYthJ5J2).\n\n"
    presenca.delete('Nome')
    presenca.each do |pessoa|
      nome = pessoa[0]
      participacao = pessoa[1].size*100/10
      url = URI.escape("https://github.com/ifpb-sr/certificados-secitec-2018/tree/master/docs/pdf/"+nome)
      f << "## [#{nome}](#{url})\n\n"
      f << "Participação: #{participacao}%\n\n"
      f << "Ausências: (em processamento)\n\n"
    end
  end
end

# certificado fechine
