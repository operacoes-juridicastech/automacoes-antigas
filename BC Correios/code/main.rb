require 'rubyXL'
require 'httparty'
require 'ox'

def main
  path = Dir.pwd + "\\unidades_postais.xlsx"
  workbook = RubyXL::Parser.parse path

  sheet = workbook[0]
  units = []
    
  (1..1000).each do |i|
  break if sheet[i][0].value == 'FIM'
    unit = { name: sheet[i][0].value, 
             unit: sheet[i][1].value,
             dependency: sheet[i][2].value,
             operator: sheet[i][3].value,
             password: sheet[i][4].value }
    units.append unit
  end

  units.each do |unit|
    unidade = unit[:unit] 
    username = "#{unidade < 10000 ? "0" << unidade.to_s : unidade}#{unit[:dependency]}.#{unit[:operator]}"
    unit[:username] = username
    response = post unit, 'ConsultarPastasAutorizadas'
    
    puts "Username: #{username}\nPassword: #{unit[:password]}"
    
    xml_response = Nokogiri::XML response.body
    teste =  xml_response.xpath '//Nome'
    p teste
    break
  end

rescue => exception
  puts "ERROR: " + exception.message
end

def post(params, op)
  uri = 'https://bccorreio.bcb.gov.br/bccorreiows/CorreioWS.asmx'
  auth = { username: params[:username], password: params[:password] }

  header = { 'Host': 'bccorreio.bcb.gov.br',
             'Content-Type': 'text/xml',
             'Content-Length': '100',
             'SOAPAction': 'http://www.bcb.gov.br/correiows/' << op }

  body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + create_xml(params, op)

  result = HTTParty.post(uri, basic_auth: auth, headers: header, body: body)
  result
end

def create_xml(params, op)
  case op 
    when 'ConsultarPastasAutorizadas'
      op_1
    else
      puts 'No function defined for this type of XML.'
      return nil
  end
end

def op_1 #Consultar pastas autorizadas.
  document = Ox::Document.new(version: '1.0')
  
  envelope = Ox::Element.new :'soap:Envelope'
  envelope[:'xmlns:xsi']  = "http://www.w3.org/2001/XMLSchema-instance"
  envelope[:'xmlns:xsd']  = "http://www.w3.org/2001/XMLSchema"
  envelope[:'xmlns:soap'] = "http://schemas.xmlsoap.org/soap/envelope/"

  body = Ox::Element.new :'soap:Body'

  operation =Ox::Element.new :'ConsultarPastasAutorizadas'
  operation[:xmlns] = "http://www.bcb.gov.br/correiows"

  document << envelope
  envelope << body
  body << operation

  xml = Ox.dump(document, {})
  return xml
end


def op_2 #Consultar correios por pasta.
  document = Ox::Document.new(version: '1.0')

  envelope = Ox::Element.new :'soap:Envelope'
  envelope[:'xmlns:xsi']  = "http://www.w3.org/2001/XMLSchema-instance"
  envelope[:'xmlns:xsd']  = "http://www.w3.org/2001/XMLSchema"
  envelope[:'xmlns:soap'] = "http://schemas.xmlsoap.org/soap/envelope/"
  document << envelope

  body = Ox::Element.new :'soap:Body'
  envelope << body

  operation = Ox::Element.new :ConsultarCorreiosPorPasta
  operation[:xmlns] = "http://www.bcb.gov.br/correiows"
  body << operation

  consulta = Ox::Element.new :consulta
  operation << consulta

  pasta = Ox::Element.new :Pasta
  consulta << pasta

  unidade = Ox::Element.new :Unidade
  pasta << unidade

  nome = Ox::Element.new :Nome
  unidade << nome

  ativa = Ox::Element.new :Ativa
  unidade << ativa

  tipo = Ox::Element.new :Tipo
  unidade << tipo

  dependencia = Ox::Element.new :Dependencia
  dependencia.replace_text 'Teste'
  pasta << dependencia

  setor = Ox::Element.new :Setor
  pasta << setor

  tipo = Ox::Element.new :Tipo
  pasta << tipo

  xml = Ox.dump(document, {})
  return xml
end

main