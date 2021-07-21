require 'httparty'
require 'date'
require 'awesome_print'
require 'rubyXL'

#require 'base64'
#require 'prawn'


class Crawler

    def start
        workbook = RubyXL::Parser.parse "unidades_postais.xlsx"

        sheet = workbook[0]
        units = []
      
        (1..1000).each do |i|
            break if sheet[i][0].value == 'FIM'
              unit = { 
                name: sheet[i][0].value, 
                unit: sheet[i][1].value,
                dependency: sheet[i][2].value,
                operator: sheet[i][3].value,
                password: sheet[i][4].value 
              }
            unit[:unit] = "0#{unit[:unit]}" if unit[:unit].to_s.length < 5            
            units.append unit
        end

        units.each do |unit| 
            get_all_mail unit 
            #break
        end
    end

    def get_all_mail(data)

        puts "\n\nChecando caixa de correio da unidade #{data[:unit]}"

        data[:operation] = 'ConsultarCorreiosPorPasta'
        correios = {}

        user = "#{data[:unit]}#{data[:dependency]}.#{data[:operator]}" 
        user = "0" + user if user.length < 20

        data[:username] = user
        data[:page] = '0'
        
        response = post(data)
        unless response.code == 400 or response.body.include? 'DOCTYPE'
            hash = response.parsed_response
            puts hash["Envelope"]["Body"]["ConsultarCorreiosPorPastaResponse"]["ConsultarCorreiosPorPastaResult"]["QuantidadeCorreios"]
        else
            puts "Erro na unidade #{data[:unit]}"
        end
        return

        emails = []

        emails.append(hash["Envelope"]["Body"]["ConsultarCorreiosPorPastaResponse"]["ConsultarCorreiosPorPastaResult"]["Correios"]["ResumoCorreioWSDTO"])

        if response.code == 200
            hash = response.parsed_response
            #ap hash, options={}
            paginas = hash["Envelope"]["Body"]["ConsultarCorreiosPorPastaResponse"]["ConsultarCorreiosPorPastaResult"]["QuantidadeCorreios"]
            paginas = (paginas.to_i / 10.0).ceil
            (1...paginas).each do |i|
                data[:page] = i
                hash = post(data).parsed_response
                emails.append(hash["Envelope"]["Body"]["ConsultarCorreiosPorPastaResponse"]["ConsultarCorreiosPorPastaResult"]["Correios"]["ResumoCorreioWSDTO"])
                #ap hash["Envelope"]["Body"]["ConsultarCorreiosPorPastaResponse"]["ConsultarCorreiosPorPastaResult"]["Correios"]["ResumoCorreioWSDTO"], options = {}
                #ap hash, options = {}
            end
        end

        
        #ap mail.flatten, options = {}
        emails.flatten.each do |email|
            data[:operation] = 'LerCorreio'

            data[:numero_correio] = email["NumeroCorreio"]
            data[:data_correio] = email["Data"]
            data[:unidade_dest] = email["UnidadeDestinataria"]
            data[:unidade_rem] = email["UnidadeRemetente"]
            data[:transicao] = email["Transicao"]
            data[:assunto] = email["Assunto"]
            data[:status] = email["Status"]
            data[:versao] = email["Versao"]
            data[:grupo] = email["grupo"]

            response = post(data).parsed_response
            puts "\n\n\n\n\n\n\n\n #{response["Envelope"]["Body"]["LerCorreioResponse"]["LerCorreioResult"]["DetalheCorreio"]["Conteudo"]}"
        end
    end 

    def post(data)
      uri = 'https://bccorreio.bcb.gov.br/bccorreiows/CorreioWS.asmx'
      auth = { username: data[:username], password: data[:password] }

      header = get_header 10, data[:operation]
      body = build_xml data

      response = HTTParty.post(uri, basic_auth: auth, headers: header, body: body)
      response
    end

    def get_header(length, operation)
      {
        "Host": "bccorreio.bcb.gov.br",
        "Content-Type": "text/xml",
        "Content-Length": "#{length}",
        "SOAPAction": "http://www.bcb.gov.br/correiows/#{operation}"
      }
    end

    def get_yesterday
        "2019-06-07"
    end

    def get_dates
      today = Time.now.strftime "%Y-%m-%d"

      year = Time.now.strftime("%Y").to_i
      month = Time.now.strftime("%m").to_i
      day = Time.now.strftime("%d").to_i

      if Date.new(year, month, day).monday?
        yesterday = Time.at(Time.now.to_i - 86400 * 3).strftime "%Y-%m-%d"
      else
        yesterday = Time.at(Time.now.to_i - 86400).strftime "%Y-%m-%d"
      end
      { today: today, yesterday: yesterday }
    end

    def build_xml(data)
      case data[:operation]
        when 'ConsultarPastasAutorizadas'
          consultar_pastas_autorizadas
        when 'ConsultarCorreiosPorPasta'
          consultar_correios_por_pasta data
        when 'LerCorreio'
          ler_correio data
        when 'ConsultarComunicacaoGeralDocumentoDivulgacao'
          consultar_comunicacao_geral
        when 'ObterAnexo'
          obter_anexo
        else
          raise "Erro na construcao do XML."
      end
    end

    def consultar_pastas_autorizadas
      "<?xml version=\"1.0\" encoding=\"utf-8\"?>"\
          "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"\
              "<soap:Body>"\
                  "<ConsultarPastasAutorizadas xmlns=\"http://www.bcb.gov.br/correiows\"/>"\
              "</soap:Body>"\
          "</soap:Envelope>"
    end

    def consultar_correios_por_pasta(data)
      "<?xml version=\"1.0\" encoding=\"utf-8\"?>"\
      "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"\
          "<soap:Body>"\
              "<ConsultarCorreiosPorPasta xmlns=\"http://www.bcb.gov.br/correiows\">"\
                  "<consulta>"\
                      "<Pasta>"\
                          "<Unidade>"\
                              "<Nome>#{data[:unit]}</Nome>"\
                              "<Ativa>true</Ativa>"\
                              "<Tipo>InstituicaoFinanceira</Tipo>"\
                          "</Unidade>"\
                          "<Dependencia></Dependencia>"\
                          "<Setor/>"\
                          "<Tipo>CaixaEntrada</Tipo>"\
                      "</Pasta>"\
                      "<DataInicial>#{get_yesterday}T06:00:01</DataInicial>"\
                      "<DataFinal>#{get_yesterday}T23:59:59</DataFinal>"\
                      "<ApenasMensagens>true</ApenasMensagens>"\
                      "<PesquisarEmTodasAsPastas>false</PesquisarEmTodasAsPastas>"\
                      "<Pagina>#{data[:page]}</Pagina>"\
                  "</consulta>"\
              "</ConsultarCorreiosPorPasta>"\
          "</soap:Body>"\
      "</soap:Envelope>"
    end

    def consultar_comunicacao_geral   
      puts 'Consultar comunicação geral.'
      "<?xml version=\"1.0\" encoding=\"utf-8\"?>"\
        "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"\
          "<soap:Body>"\
            "<ConsultarComunicacaoGeralDocumentoDivulgacao xmlns=\"http://www.bcb.gov.br/correiows\">"\
              "<parametros>"\
                "<Pagina>0</Pagina>"\
                "<NumeroDocumento></NumeroDocumento>"\
                "<TipoDocumento>MENSAGEM</TipoDocumento>"\
                "<ExpressaoBusca></ExpressaoBusca>"\
                "<DataInicial>2018-06-06T06:00:01</DataInicial>"\
                "<DataFinal>2019-06-06T23:59:59</DataFinal>"\
                "<PesquisarPalavrasChave>false</PesquisarPalavrasChave>"\
              "</parametros>"\
            "</ConsultarComunicacaoGeralDocumentoDivulgacao>"\
          "</soap:Body>"\
        "</soap:Envelope>"
    end
    
    def ler_correio(data)
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"\
        "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"\
            "<soap:Body>"\
                "<LerCorreio xmlns=\"http://www.bcb.gov.br/correiows\">"\
                    "<parametros>"\
                        "<Correio>"\
                            "<Assunto>#{data[:assunto]}</Assunto>"\
                            "<Data>#{data[:data_correio]}</Data>"\
                            "<UnidadeDestinataria>#{data[:unidade_dest]}</UnidadeDestinataria>"\
                            "<DependenciaDestinataria></DependenciaDestinataria>"\
                            "<UnidadeRemetente>#{data[:unidade_rem]}</UnidadeRemetente>"\
                            "<DependenciaRemetente></DependenciaRemetente>"\
                            "<Grupo>#{data[:grupo]}</Grupo>"\
                            "<Status>#{data[:status]}</Status>"\
                            "<TipoCorreio>MENSAGEM</TipoCorreio>"\
                            "<NumeroCorreio>#{data[:numero_correio]}</NumeroCorreio>"\
                            "<Pasta>"\
                                "<Unidade xsi:nil=\"true\" />"\
                                "<Dependencia>#{data[:dependency]}</Dependencia>"\
                                "<Setor xsi:nil=\"true\" />"\
                                "<Tipo>CaixaEntrada</Tipo>"\
                            "</Pasta>"\
                            "<Transicao>#{data[:transicao]}</Transicao>"\
                            "<Versao>#{data[:versao]}</Versao>"\
                        "</Correio>"\
                    "</parametros>"\
                "</LerCorreio>"\
            "</soap:Body>"\
        "</soap:Envelope>"  
    end

    def ler_correio_1
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"\
        "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"\
            "<soap:Body>"\
                "<LerCorreio xmlns=\"http://www.bcb.gov.br/correiows\">"\
                    "<parametros>"\
                        "<Correio>"\
                            "<Assunto>SOLCCS 2019055547</Assunto>"\
                            "<Data>2019-06-06T18:09:42.51</Data>"\
                            "<UnidadeDestinataria>23181</UnidadeDestinataria>"\
                            "<DependenciaDestinataria></DependenciaDestinataria>"\
                            "<UnidadeRemetente>DEATI</UnidadeRemetente>"\
                            "<DependenciaRemetente></DependenciaRemetente>"\
                            "<Grupo>F1 </Grupo>"\
                            "<Status>Lido/Recebido</Status>"\
                            "<TipoCorreio>MENSAGEM</TipoCorreio>"\
                            "<NumeroCorreio>119045371</NumeroCorreio>"\
                            "<Pasta>"\
                                "<Unidade xsi:nil=\"true\" />"\
                                "<Dependencia>0001</Dependencia>"\
                                "<Setor xsi:nil=\"true\" />"\
                                "<Tipo>CaixaEntrada</Tipo>"\
                            "</Pasta>"\
                            "<Transicao>17287805</Transicao>"\
                            "<Versao>550</Versao>"\
                        "</Correio>"\
                    "</parametros>"\
                "</LerCorreio>"\
            "</soap:Body>"\
        "</soap:Envelope>"
    end

    def obter_anexo
  
      "<?xml version=\"1.0\" encoding=\"utf-8\"?>"\
      "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"\
        "<soap:Body>"\
          "<ObterAnexo xmlns=\"http://www.bcb.gov.br/correiows\">"\
            "<parametros>"\
              "<NumeroCorreio>119045371</NumeroCorreio>"\
              "<Versao>550</Versao>"\
              "<Transicao>17287805</Transicao>"\
              "<Pasta>"\
                "<Unidade>"\
                  "<Nome>23181</Nome>"\
                  "<Ativa>true</Ativa>"\
                  "<Tipo>InstituicaoFinanceira</Tipo>"\
                "</Unidade>"\
                "<Dependencia>0001</Dependencia>"\
                "<Tipo>CaixaEntrada</Tipo>"\
              "</Pasta>"\
              "<Anexo>"\
                "<IdAnexo>198792</IdAnexo>"\
                "<NomeAnexo>-2019-054854-77006--04062019180759.PDF</NomeAnexo>"\
              "</Anexo>"\
            "</parametros>"\
          "</ObterAnexo>"\
        "</soap:Body>"\
      "</soap:Envelope>"

    end

    def obter_anexo_2

      "<?xml version=\"1.0\" encoding=\"utf-8\"?>"\
      "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"\
        "<soap:Body>"\
          "<ObterAnexo xmlns=\"http://www.bcb.gov.br/correiows\">"\
            "<parametros>"\
              "<NumeroCorreio>int</NumeroCorreio>"\
              "<Versao>int</Versao>"\
              "<Transicao>int</Transicao>"\
              "<Pasta>"\
                "<Unidade>"\
                  "<Nome>string</Nome>"\
                  "<Ativa>boolean</Ativa>"\
                  "<Tipo>UnidadeBanco or UnidadeExterna or InstituicaoFinanceira</Tipo>"\
                "</Unidade>"\
                "<Dependencia>string</Dependencia>"\
                "<Setor/>"\
                "<Tipo>CaixaEntrada</Tipo>"\
              "</Pasta>"\
              "<Anexo>"\
                "<IdAnexo>int</IdAnexo>"\
                "<NomeAnexo>string</NomeAnexo>"\
                "<Conteudo>base64Binary</Conteudo>"\
              "</Anexo>"\
            "</parametros>"\
          "</ObterAnexo>"\
        "</soap:Body>"\
      "</soap:Envelope>"

    end
  
end

c = Crawler.new
c.start

#data = Hash.new('')

#data[:operation] = 'ConsultarComunicacaoGeralDocumentoDivulgacao'
#data[:operation] = 'ConsultarPastasAutorizadas'

#data[:username] = "231810001.S-WJUD1501"
#data[:password] = "BJUD2019"
#
#data[:unidade_nome] = '23181'
#data[:unidade_tipo] = 'InstituicaoFinanceira'
#data[:dependencia] = '0001'
#data[:tipo] = 'CaixaEntrada'
#data[:pagina] = '1'
#
#
#data[:operation] = 'ConsultarCorreiosPorPasta'
#response = c.post(data)
#
#unless response.body.include? 'DOCTYPE'
#  puts "\n\n----------------------------------- CONSULTAR PASTA -----------------------------------"
#  hash = response.parsed_response
#  ap hash, options = {} 
#end
#
#data[:operation] = 'LerCorreio'
#response = c.post(data)
#
#unless response.body.include? 'DOCTYPE'
#  puts "\n\n------------------------------------- LER CORREIO -------------------------------------"
#  hash = response.parsed_response
#  ap hash, options = {} 
#end
#
#
#data[:operation] = 'ObterAnexo'
#response = c.post(data)

#unless response.body.include? 'DOCTYPE'
#  puts "\n\n---------------------------------------- ANEXO ----------------------------------------"
#  hash = response.parsed_response
#  ap hash, options = {} 
#
#  
#  #Prawn::Document.new("decoded_pdf.pdf") do 
#  #  font_families.update("Arial" => {
#  #    :normal => "C:\\Windows\\Fonts\\arial.ttf",
#  #    :italic => "C:\\Windows\\Fonts\\ariali.ttf",
#  #    :bold => "C:\\Windows\\Fonts\\arialbd.ttf",
#  #    :bold_italic => "C:\\Windows\\Fonts\\arialbi.ttf",
#  #  })
#  #  font "Arial"
#  #  texto = Base64.decode64(hash["soap:Envelope"]["soap:Body"]["ObterAnexoResponse"]["ObterAnexoResult"]["Anexo"]["Conteudo"])   
#  #  text texto
#  #  #text texto.force_encoding('iso-639-3').encode('utf-8')
#  #  #pdf.render_file(Dir.pwd + '/decoded.pdf')
#  #end
#
#  File.open("text.txt", "w") { |file| file.write hash["soap:Envelope"]["soap:Body"]["ObterAnexoResponse"]["ObterAnexoResult"]["Anexo"]["Conteudo"] }
#  File.open("decoded.pdf", "w") { |file| file.write(Base64.decode64(hash["soap:Envelope"]["soap:Body"]["ObterAnexoResponse"]["ObterAnexoResult"]["Anexo"]["Conteudo"])) }
#end

