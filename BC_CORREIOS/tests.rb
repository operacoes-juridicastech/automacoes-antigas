require 'rspec/autorun'
require_relative 'main'

describe Crawler do
  subject { Crawler.new }
    
  context 'when building the post requests' do
    it 'should get the valid header' do
      valid_header = {
          "Host": "bccorreio.bcb.gov.br",
          "Content-Type": "text/xml",
          "Content-Length": "100",
          "SOAPAction": "http://www.bcb.gov.br/correiows/ConsultarCorreiosPorPasta"
      }
      expect(subject.get_header(100, "ConsultarCorreiosPorPasta")).to eq(valid_header)
    end

    it 'should build the correct XML when operation is ConsultarPastasAutorizadas' do
      expected_xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"\
                     "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"\
                     "<soap:Body>"\
                     "<ConsultarPastasAutorizadas xmlns=\"http://www.bcb.gov.br/correiows\"/>"\
                     "</soap:Body>"\
                     "</soap:Envelope>"
      data = { operation: 'ConsultarPastasAutorizadas' }
      expect(subject.build_xml(data)).to eq(expected_xml)
    end

    it 'should build the correct XML when operation is ConsultarCorreiosPorPasta' do
      dates = subject.get_dates

      data = { 
        operation: 'ConsultarCorreiosPorPasta',
        unidade_nome: '23181',
        unidade_tipo: 'InstituicaoFinanceira',
        dependencia: '0001',
        tipo: 'CaixaEntrada',
        pagina: '0'
      }

      expected_xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"\
      "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"\
      "<soap:Body>"\
        "<ConsultarCorreiosPorPasta xmlns=\"http://www.bcb.gov.br/correiows\">"\
          "<consulta>"\
            "<Pasta>"\
              "<Unidade>"\
                "<Nome>23181</Nome>" \
                "<Ativa>true</Ativa>"\
                "<Tipo>InstituicaoFinanceira</Tipo>"\
              "</Unidade>"\
              "<Dependencia>0001</Dependencia>"\
              "<Setor/>"\
              "<Tipo>CaixaEntrada</Tipo>"\
            "</Pasta>"\
            "<Setor>#{data[:setor_nome]}</Setor>"\
            "<DataInicial>#{dates[:yesterday]}T00:00:01</DataInicial>"\
            "<DataFinal>#{dates[:yesterday]}T23:59:59</DataFinal>"\
            "<AssuntoConteudo>Correio</AssuntoConteudo>"\
            "<Unidade>SECRE</Unidade>"\
            "<ApenasMensagens>true</ApenasMensagens>"\
            "<PesquisarEmTodasAsPastas>false</PesquisarEmTodasAsPastas>"\
            "<Pagina>#{data[:pagina]}</Pagina>"\
          "</consulta>"\
        "</ConsultarCorreiosPorPasta>"\
      "</soap:Body>"\
    "</soap:Envelope>"

      expect(subject.build_xml(data)).to eq(expected_xml)
    end
  end

  context 'when communicating with the API' do
    it 'should get a valid response' do
      data = {
          operation: "ConsultarPastasAutorizadas",
          username: "231810001.S-WJUD1501" ,
          password: "BJUD2019"
      }

      response = subject.post(data)
      puts response.parsed_response
      expect(response.code).to eq(200)
    end
  end
end