require_relative '../excel_handler'

require 'minitest/autorun'
require 'rubyXL'

class ExcelHandlerTest < Minitest::Test
  def test_ask_returns_an_answer
    assert ExcelHandler.ask != nil
  end

  def test_get_workbook_given_right_path
    workbook = RubyXL::Workbook.new
    path = Dir.pwd + 'teste.xlsx'
    workbook.write path
    assert ExcelHandler.get_workbook(path).is_a? RubyXL::Workbook
    File.delete path if File.exists? path
  end

  def test_get_workbook_given_wrong_path
    assert_nil ExcelHandler.get_workbook 'some_random_string'
  end
end