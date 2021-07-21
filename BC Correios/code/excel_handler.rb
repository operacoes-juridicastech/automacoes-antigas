require 'rubyXL'

module ExcelHandler
  def self.ask
    "I'm here!"
  end

  def self.get_workbook(path)
    workbook = RubyXL::Parser.parse path
  rescue => exception
    puts "ERROR: " + exception.message
    return nil
  end

  def self.get_postal_units(workbook)

    sheet = workbook[0]
    units = []
    
    (1..1000).each do |i|
      break if sheet[i][0].value == 'FIM'
      unit = { name: sheet[i][0].value, 
               dependency: sheet[i][2].value,
               operator: sheet[i][3].value,
               password: sheet[i][4].value }
      units.append unit
    end
    
    units
  end
end
