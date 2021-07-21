require 'creek'
require 'date'

module Auditor 

    def self.get_files
        files = Dir.entries(Dir.pwd)
        files.each { |file| file = Dir.pwd.gsub('/', '\\') + '\\' + file }
        files.filter { |file| file.include? '.xls' }
    end

    def self.filter_date(date)
        if date == nil or date == '-' or date == ""
            return nil
        end

        separator = date.index('-')

        if separator == 4
            return Date.parse "#{date[8..9]}/#{date[5..6]}/#{date[0..3]}"
        else
            if separator == nil
                separator = -1                 
            end
            return Date.parse date[separator+1..-1]
        end
    end

    def self.run
        File.open("analise.txt", "w+") do |arquivo|
            files = get_files
            files_number = get_files.length
            files_counter = 0 
            files.each do |file|
                files_counter += 1
                puts "Lendo arquivo (#{files_counter}/#{files_number}) - #{file}"
                creek = Creek::Book.new file
                index = 0
                while true
                    sheet = creek.sheets[index]
                    if sheet == nil
                        index = 0
                        break
                    else
                        index += 1
                        if sheet.name == 'Smart Report'
                            row_number = 0
                            process_number = 0
                            violations = [0, 0, 0]
                            data_relatorio = ""

                            sheet.simple_rows.each do |row|
                                row_number += 1
                                if row_number >= 8 and row['G'] != 'Ativo'
                                    break
                                end
                                
                                if row_number == 2 
                                    data = row["A"]
                                    data_relatorio = data[data.index('/')-2..data.index('/')+8]
                                    data_relatorio = Date.parse(data_relatorio)
                                end
                                
                                if row_number < 5 
                                    next
                                end

                                process_number += 1
                                
                                data_ajuizamento = filter_date(row["I"].to_s)
                                data_movimentacao = filter_date(row["J"].to_s)
                                data_recebimento = filter_date(row["K"].to_s)
                                data_citacao = filter_date(row["L"].to_s)
                                    
                                #Ponto 1
                                if data_citacao == nil or data_recebimento == nil or (data_citacao - data_recebimento).to_i.abs >= 15
                                    violations[0] += 1
                                end
                                #Ponto 2
                                if data_citacao == nil or data_ajuizamento == nil or (data_citacao - data_ajuizamento).to_i.abs >= 90
                                    violations[1] += 1
                                end
                                #Ponto 3
                                if data_movimentacao == nil or (data_movimentacao - data_relatorio).to_i.abs >= 30
                                    violations[2] += 1
                                end                            

                            end
                            
                            arquivo.write "#{file} -> Problema 1 - [#{violations[0]}/#{process_number}] | Problema 2 - [#{violations[1]}/#{process_number}] | Problema 3 - [#{violations[2]}/#{process_number}]\n"
                            #puts "#{file} -> Problema 1 - [#{violations[0]}/#{process_number}] | Problema 2 - [#{violations[1]}/#{process_number}] | Problema 3 - [#{violations[2]}/#{process_number}]"
                        end
                    end
                end
            end
        end
    rescue => exception
        puts exception.backtrace
        puts exception.message
        sleep(100)
    end

end

Auditor.run