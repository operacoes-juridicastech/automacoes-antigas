TAB = '	'

def check_cpf(cpf_number)
    return false if cpf_number.length != 11 
    sum = 0
    9.times do |i|
        sum += (10 - i) * cpf_number[i].to_i
    end
    
    quo = sum % 11
    if quo < 2
        return false if cpf_number[-2] != '0'
    else
        return false if 11 - quo != cpf_number[-2].to_i
    end
    
    sum = 0
    10.times do |i|
        sum += (11 - i) * cpf_number[i].to_i
    end
    
    quo = sum % 11
    if quo < 2
        return false if cpf_number[-1] != '0'
    else
        return false if 11 - quo != cpf_number[-1].to_i
    end

    true
end

def fix_agencias(array)
    if array[8] == ""
        File.readlines("CONTATOS.TXT").each do |line|
            data = line.scrub('').split(TAB).each { |i| i.gsub("\n", "") }
            array = data if array[1] == data[1]  
        end
    end
    array
rescue => exc
    puts exc.backtrace
    puts exc.message
    sleep(30)
    return nil
end

def fix_contas(array)
    array[1] = "0100" if array[3] == "2"
    array
rescue => exc
    puts exc.backtrace
    puts exc.message
    sleep(30)
    return nil
end

def fix_titulares(array)
    array[1] = "0100" if array[3] == "2"
    array
end

def fix_extrato(array)
    array[2] = "0100" if array[4] == "2"
    if array[8] == ''
        File.readlines("CNAB.txt").each do |line|
            data = line.scrub('').split(TAB)
            array[8] = data[1] if array[7] == data[0]
        end
    end
    array
end

def fix_origem_destino(array)
    array[8] = "1" if array[9][0..2] == "000" and check_cpf(array[9][3..-1])
    array
end

def fix_files(file)
    file_type = ["AGENCIAS.TXT", "CONTAS.TXT", "TITULARES.TXT", "EXTRATO.TXT", "ORIGEM_DESTINO.TXT"]
    return nil if not file_type.any? { |word| file.include? word }
    puts "Atualizando " + file 
    final_string = ""

    File.readlines(file).each do |line|
        data = line.scrub('').split(TAB).each { |i| i.gsub("\n", "") }
        #data = line.scrub('').split(TAB)

        fixed_data = fix_agencias(data) if file.include? "AGENCIAS.TXT"
        fixed_data = fix_contas(data) if file.include? "CONTAS.TXT"
        fixed_data = fix_titulares(data) if file.include? "TITULARES.TXT"
        fixed_data = fix_extrato(data) if file.include? "EXTRATO.TXT"
        fixed_data = fix_origem_destino(data) if file.include? "ORIGEM_DESTINO.TXT"
            
        final_string += fixed_data.join(TAB) + "\n"
    end
    
    return nil if final_string == ""
    final_string
rescue ArgumentError => e
    return nil
end

def start
    files = Dir.entries(Dir.pwd)
    
    if files.include? 'Resultados'
        system "del /q Resultados\\*"
    else
        system "mkdir Resultados"
    end
    
    files.each do |file| 
        fixed_file = fix_files(file)
        next if fixed_file == nil
        
        File.open("Resultados\\" + file, "w") do |file|
            file.puts fixed_file
        end
    end
end 

start