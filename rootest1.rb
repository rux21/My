require 'rubygems'
require 'roo'

ts = Excelx.new('pics.trial.xlsx')

ts.default_sheet = ts.sheets.first
2.upto(4) do |line|

  exercise = ts.cell(line,'A')

  start_image = ts.cell(line,'C')

  end_image = ts.cell(line,'D')

  mid_image = ts.cell(line,'E')

  cue1 = ts.cell(line,'F')

  cue2 = ts.cell(line,'G')

  cue3 = ts.cell(line,'H')

  cue4 = ts.cell(line,'I')
  
  if exercise
    htmlfile = exercise.gsub(' ','_')
    filename = "#{htmlfile}.html"
    
    #puts filename
    File.open(filename,'w') do |f|

    f.puts "#{exercise}\t#{start_image}\t#{end_image}"
    f.puts "first que is: #{cue1}"
    f.puts "2nd que is: #{cue2}"
    f.puts "3rd que is: #{cue3}"
    
    end
  end

end

