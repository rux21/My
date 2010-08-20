require 'rubygems'
require 'roo'

ts = Excelx.new('pics.trial.xlsx')

ts.default_sheet = ts.sheets.first
2.upto(3) do |line|

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

    f.puts "<H1> #{exercise} </H1>"
    f.puts ""
    f.puts "<table border=1 width=\"90%\" align=\"center\">"
    f.puts "<TR>"
    #if more than one picture then "Exercise position" changes to "Exercise Start Position"
    #then another TH is added for "Exercise End Position"... and there's a mid position
    f.puts "\t<TH align=\"center\">Exercise Position</TH>"
    f.puts "</TR>"
    f.puts "<TR>"
    f.puts "\t<TD width=\"50%\"><IMG src=\"/home_ed/images/#{start_image}\" style=\"width:100%;height:auto\"></TD>"
    # I need some kind of conditional statement that checks for a second and thrid picture and places them here
    # will do that later. I'm sure its easy
    f.puts "</TR>"
    f.puts "</table>"
    f.puts "<BR/><BR/>"
    f.puts "<h2>Performance Cues</h2>"
    f.puts "<ul style=\"margin-left:5%\">"
    f.puts ""
    f.puts "\t<li>#{cue1}</li>"
    f.puts "\t<li>#{cue2}</li>"
    f.puts "\t<li>#{cue3}</li>"
    f.puts ""
    f.puts "</ul>"
    f.puts ""
    f.puts "<P></P>"
    f.puts "<div style=\"width:100%;text-align:center\"><%= @case_home_education.provider_instructions %></div>"
    
    end
  end

end

