class SetupXls
  require 'win32ole'
# variable_for_sheet = SetupXls.new("one")

    def self.setsheet(sheetname) #create an instance variable for a sheet
      excel = WIN32OLE::connect('excel.Application')
      worksheet = nil
      excel.Workbooks.each{|wb| 
        wb.Worksheets.each{|ws| 
          if ws.name == sheetname
            worksheet = ws
            return ws
            break
          end
        }
        break unless worksheet.nil?
      }
    end
end