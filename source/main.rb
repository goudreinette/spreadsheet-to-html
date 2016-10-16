require 'spreadsheet'

class SpreadsheetConverter
    def initialize(path)
        @path        = path
        @spreadsheet = Spreadsheet.open(path)
        @worksheet   = @spreadsheet.worksheet 0 
        @headings    = @worksheet.row 0
        @data        = @worksheet.drop 1
    end

    def write()
        directory = File.dirname(@path)
        basename  = File.basename(@path, '.*')
        filepath  = "#{directory}/#{basename}.html"

        File.new(filepath, "w").write(to_html)
    end


    def to_html
        "<table>
            <thead>#{render_headings}</thead>
            <tbody>#{render_data}</tbody>
        </table>"
    end

    def render_headings()
        render_columns('th', @headings)
    end

    def render_data
        @data.inject("") do |html, row|
            render_columns('td', row)
        end
    end

    def render_columns(tag, rows)
        rows.inject("") do |html, row|
            html + "<#{tag}>#{row}</#{tag}>"
        end
    end
end



path = '/home/reinvdwoerd/Documents/testsheet.xls'
spreadsheet_converter = SpreadsheetConverter.new(path)
result = spreadsheet_converter.write
p result