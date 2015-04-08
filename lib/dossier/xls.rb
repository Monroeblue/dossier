module Dossier
  class Xls
    attr_reader :report, :headers, :collection

    HEADER = %Q{<?xml version="1.0" encoding="UTF-8"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">\n<Worksheet ss:Name="Sheet1">\n<Table>\n}
    FOOTER = %Q{</Table>\n</Worksheet>\n</Workbook>\n}

    def initialize(report)
      @report = report
      @headers    = report.columns.select{|col| col unless report.format_header(col).blank? }
      @collection = report.raw_results
    end

    def each
      yield HEADER
      yield headers_as_row
      @collection.each { |record| yield as_ar_row(record) }
      yield FOOTER
    end

    private

    def as_cell(el)
      if el.kind_of?(Array)
        %{<Cell><Data ss:Type="String">#{el.map(&:to_s).join(', ')}</Data></Cell>}
      elsif el.is_a?(DateTime)
        %{<Cell><Data ss:Type="DateTime">#{el}</Data></Cell>}
      elsif el.is_a?(Float) || el.is_a?(Integer) || el.is_a?(BigDecimal)
        %{<Cell><Data ss:Type="Number">#{el}</Data></Cell>}
      else
        %{<Cell><Data ss:Type="String">#{CGI::escapeHTML(el.to_s)}</Data></Cell>}
      end
    end

    def as_ar_row(row)
      "<Row>\n" +  headers.collect{|column|

        args = [column]        
        if (row.method(column).arity == -1  rescue false)
           args << report.options
        end
        value = row.public_send(*args)

        #use export formatter for this cell, if we have one
        if report.respond_to?("export_format_#{column}") 
          args = ["export_format_#{column}", value, row]
          value = report.public_send(*args) 
        end
          
        as_cell(value)

      }.join("\n") + "\n</Row>\n"
    end

    def as_row(array)
      my_array = array.map{|a| as_cell(a)}.join("\n")
      "<Row>\n" + my_array + "\n</Row>\n"
    end

    def headers_as_row
      as_row(@headers.map { |header| report.format_header(header) })
    end
  end
end
