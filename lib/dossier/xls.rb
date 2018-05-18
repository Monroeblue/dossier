require 'write_xlsx'

module Dossier
  class Xls
    attr_reader :report, :headers, :collection

    def initialize(report)
      @report = report
      @headers = report.columns.select { |col| col unless report.format_header(col).blank? }
      @collection = report.raw_results
      @string_buffer = StringIO.new
      @workbook = WriteXLSX.new(@string_buffer)
      @worksheet = @workbook.add_worksheet
    end

    def each
      add_header_row
      @collection.each_with_index { |record, row_index| as_ar_row(record, row_index + 1) }
      @workbook.close
      yield @string_buffer.string
    end

    private

    def as_ar_row(row, row_index)
      headers.each_with_index do |column, column_index|
        args = [column]
        if begin
              row.method(column).arity == -1
            rescue
              false
            end
           args << report.options
        end
        value = row.public_send(*args)

        # use export formatter for this cell, if we have one
        if report.respond_to?("export_format_#{column}")
          args = ["export_format_#{column}", value, row]
          value = report.public_send(*args)
        end

        value = value.map(&:to_s).join(', ') if value.is_a?(Array)
        @worksheet.write(row_index, column_index, value)
      end
    end

    def add_header_row
      header_format = @workbook.add_format
      header_format.set_bold

      @headers.each_with_index do |header, column_index|
        @worksheet.write(0, column_index, report.format_header(header), header_format)
      end
    end
  end
end
