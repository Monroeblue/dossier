require 'csv'

module Dossier
  class StreamCSV
    attr_reader :report, :headers, :collection

    def initialize(report)
      @report = report
      @headers    = report.columns.select{|col| col unless report.format_header(col).blank? }
      @collection = report.raw_results
    end

    def each
      yield headers.map { |header| report.format_header(header) }.to_csv if headers?
      collection.each do |record|
        yield headers.collect{|h| record.send(h)}.to_csv
      end
    rescue => e
      if Rails.application.config.consider_all_requests_local
        yield e.message
        e.backtrace.each do |line|
          yield "#{line}\n"
        end
      else
        yield "We're sorry, but something went wrong." 
      end
    end

    private

    def headers?
      headers.present?
    end

  end
end
