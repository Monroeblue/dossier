module Dossier
  module Adapter
    class ActiveRecord
      class Result

        attr_accessor :result

        def initialize(activerecord_result)
          self.result = activerecord_result
        end
  
        def headers
          self.result.column_names
        end

        def rows
          result
        end

      end
    end

  end
end
