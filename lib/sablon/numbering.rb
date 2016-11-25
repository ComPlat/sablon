module Sablon
  class Numbering
    include Singleton
    attr_reader :definitions

    Definition = Struct.new(:numid, :style) do
      def inspect
        "#<Numbering #{numid}:#{style}"
      end
    end

    def initialize
      reset!
    end

    def reset!
      @numid = 0
      @definitions = []
    end

    def register(style)
      @numid += 1
      definition = Definition.new(@numid, style)
      @definitions << definition
      definition
    end
  end
end
