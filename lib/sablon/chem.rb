module Sablon
  class Chem
    include Singleton
    attr_reader :definitions

    Definition = Struct.new(:ole, :img) do
      def inspect
        "#<Chem #{ole}:#{img}"
      end
    end

    def self.create(ole, img)
      Sablon::Chem::Definition.new(ole, img)
    end
  end
end
