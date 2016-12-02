module Sablon
  class OleImg
    include Singleton
    attr_reader :definitions

    Definition = Struct.new(:ole, :img) do
      def inspect
        "#<OleImg #{ole}:#{img}"
      end
    end

    def self.create(ole, img)
      Sablon::OleImg::Definition.new(ole, img)
    end
  end
end
