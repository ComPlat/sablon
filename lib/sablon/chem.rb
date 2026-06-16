module Sablon
  # Public value object pairing an embedded OLE object with its preview
  # image. Inserted into a template through a chem merge field of the form
  # `$$field:start` / `$$field:end`. Mirrors the ComPlat/chemotion_ELN API
  # (`Sablon::Chem.create(ole, img)`).
  class Chem
    Definition = Struct.new(:ole, :img) do
      def inspect
        "#<Chem #{ole.inspect}:#{img.inspect}>"
      end
    end

    def self.create(ole, img)
      Definition.new(ole, img)
    end
  end
end
