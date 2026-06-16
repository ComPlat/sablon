require 'pathname'

module Sablon
  # Public value object describing an OLE object (e.g. an embedded ChemDraw
  # document) to be inserted into a template via a `$$field:start`/`:end`
  # chem merge field. The `rid` is assigned during processing once the
  # embedding has been added to the document's relationships.
  class Ole
    Definition = Struct.new(:name, :data, :rid) do
      def inspect
        "#<Ole #{name}:#{rid}>"
      end
    end

    def self.create(name, data)
      Definition.new(name, data)
    end

    def self.create_by_path(path)
      name = "#{Random.new_seed}-#{Pathname.new(path).basename}"
      Definition.new(name, IO.binread(path))
    end
  end
end
