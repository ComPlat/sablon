module Sablon
  class Ole
    include Singleton
    attr_reader :definitions

    Definition = Struct.new(:name, :data, :rid) do
      def inspect
        "#<Ole #{name}:#{data}:#{rid}"
      end
    end

    def self.create_by_path(path)
      ole_name = "#{Random.new_seed}-#{Pathname.new(path).basename.to_s}"
      Sablon::Ole::Definition.new(ole_name, IO.binread(path))
    end
  end
end
