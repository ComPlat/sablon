require 'pathname'

module Sablon
  # Public value object describing an image (raw bytes + optional pixel
  # dimensions). Used both as the preview picture of a {Sablon::Chem}
  # embedding and as the ComPlat/chemotion_ELN compatible image API.
  #
  # NOTE: native standalone image insertion (`@field:start`) is handled by
  # upstream's {Sablon::Content::Image}. This class is retained as the
  # ELN-facing factory/value object.
  class Image
    Definition = Struct.new(:name, :data, :width, :height, :rid) do
      def inspect
        "#<Image #{name}:#{width}x#{height}:#{rid}>"
      end
    end

    def self.create(name, data, width = nil, height = nil)
      Definition.new(name, data, width, height)
    end

    def self.create_by_path(path)
      name = "#{Random.new_seed}-#{Pathname.new(path).basename}"
      Definition.new(name, IO.binread(path))
    end
  end
end
