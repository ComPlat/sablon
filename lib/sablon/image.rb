require 'pathname'

module Sablon
  # Public value object describing an image (raw bytes + optional pixel
  # dimensions). Used both as the preview picture of a {Sablon::Chem}
  # embedding and as the ComPlat/chemotion_ELN compatible image API.
  #
  # NOTE: native standalone image insertion (`@field:start`) is handled by
  # upstream's {Sablon::Content::Image}. This value object is retained as the
  # ELN-facing factory and additionally duck-types the contract that
  # {Sablon::Statement::Image} / {Sablon::Processor::Document::ImageBlock}
  # require of a context value (+rid_by_file+, +local_rid+, EMU +width+/
  # +height+), so the same object works for both the chem preview image
  # (which reads +rid+) and standalone +@field+ insertion.
  class Image
    # Pixels-to-EMU scale. The ELN supplies image dimensions in pixels
    # (e.g. ImageMagick +columns+/+rows+), but native ImageBlock writes
    # width/height straight into the DrawingML EMU attributes. This factor
    # preserves the legacy ComPlat fork's sizing so existing reports do not
    # reflow; it is an intentional dpi-compat constant, NOT the OOXML
    # canonical value (which is 9525 = 1px @ 96dpi).
    PX_TO_EMU = 3000

    # Pixel dimensions are stored under +px_width+/+px_height+; +width+/
    # +height+ expose them in EMU for the native image path.
    Definition = Struct.new(:name, :data, :px_width, :px_height, :rid) do
      # Per-file relationship ids, populated by Statement::Image#set_local_rid
      # when the image is inserted via a native +@field+ block.
      def rid_by_file
        @rid_by_file ||= {}
      end

      # Native ImageBlock reads +local_rid+; the chem path reads +rid+. Alias
      # them onto the single +rid+ slot so both stay consistent.
      def local_rid
        rid
      end

      def local_rid=(value)
        self.rid = value
      end

      # EMU width, or nil when unset so ImageBlock keeps the placeholder's
      # geometry (used by the size-less status icon).
      def width
        px_width && px_width * PX_TO_EMU
      end

      # EMU height, or nil when unset (see {#width}).
      def height
        px_height && px_height * PX_TO_EMU
      end

      def inspect
        "#<Image #{name}:#{px_width}x#{px_height}px:#{rid}>"
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
