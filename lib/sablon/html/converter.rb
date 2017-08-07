require "sablon/html/ast"
require "sablon/html/visitor"

module Sablon
  class HTMLConverter
    class ASTBuilder
      Layer = Struct.new(:items, :ilvl)

      def initialize(nodes)
        @layers = [Layer.new(nodes, false)]
        @root = Root.new([])
      end

      def to_ast
        @root
      end

      def new_layer(ilvl: false)
        @layers.push Layer.new([], ilvl)
      end

      def next
        current_layer.items.shift
      end

      def push(node)
        @layers.last.items.push node
      end

      def push_all(nodes)
        nodes.each(&method(:push))
      end

      def done?
        !current_layer.items.any?
      end

      def nested?
        ilvl > 0
      end

      def ilvl
        @layers.select { |layer| layer.ilvl }.size - 1
      end

      def emit(node)
        @root.nodes << node
      end

      private
      def current_layer
        if @layers.any?
          last_layer = @layers.last
          if last_layer.items.any?
            last_layer
          else
            @layers.pop
            current_layer
          end
        else
          Layer.new([], false)
        end
      end
    end

    def process(input)
      processed_ast(input).to_docx
    end

    def processed_ast(input)
      ast = build_ast(input)
      ast.accept LastNewlineRemoverVisitor.new
      ast
    end

    def build_ast(input)
      doc = Nokogiri::HTML.fragment(input)
      @builder = ASTBuilder.new(doc.children)

      while !@builder.done?
        ast_next_paragraph
      end
      @builder.to_ast
    end

    private
    def ast_next_paragraph
      node = @builder.next
      if node.name == 'div'
        @builder.new_layer
        @builder.emit Paragraph.new('Normal', ast_text(node.children))
      elsif node.name == 'p'
        @builder.new_layer
        @builder.emit Paragraph.new('Paragraph', ast_text(node.children))
      elsif node.name =~ /h(\d+)/
        @builder.new_layer
        @builder.emit Paragraph.new("Heading#{$1}", ast_text(node.children))
      elsif node.name == 'ul'
        @builder.new_layer ilvl: true
        unless @builder.nested?
          @definition = Sablon::Numbering.instance.register('ListParagraph')
        end
        @builder.push_all(node.children)
      elsif node.name == 'ol'
        @builder.new_layer ilvl: true
        unless @builder.nested?
          @definition = Sablon::Numbering.instance.register('ListParagraph')
        end
        @builder.push_all(node.children)
      elsif node.name == 'li'
        @builder.new_layer
        @builder.emit ListParagraph.new(@definition.style, ast_text(node.children), @definition.numid, @builder.ilvl)
      elsif node.text?
        # SKIP?
      else
        raise ArgumentError, "Don't know how to handle node: #{node.inspect}"
      end
    end

    def get_highlight_from_hex hex
      hex = hex.to_i(16)

      return "black" if hex.between?(0x000000, 0x000080)
      return "darkblue" if hex.between?(0x000080, 0x0000FF)
      return "blue" if hex.between?(0x0000FF, 0x008000)
      return "darkGreen" if hex.between?(0x008000, 0x008080)
      return "darkCyan" if hex.between?(0x008080, 0x00FF00)
      return "green" if hex.between?(0x00FF00, 0x00FFFF)
      return "cyan" if hex.between?(0x00FFFF, 0x800000)
      return "darkRed" if hex.between?(0x800000, 0x800080)
      return "darkMagenta" if hex.between?(0x800080, 0x808000)
      return "darkYellow" if hex.between?(0x808000, 0x808080)
      return "darkGray" if hex.between?(0x808080, 0xC0C0C0)
      return "lightGray" if hex.between?(0xC0C0C0, 0xFF0000)
      return "red" if hex.between?(0xFF0000, 0xFF00FF)
      return "magenta" if hex.between?(0xFF00FF, 0xFFFF00)
      return "yello" if hex.between?(0xFFFF00, 0xFFFFFF)
      return "white"
    end

    def ast_text(nodes, format: TextFormat.default)
      runs = nodes.flat_map do |node|
        node_format = format.clone
        if node.attributes.count > 0
          styles = node.attr("style").split(";").compact
          styles.each do |style|
            style_attr = style.split(":").compact.collect(&:strip)
            hex_color = style_attr[1].delete "#; "
            font_family = style_attr[1]
            case style_attr[0]
            when 'color'
              node_format.set_color hex_color
            when 'background-color'
              node_format.set_highlight(get_highlight_from_hex(hex_color))
            when 'font-family'
              node_format.set_font_family(font_family)
            end
          end
        end

        if node.text?
          Text.new(node.text, node_format)
        elsif node.name == 'br'
          Newline.new
        elsif node.name == 'strong' || node.name == 'b'
          ast_text(node.children, format: node_format.with_bold).nodes
        elsif node.name == 'em' || node.name == 'i'
          ast_text(node.children, format: node_format.with_italic).nodes
        elsif node.name == 'u'
          ast_text(node.children, format: node_format.with_underline).nodes
        elsif node.name == 'sub'
          ast_text(node.children, format: node_format.with_subscript).nodes
        elsif node.name == 'sup' || node.name == 'super'
          ast_text(node.children, format: node_format.with_superscript).nodes
        elsif node.name == 'span'
          ast_text(node.children, format: node_format).nodes
        elsif ['ul', 'ol', 'p', 'div'].include?(node.name) ||
              (node.name =~ /h(\d+)/) != nil
          @builder.push(node)
          nil
        else
          raise ArgumentError, "Don't know how to handle node: #{node.inspect}"
        end
      end

      return Collection.new(runs.compact)
    end
  end
end
