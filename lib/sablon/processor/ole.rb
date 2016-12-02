# The code of that class was inspired in "kubido Fork - https://github.com/kubido/"

module Sablon
  module Processor
    class Ole
      # PICTURE_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
      # MAIN_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/main'
      # RELATIONSHIPS_NS_URI = 'http://schemas.openxmlformats.org/package/2006/relationships'
      # IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

      def self.process(doc, properties, out)
        processor = new(doc, properties, out)
        processor.manipulate
      end

      def initialize(doc, properties, out)
        @doc = doc
        @properties = properties
        @out = out
      end

      def manipulate
        next_id = next_rel_id
        @@images_rids = {}
        relationships = @doc.at_xpath('r:Relationships', r: RELATIONSHIPS_NS_URI)

        @@images.to_a.each do |image|
          relationships.add_child("<Relationship Id='rId#{next_id}' Type='#{IMAGE_TYPE}' Target='media/#{image.name}'/>")
          image.rid = next_id
          @@images_rids[image.name.match(/(.*)\.[^.]+$/)[1]] = next_id
          next_id += 1
        end

        @doc
      end




















      def self.add_images_to_zip!(content, zip_out)
        (@@images = get_all_images(content)).each do |image|
          zip_out.put_next_entry(File.join('word', 'media', image.name))
          zip_out.write(image.data)
        end
      end

      def self.list_ids
        @@images_rids
      end

      def self.get_all_images(content)
        result = []

        if content.is_a?(Sablon::Image::Definition)
          result << content
        elsif content && (content.is_a?(Enumerable) || content.is_a?(OpenStruct))
          content = content.to_h if content.is_a?(OpenStruct)
          result += content.collect do |key, value|
            if value
              get_all_images(value)
            else
              get_all_images(key)
            end
          end.compact
        end

        result.flatten.compact
      end

      private

      def next_rel_id
        @doc.xpath('r:Relationships/r:Relationship', 'r' => RELATIONSHIPS_NS_URI).inject(0) do |max ,n|
          id = n.attributes['Id'].to_s[3..-1].to_i
          [id, max].max
        end + 1
      end
    end
  end
end
