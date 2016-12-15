# The code of that class was inspired in "kubido Fork - https://github.com/kubido/"

module Sablon
  module Processor
    class Chem
      # PICTURE_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
      # MAIN_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/main'
      RELATIONSHIPS_NS_URI = 'http://schemas.openxmlformats.org/package/2006/relationships'
      IMG_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
      OLE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'

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
        @@oles_rids = {}
        # @@imgs_rids = {}
        relationships = @doc.at_xpath('r:Relationships', r: RELATIONSHIPS_NS_URI)

        @@chems.to_a.each do |chems|
          # add_img_to_relationships is done in Sablon::Processor::Image
          add_ole_to_relationships(relationships, next_id, chems.ole)
          next_id += 1
        end

        @doc
      end

      def add_ole_to_relationships(relationships, next_id, ole)
        relationships.add_child("<Relationship Id='rId#{next_id}' Type='#{OLE_TYPE}' Target='embeddings/#{ole.name}'/>")
        ole.rid = next_id
        @@oles_rids[ole.name.match(/(.*)\.[^.]+$/)[1]] = next_id
      end

      # def add_img_to_relationships(relationships, next_id, img)
      #   relationships.add_child("<Relationship Id='rId#{next_id}' Type='#{IMG_TYPE}' Target='media/#{img.name}'/>")
      #   img.rid = next_id
      #   @@imgs_rids[img.name.match(/(.*)\.[^.]+$/)[1]] = next_id
      # end

      def self.add_chems_to_zip!(content, zip_out)
        (@@chems = get_all_chems(content)).each do |chem|
          add_imgs_to_zip!(chem.img, zip_out)
          add_oles_to_zip!(chem.ole, zip_out)
        end
      end

      def self.add_imgs_to_zip!(img, zip_out)
        zip_out.put_next_entry(File.join('word', 'media', img.name))
        zip_out.write(img.data)
      end

      def self.add_oles_to_zip!(ole, zip_out)
        zip_out.put_next_entry(File.join('word', 'embeddings', ole.name))
        zip_out.write(ole.data)
      end

      def self.list_ole_ids
        @@oles_rids
      end

      # def self.list_img_ids
      #   @@imgs_rids
      # end

      def self.get_all_chems(content)
        result = []

        if content.is_a?(Sablon::Chem::Definition)
          result << content
        elsif content && (content.is_a?(Enumerable) || content.is_a?(OpenStruct))
          content = content.to_h if content.is_a?(OpenStruct)
          result += content.collect do |key, value|
            if value
              get_all_chems(value)
            else
              get_all_chems(key)
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
