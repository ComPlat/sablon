# -*- coding: utf-8 -*-
require "test_helper"
require "zip"

# Verifies that the ELN-facing {Sablon::Image::Definition} value object can be
# inserted through upstream's native `@field` image path. Native
# Statement::Image / ImageBlock expect a Content::Image-like object
# (+rid_by_file+, +local_rid+, EMU +width+/+height+); Image::Definition
# duck-types that contract so the ComPlat/chemotion_ELN API keeps working
# without emitting `Sablon.content(:image, ...)` at every call site.
class SablonImageBridgeTest < Sablon::TestCase
  def setup
    super
    @base_path = Pathname.new(File.expand_path("../", __FILE__))
    @template_path = @base_path + "fixtures/images_template.docx"
    @image_fixtures = @base_path + "fixtures/images"
  end

  def render(context)
    entries = {}
    xml = Sablon.template(@template_path).render_to_string(context)
    Zip::File.open_buffer(StringIO.new(xml)) do |zip|
      zip.each { |e| entries[e.name] = e.get_input_stream.read if e.file? }
    end
    entries
  end

  def image_def(file, width = nil, height = nil)
    Sablon::Image.create(file, IO.binread(@image_fixtures.join(file)), width, height)
  end

  # The crux: feeding raw Image::Definition objects (not Sablon.content) must
  # render without NoMethodError and produce a valid package — this is the
  # path that would otherwise break the ELN status/attachment images.
  def test_image_definition_renders_through_native_field
    sized = image_def('r2d2.jpg', 100, 50)
    context = {
      items: [
        { title: 'C-3PO', image: image_def('c3po.jpg') },
        { title: 'R2-D2', image: sized }
      ],
      trooper: image_def('clone.jpg'),
      r2d2: sized
    }
    entries = render(context)
    doc = entries["word/document.xml"]
    rels = entries["word/_rels/document.xml.rels"]

    # every embedded blip resolves to a declared image relationship whose
    # target part is present in the package
    rids = doc.scan(/r:embed="(rId\d+)"/).flatten.uniq
    refute_empty rids, "no images were embedded"
    rids.each do |rid|
      target = rels[/Id="#{rid}"[^>]*Target="([^"]+)"/, 1] ||
               rels[/Target="([^"]+)"[^>]*Id="#{rid}"/, 1]
      assert target, "rId #{rid} has no relationship in rels: #{rels}"
      assert entries.key?("word/#{target}"), "missing media part word/#{target}"
    end

    # pixel dimensions are scaled to EMU by the bridge (100px -> 100*3000)
    assert_includes doc, 'cx="300000"'
    assert_includes doc, 'cy="150000"'

    # no merge-field control text left behind
    refute_includes doc, "MERGEFIELD"
    refute_includes doc, ":start"
  end

  def test_px_to_emu_constant_matches_legacy_fork
    assert_equal 3000, Sablon::Image::PX_TO_EMU
    assert_equal 300_000, image_def('r2d2.jpg', 100, 50).width
    assert_nil image_def('c3po.jpg').width
  end
end
