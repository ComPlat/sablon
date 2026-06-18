# -*- coding: utf-8 -*-
require "test_helper"
require "zip"

# Proof-of-concept coverage for the ported Chem/OLE feature: inserting an
# embedded OLE object together with its preview image via a
# `$$expr:start`/`$$expr:end` merge field.
class SablonChemTest < Sablon::TestCase
  def setup
    super
    @base_path = Pathname.new(File.expand_path("../", __FILE__))
    @template_path = @base_path + "fixtures/chem_template.docx"
    @loop_template_path = @base_path + "fixtures/chem_loop_template.docx"
    @ole_path = @base_path + "fixtures/chem_ole_sample.bin"
    @image_path = @base_path + "fixtures/images/c3po.jpg"
  end

  def render(context, template = @template_path)
    entries = {}
    xml = Sablon.template(template).render_to_string(context)
    Zip::File.open_buffer(StringIO.new(xml)) do |zip|
      zip.each { |e| entries[e.name] = e.get_input_stream.read if e.file? }
    end
    entries
  end

  def build_chem(ole_name = "ole_object.bin", img_name = "preview.jpg")
    ole = Sablon::Ole.create(ole_name, IO.binread(@ole_path))
    img = Sablon::Image.create(img_name, IO.binread(@image_path))
    Sablon::Chem.create(ole, img)
  end

  def test_embeds_ole_and_preview_and_rewrites_references
    chem = build_chem
    entries = render(structure: chem)

    # 1. OLE payload added under word/embeddings with the original bytes
    assert entries.key?("word/embeddings/ole_object.bin"),
           "expected embedded OLE part, got: #{entries.keys.grep(%r{embeddings})}"
    assert_equal IO.binread(@ole_path), entries["word/embeddings/ole_object.bin"]

    # 2. preview image added under word/media with the original bytes
    media = entries.keys.grep(%r{word/media/.*preview\.jpg}).first
    assert media, "expected preview image in word/media, got: #{entries.keys.grep(%r{media})}"
    assert_equal IO.binread(@image_path), entries[media]

    # 3. relationships registered for both parts
    rels = entries["word/_rels/document.xml.rels"]
    image_rid = rels[%r{Target="media/[^"]*preview\.jpg"[^>]*Id="(rId\d+)"}, 1]
    ole_rid   = rels[%r{Target="embeddings/ole_object\.bin"[^>]*Id="(rId\d+)"}, 1]
    assert image_rid, "no image relationship in rels: #{rels}"
    assert ole_rid,   "no oleObject relationship in rels: #{rels}"
    assert_includes rels, "relationships/oleObject"

    # 4. content type registered for the OLE extension
    assert_includes entries["[Content_Types].xml"],
                    "application/vnd.openxmlformats-officedocument.oleObject"

    # 5. document.xml references rewritten to the freshly created rIds
    doc = entries["word/document.xml"]
    assert_match %r{<v:imagedata r:id="#{image_rid}"}, doc
    assert_match %r{<o:OLEObject[^>]*r:id="#{ole_rid}"}, doc
    # shape id and OLEObject ShapeID stay consistent
    sid = image_rid[/\d+/]
    assert_match %r{<v:shape id="id_s#{sid}"}, doc
    assert_match %r{ShapeID="id_s#{sid}"}, doc

    # 6. merge field control text removed
    refute_includes doc, "$$structure:start"
    refute_includes doc, "$$structure:end"
    refute_includes doc, "MERGEFIELD"
  end

  def test_missing_chem_removes_fields_without_embedding
    entries = render(structure: nil)
    doc = entries["word/document.xml"]
    refute_includes doc, "MERGEFIELD"
    assert_empty entries.keys.grep(%r{word/embeddings})
  end

  # Regression guard for the "Word found unreadable content" warning: a report
  # inserts many structures via an `:each` loop. Each chem object must get its
  # own embedding/media part and a consistent, unique set of references — the
  # inconsistent rewriting in the pre-0.4.3 fork is what made Word complain.
  def test_loop_inserts_each_chem_consistently
    chems = [
      build_chem("ole_a.bin", "preview_a.jpg"),
      build_chem("ole_b.bin", "preview_b.jpg"),
      build_chem("ole_c.bin", "preview_c.jpg")
    ]
    entries = render({ items: chems }, @loop_template_path)
    doc = entries["word/document.xml"]
    rels = entries["word/_rels/document.xml.rels"]

    # one distinct embedding + one distinct preview image per chem object
    assert_equal 3, entries.keys.grep(%r{word/embeddings/}).size, entries.keys.inspect
    assert_equal 3, entries.keys.grep(%r{word/media/}).size, entries.keys.inspect

    # OLE content type declared (idempotent even across three inserts)
    assert_includes entries["[Content_Types].xml"],
                    "application/vnd.openxmlformats-officedocument.oleObject"

    # every reference in the document resolves to a declared relationship
    # whose target part exists in the package (no dangling/undeclared refs)
    ref_rids = (doc.scan(/<v:imagedata r:id="(rId\d+)"/) +
                doc.scan(/<o:OLEObject[^>]*r:id="(rId\d+)"/)).flatten
    assert_equal 6, ref_rids.size, "expected 3 image + 3 ole refs, got #{ref_rids.inspect}"
    ref_rids.each do |rid|
      target = rels[/Id="#{rid}"[^>]*Target="([^"]+)"/, 1] ||
               rels[/Target="([^"]+)"[^>]*Id="#{rid}"/, 1]
      assert target, "rId #{rid} not declared in rels: #{rels}"
      assert entries.key?("word/#{target}"), "missing part word/#{target}"
    end

    # references are all distinct and the VML shape ids are unique — duplicate
    # shape ids would themselves trip Word's recovery dialog
    assert_equal ref_rids.uniq, ref_rids, "duplicate relationship references"
    shape_ids = doc.scan(/<v:shape id="([^"]+)"/).flatten
    assert_equal 3, shape_ids.size
    assert_equal shape_ids.uniq, shape_ids, "duplicate v:shape ids: #{shape_ids.inspect}"

    # no merge-field control text left behind
    refute_includes doc, "MERGEFIELD"
    refute_includes doc, ":start"
    refute_includes doc, ":endEach"
  end
end
