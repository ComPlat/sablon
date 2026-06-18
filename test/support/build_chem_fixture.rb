# Builds minimal DOCX templates exercising the Chem/OLE feature:
#   chem_template.docx       - a single $$structure:start ... :end field.
#   chem_loop_template.docx  - the same VML object inside an items:each loop,
#                              i.e. several chem objects inserted in one render
#                              (mirrors the ELN's objs:each(obj) report path).
# Each wraps a placeholder VML OLE object (v:shape > v:imagedata plus
# o:OLEObject). Run with: ruby test/support/build_chem_fixture.rb
require 'zip'

OUT = File.expand_path('../fixtures/chem_template.docx', __dir__)
LOOP_OUT = File.expand_path('../fixtures/chem_loop_template.docx', __dir__)

def complex_field(expr)
  <<~XML
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> MERGEFIELD  #{expr}  \\* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:rPr><w:noProof/></w:rPr><w:t>«#{expr}»</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  XML
end

vml_object = <<~XML
  <w:r><w:object w:dxaOrig="18787" w:dyaOrig="5510">
    <v:shape id="_x0000_i1027" type="#_x0000_t75" style="width:453pt;height:129pt" o:ole="">
      <v:imagedata r:id="rId9" o:title=""/>
    </v:shape>
    <o:OLEObject Type="Embed" ProgID="ChemDraw.Document.6.0" ShapeID="_x0000_i1027" DrawAspect="Content" ObjectID="_1832216511" r:id="rId12"/>
  </w:object></w:r>
XML

def document(body)
  <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:v="urn:schemas-microsoft-com:vml"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body>
        #{body}
        <w:sectPr/>
      </w:body>
    </w:document>
  XML
end

document_xml = document(<<~BODY)
  <w:p>
    #{complex_field('$$structure:start')}
    #{vml_object}
    #{complex_field('$$structure:end')}
  </w:p>
BODY

# items:each(item) loop wrapping a per-iteration chem field $$item:start/end.
# The loop start/end live in their own paragraphs (ParagraphBlock), the chem
# field is inline in the body paragraph that gets duplicated per item.
loop_document_xml = document(<<~BODY)
  <w:p>#{complex_field('items:each(item)')}</w:p>
  <w:p>
    #{complex_field('$$item:start')}
    #{vml_object}
    #{complex_field('$$item:end')}
  </w:p>
  <w:p>#{complex_field('items:endEach')}</w:p>
BODY

document_rels = <<~XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  </Relationships>
XML

root_rels = <<~XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  </Relationships>
XML

content_types = <<~XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  </Types>
XML

def write_docx(out, doc_xml, content_types, root_rels, document_rels)
  File.delete(out) if File.exist?(out)
  Zip::File.open(out, create: true) do |zip|
    zip.get_output_stream('[Content_Types].xml') { |f| f.write(content_types) }
    zip.get_output_stream('_rels/.rels') { |f| f.write(root_rels) }
    zip.get_output_stream('word/document.xml') { |f| f.write(doc_xml) }
    zip.get_output_stream('word/_rels/document.xml.rels') { |f| f.write(document_rels) }
  end
  puts "wrote #{out}"
end

write_docx(OUT, document_xml, content_types, root_rels, document_rels)
write_docx(LOOP_OUT, loop_document_xml, content_types, root_rels, document_rels)
