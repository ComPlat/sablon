# Builds a minimal DOCX template containing a single chem merge field
# ($$structure:start ... $$structure:end) wrapping a placeholder VML OLE
# object (v:shape > v:imagedata, plus o:OLEObject). Used by the Chem/OLE
# proof-of-concept test. Run with: ruby test/support/build_chem_fixture.rb
require 'zip'

OUT = File.expand_path('../fixtures/chem_template.docx', __dir__)

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

document_xml = <<~XML
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
              xmlns:v="urn:schemas-microsoft-com:vml"
              xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
    <w:body>
      <w:p>
        #{complex_field('$$structure:start')}
        #{vml_object}
        #{complex_field('$$structure:end')}
      </w:p>
      <w:sectPr/>
    </w:body>
  </w:document>
XML

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

File.delete(OUT) if File.exist?(OUT)
Zip::File.open(OUT, create: true) do |zip|
  zip.get_output_stream('[Content_Types].xml') { |f| f.write(content_types) }
  zip.get_output_stream('_rels/.rels') { |f| f.write(root_rels) }
  zip.get_output_stream('word/document.xml') { |f| f.write(document_xml) }
  zip.get_output_stream('word/_rels/document.xml.rels') { |f| f.write(document_rels) }
end
puts "wrote #{OUT}"
