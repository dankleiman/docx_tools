require 'test_helper'

class DocumentTest < Minitest::Test

  def test_initialize
    assert_kind_of Zip::File, build.zip
  end

  def test_entries
    assert_equal 22, build.entries.size
  end

  def test_get
    document = build
    ['[Content_Types].xml', 'word/document.xml', 'word/footer1.xml', 'word/header1.xml'].each do |filename|
      assert_kind_of Nokogiri::XML::Document, build.get(filename)
    end
  end

  private

    def build
      filepath = File.expand_path('../fixtures/test-doc.docx', __FILE__)
      DocxTools::Document.new(filepath)
    end
end
