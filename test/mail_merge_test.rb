require 'test_helper'

class MailMergeTest < Minitest::Test

  MERGE_FIELDS = {
    'apts_agreement_principal_investigator' => 'test1',
    'apts_agreement_principal_investigator_credentials' => 'test2',
    'apts_agreement_institution_address' => 'test3',
    'apts_agreement_institution_city' => 'test4',
    'apts_agreement_institution_state_province' => 'test5',
    'apts_agreement_institution_zip_postal_code' => 'test6',
    'apts_agreement_account' => 'test7'
  }

  def test_regexp
    refute_nil DocxTools::MailMerge::REGEXP
  end

  def test_initialize
    mail_merge = build
    assert_kind_of DocxTools::Document, mail_merge.document
    assert_kind_of DocxTools::PartList, mail_merge.part_list
  end

  def test_fields
    assert_equal MERGE_FIELDS.keys, build.fields
  end

  def test_merge
    mail_merge = build
    assert_equal MERGE_FIELDS.size, mail_merge.fields.size

    mail_merge.merge(MERGE_FIELDS)
    assert mail_merge.fields.size.zero?
  end

  def test_write
    mail_merge = build
    filepath = File.expand_path('../fixtures/written.docx', __FILE__)

    begin
      refute File.file?(filepath)
      mail_merge.write(filepath)
      assert File.file?(filepath)
    ensure
      File.delete(filepath)
    end
  end

  private

    def build
      filepath = File.expand_path('../fixtures/test-doc.docx', __FILE__)
      DocxTools::MailMerge.new(filepath)
    end
end
