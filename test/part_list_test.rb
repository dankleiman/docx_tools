require 'test_helper'

module TestDoubles
  class Document
    def get(filename)
      Part.new
    end
  end

  class Part
    def xpath(path)
      []
    end
  end
end

class PartListTest < Minitest::Test

  def test_each_part
    assert_equal 3, build(a: 1, b: 2, c: 3).each_part.size
  end

  def test_get
    value = Object.new
    assert_equal value, build(key: value).get(:key)
  end

  def test_has?
    value = Object.new
    assert build(value => nil).has?(value)
  end

  private

    def build(parts)
      part_list = DocxTools::PartList.new(TestDoubles::Document.new, [])
      part_list.parts = parts
      part_list
    end
end
