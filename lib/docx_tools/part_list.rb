module DocxTools
  class PartList

    # the stored list of parts
    attr_accessor :parts

    # parse the content type entries out of the given document
    def initialize(document, content_types)
      self.parts = {}

      content_types.map!(&method(:expand_type))
      document.get('[Content_Types].xml').xpath(content_types.join(' | ')).each do |tag|
        filename = tag['PartName'].split('/', 2)[1]
        parts[filename] = document.get(filename)
      end
    end

    # yield each part to the block
    def each_part(&block)
      parts.values.each(&block)
    end

    # get the requested part
    def get(filename)
      parts[filename]
    end

    # true if this part list has extracted this part from the document
    def has?(filename)
      parts.key?(filename)
    end

    private

      def expand_type(content_type)
        "xmlns:Types/xmlns:Override[@ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.#{content_type}+xml']"
      end
  end
end
