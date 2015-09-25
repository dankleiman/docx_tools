module DocxTools
  class Document

    # the stored zip file (docx)
    attr_accessor :zip

    # read the filepath and parse it as a zip file
    def initialize(filepath)
      self.zip = Zip::File.open(filepath)
    end

    # the entries contained within the zip
    def entries
      zip.entries
    end

    # read the file within the zip at the given filename
    def get(filename)
      Nokogiri::XML(zip.read(filename))
    end
  end
end
