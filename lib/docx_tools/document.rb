module DocxTools
  class Document

    # the stored zip file (docx)
    attr_accessor :zip

    # read the given object and parse it as a zip file
    def initialize(file_object)
      if file_object.respond_to?(:read)
        self.zip = Zip::File.new(file_object, true)
      elsif file_object.is_a?(String)
        self.zip = Zip::File.open(file_object)
      else
        fail ArgumentError, 'File must be either a filename or an IO-like object.'
      end
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
