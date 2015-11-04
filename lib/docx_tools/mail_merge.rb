module DocxTools
  class MailMerge

    REGEXP = / MERGEFIELD "?([^ ]+?)"? (| \\\* MERGEFORMAT )/i.freeze
    attr_accessor :document, :part_list

    def initialize(file_object)
      self.document  = Document.new(file_object)
      self.part_list = PartList.new(document, %w[document.main header footer])
      process_merge_fields
    end

    def fields
      fields = Set.new
      part_list.each_part do |part|
        part.xpath('.//w:MergeField').each do |mf|
          fields.add(mf.content)
        end
      end
      fields.to_a
    end

    def merge(replacements = {})
      part_list.each_part do |part|
        replacements.each do |field, text|
          merge_field(part, field, text)
        end
      end
    end

    def write(filename)
      File.open(filename, 'w') do |file|
        file.write(generate.string)
      end
    end

    private

      def clean_up
        remaining = fields.map { |field| [field.to_sym, ''] }
        merge(remaining.to_h)
      end

      def generate
        clean_up
        buffer = Zip::OutputStream.write_buffer do |out|
          document.entries.each do |entry|
            unless entry.ftype == :directory
              out.put_next_entry(entry.name)
              if self.part_list.has?(entry.name)
                out.write self.part_list.get(entry.name).to_xml(indent: 0).gsub('\n', '')
              else
                out.write entry.get_input_stream.read
              end
            end
          end
        end
        buffer.seek(0)
        buffer
      end

      def merge_field(part, field, text)
        part.xpath(".//w:MergeField[text()=\"#{field}\"]").each do |merge_field|
          t_elem = Nokogiri::XML::Node.new('t', part)
          t_elem.content = text
          t_elem.parent = merge_field.parent
          merge_field.replace(t_elem)
        end
      end

      # replace the original convoluted tag with a simplified tag for easy searching and processing
      def process_merge_fields
        self.part_list.each_part do |part|
          part.root.remove_attribute('Ignorable')

          part.xpath('.//w:fldSimple/..').each do |parent|
            parent.children.each do |child|
              match_data = REGEXP.match(child.attribute('instr'))
              next if (child.node_name != 'fldSimple') || !match_data

              new_tag = Nokogiri::XML::Node.new('MergeField', part)
              new_tag.content = match_data[1]
              child.replace(new_tag)
            end
          end

          part.xpath('.//w:instrText/../..').each do |parent|
            begin_tags = parent.xpath('w:r/w:fldChar[@w:fldCharType="begin"]/..')
            end_tags = parent.xpath('w:r/w:fldChar[@w:fldCharType="end"]/..')
            instr_tags = parent.xpath('w:r/w:instrText')
            instr_tag_content = instr_tags.map(&:content)

            instr_tag_content.take(begin_tags.length).each_with_index do |instr, idx|
              next unless match_data = REGEXP.match(instr)

              new_tag = Nokogiri::XML::Node.new('MergeField', part)
              new_tag.content = match_data[1]

              children = parent.children
              start_idx = children.index(begin_tags[idx]) + 1
              end_idx = children.index(end_tags[idx])
              children[start_idx..end_idx].each do |child|
                instr_node = child.xpath('w:instrText')
                if instr_node.empty?
                  child.remove
                else
                  instr_node.first.replace(new_tag)
                end
              end
            end
          end
        end
      end
  end
end
