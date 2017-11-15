require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element

        def self.tag
          'p'
        end


        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node, document_properties = {})
          @node = node
          @properties_tag = 'pPr'
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
        end

        # Set text of paragraph
        def text=(content)
          if text_runs.size == 1
            text_runs.first.text = content
          elsif text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            text_runs.each {|r| r.node.remove }
            new_r = TextRun.create_within(self)
            new_r.text = content
          end
        end

        # Return text of paragraph
        def to_s
          text_runs.map(&:text).join('')
        end

        # Return paragraph as a <p></p> HTML fragment with formatting based on properties.
        def to_html
          html = ''
          text_runs.each do |text_run|
            html << text_run.to_html
          end
          styles = { 'font-size' => "#{font_size}pt" }
          styles['text-align'] = alignment if alignment
          styles['color'] = "##{color}" if color
          if l_id = list_id
            list_start = @document_properties[:lists][l_id.to_s].to_i + 1
            @document_properties[:lists][l_id.to_s] = list_start
            html_tag(:ol,
                     content: html_tag(:li,
                                       content: html,
                                       styles: styles),
                     styles: {margin: '2px'},
                     attributes: "start=#{list_start}")
          else
            html_tag(:p, content: html, styles: styles)
          end
        end

        # Array of text runs contained within paragraph
        def text_runs
          @node.xpath('w:r|w:hyperlink/w:r|w:ins/w:r').map { |r_node| Containers::TextRun.new(r_node, @document_properties) }
        end

        # Iterate over each text run within a paragraph
        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end

        def aligned_left?
          ['left', nil].include?(alignment)
        end

        def aligned_right?
          alignment == 'right'
        end

        def aligned_center?
          alignment == 'center'
        end

        def font_size
          size_tag = @node.xpath('w:pPr//w:sz').first
          size_tag ? size_tag.attributes['val'].value.to_i / 2 : @font_size
        end
        
        alias_method :text, :to_s

        private

        # Returns the alignment if any, or nil if left
        def alignment
          alignment_tag = @node.xpath('.//w:jc').first
          alignment_tag ? alignment_tag.attributes['val'].value : nil
        end

        def list_id
          list_tag = @node.xpath('w:pPr//w:numPr//w:numId').first
          list_tag ? list_tag.attributes['val'].value : nil
        end

        def color
          color_tag = @node.xpath('w:pPr//w:color').first
          color_tag ? color_tag.attributes['val'].value : nil
        end

      end
    end
  end
end
