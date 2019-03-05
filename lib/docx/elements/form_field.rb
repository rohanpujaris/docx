require 'docx/elements/element'

module Docx
  module Elements
    class FormField
      # probably included Element is not required here
      include Element

      def self.tag
        'ffData'
      end

      attr_reader :node, :ancestor_paragraph, :type, :name, :value

      def initialize(node)
        @node = node
        @ancestor_paragraph = node.xpath("ancestor::w:p").first
        @type = get_type_for_form_field
        @name = get_name_for_form_field
        @value = get_value_for_form_field
      end

      def get_type_for_form_field
        node_type = ancestor_paragraph.xpath("descendant::w:instrText").first.text.strip
        case node_type
        when "FORMTEXT"
          "text"
        when "FORMCHECKBOX"
          "checkbox"
        when "FORMDROPDOWN"
          "dropdown"
        end
      end

      def get_name_for_form_field
        node.xpath('descendant::w:name/@w:val').first&.content
      end

      def get_value_for_form_field
        case type
        when "text"
          get_value_for_text_field
        when "checkbox"
          get_value_for_checkbox_field
        when "dropdown"
          get_value_for_dropdown_field
        end
      end

      def get_value_for_text_field
        Elements::Containers::Paragraph.new(ancestor_paragraph).to_s
      end

      def get_value_for_checkbox_field
          # Cases of default checked checkboxes need to be handled here
        !ancestor_paragraph.xpath('descendant::w:r[1]/descendant::w:checked').empty?
      end

      def get_value_for_dropdown_field
      end

      def set_value_for_form_field(value)
        case type
        when "text"
          set_value_for_text_field(value)
        when "checkbox"
          set_value_for_checkbox_field(value)
        when "dropdown"
          set_value_for_dropdown_field(value)
        end
      end

      def set_value_for_text_field(value)
        ancestor_paragraph.xpath("descendant::w:t/parent::w:r").each_with_index do |node, index|
          if index == 0
            node.xpath('w:t').first.content = value
          else
            node.remove
          end
        end
      end

      def set_value_for_checkbox_field(value)
      end

      def set_value_for_dropdown_field(value)
      end

      def to_h
        { name: name, type: type, value: value }
      end
    end
  end
end

# Docx::Document.open('example1.docx').set_value_for_form_field("Name" => "LLL")