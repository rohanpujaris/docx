require 'docx/elements/element'
require "byebug"

module Docx
  module Elements
    class FormField
      # probably included Element is not required here
      include Element

      def self.tag
        'ffData'
      end

      attr_reader :node, :ancestor_paragraph, :type, :name, :value, :meta

      def initialize(node)
        @node = node
        @ancestor_paragraph = node.xpath("ancestor::w:p").first
        @type = get_type_for_form_field
        @name = get_name_for_form_field
        @meta = get_meta_value_for_form_field
        @value = get_value_for_form_field
      end

      def get_meta_value_for_form_field
        return {} if type != "dropdown"
        {
          options: node.xpath("descendant::w:listEntry").map {|n| n["w:val"] }
        }
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
        ancestor_paragraph.xpath("w:r/w:instrText/ancestor::w:r/following-sibling::w:r/w:t").map(&:text).join(" ")
      end

      def get_value_for_checkbox_field
        # Cases of default checked checkboxes need to be handled here
        !ancestor_paragraph.xpath('descendant::w:r[1]/descendant::w:checked').empty?
      end

      def get_value_for_dropdown_field
        selected_option = node.xpath("descendant::w:result").first
        if selected_option
          meta[:options][selected_option["w:val"].to_i]
        else
          meta[:options].first
        end
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
        # Should not delete all w:r node. Instead should delete it from bookmark start to bookmar end
        ancestor_paragraph.xpath("w:r/w:instrText/ancestor::w:r/following-sibling::w:r/w:t/parent::w:r").each_with_index do |node, index|
          if index == 0
            node.xpath('w:t').first.content = value
          else
            node.remove
          end
        end
      end

      def set_value_for_checkbox_field(checked)
        # Cases of default checked checkboxes need to be handled here
        if checked
          unless value
            new_node = Nokogiri::XML("<w:checked/>").root
            node.xpath("descendant::w:checkBox").first.add_child(new_node)
          end
        else
          node.xpath("descendant::w:checked").first&.remove
        end
      end

      def set_value_for_dropdown_field(value)
        index = meta[:options].index(value)
        if index
          result_node = node.xpath("descendant::w:result").first
          if result_node
            result_node["w:val"] = index
          else
            new_node = Nokogiri::XML("<w:result w:val='#{index}'/>").root
            node.xpath("descendant::w:ddList").first.add_child(new_node)
          end
        end
      end

      def to_h
        { name: name, type: type, value: value }
      end
    end
  end
end

# data = {
#   "first_name" => "Rohan",
#   "last_name" => "Pujari",
#   "english" => true,
#   "french" => false,
#   "spanish" => false,
#   "qualification" => "Graduate"
# }

# data = {
#   "first_name" => "Kavita",
#   "last_name" => "Joshi",
#   "english" => true,
#   "french" => false,
#   "spanish" => true,
#   "qualification" => "Post Graduate"
# }
# Docx::Document.open('fillable_form.docx').set_value_for_form_field(data, flatten: true)
# Docx::Document.open('fillable_form.docx').get_form_fields.map(&:to_h)