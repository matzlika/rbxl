module Rbxl
  # Streams +xl/sharedStrings.xml+ out of an opened +.xlsx+ ZIP and decodes
  # the table to an immutable +Array<String>+.
  #
  # Both the read-only and edit modes need this same view of the SST. The
  # logic is identical — phonetic guides are skipped, +<r>+/+<t>+ runs inside
  # an +<si>+ are concatenated, the count and byte caps configured on
  # {Rbxl} are enforced — so it lives here as a single source of truth
  # rather than being inlined twice.
  #
  # @api private
  module SharedStringsLoader
    module_function

    # @param zip [Zip::File] the open package
    # @return [Array<String>] frozen, index-aligned shared strings table
    # @raise [Rbxl::SharedStringsTooLargeError] if the table exceeds the
    #   configured count or byte limits
    def load(zip)
      entry = zip.find_entry("xl/sharedStrings.xml")
      return [].freeze unless entry

      max_count = Rbxl.max_shared_strings
      max_bytes = Rbxl.max_shared_string_bytes

      # Reject zip-bomb style entries up front using the ZIP directory's
      # declared uncompressed size, before allocating any decompression buffer.
      if max_bytes && entry.size && entry.size > max_bytes
        raise SharedStringsTooLargeError,
              "shared strings uncompressed size #{entry.size} exceeds limit #{max_bytes}"
      end

      strings = []
      total_bytes = 0
      io = entry.get_input_stream
      reader = Nokogiri::XML::Reader(io)

      in_si = false
      in_run = false
      in_phonetic = false
      collecting_text = false
      buffer = +""
      current_fragments = []

      reader.each do |node|
        case node.node_type
        when Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.local_name
          when "si"
            in_si = true
            current_fragments = []
          when "r"
            in_run = true if in_si
          when "rPh"
            in_phonetic = true if in_si
          when "t"
            next unless in_si && !in_phonetic

            collecting_text = !in_run || node.depth.positive?
            buffer.clear if collecting_text
          end
        when Nokogiri::XML::Reader::TYPE_TEXT, Nokogiri::XML::Reader::TYPE_CDATA
          buffer << node.value if collecting_text
        when Nokogiri::XML::Reader::TYPE_END_ELEMENT
          case node.local_name
          when "t"
            if collecting_text
              current_fragments << buffer.dup
              collecting_text = false
            end
          when "r"
            in_run = false
          when "rPh"
            in_phonetic = false
          when "si"
            value = current_fragments.join.freeze
            total_bytes += value.bytesize
            if max_bytes && total_bytes > max_bytes
              raise SharedStringsTooLargeError,
                    "shared strings total size exceeds limit #{max_bytes}"
            end
            strings << value
            if max_count && strings.size > max_count
              raise SharedStringsTooLargeError,
                    "shared strings count exceeds limit #{max_count}"
            end
            in_si = false
            in_run = false
            in_phonetic = false
            collecting_text = false
          end
        end
      end

      strings.freeze
    ensure
      io&.close
    end
  end
end
