module Rbxl
  # @api private
  #
  # DOM-insertion helpers for edit mode. Every fragment is parsed in the
  # context of its new parent and then moved into the parent's namespace, so
  # the inserted elements land in the right namespace whether the host
  # document binds SpreadsheetML as the default namespace (Excel, LibreOffice)
  # or under a prefix such as +x:+ (the .NET OpenXML SDK).
  module XmlEdit
    module_function

    # Serializes a document or node without libxml2's pretty-printing, so
    # a re-serialized part keeps the source's layout apart from the edits.
    #
    # @return [String]
    def serialize(node)
      node.to_xml(save_with: Nokogiri::XML::Node::SaveOptions::AS_XML)
    end

    # Appends +xml+ as the last child of +parent+.
    #
    # @return [Nokogiri::XML::Element] the first inserted element
    def append(parent, xml)
      adopt(parent, parent.add_child(xml))
    end

    # Inserts +xml+ immediately before +sibling+.
    #
    # @return [Nokogiri::XML::Element] the first inserted element
    def insert_before(sibling, xml)
      adopt(sibling.parent, sibling.add_previous_sibling(xml))
    end

    # Inserts +xml+ as a child of +parent+ respecting a schema sequence:
    # the new element goes before the first existing child whose local name
    # appears after +name+ in +order+, or at the end when there is none.
    #
    # @param order [Array<String>] local names in schema sequence order
    # @return [Nokogiri::XML::Element] the inserted element
    def insert_in_order(parent, name, xml, order)
      rank = order.index(name) or raise ArgumentError, "#{name} not in order list"
      following = parent.element_children.find do |child|
        idx = order.index(child.name)
        idx && idx > rank
      end
      following ? insert_before(following, xml) : append(parent, xml)
    end

    # Returns the first element child of +parent+ named +name+ in +parent+'s
    # namespace, or +nil+.
    def child(parent, name)
      href = parent.namespace&.href
      parent.element_children.find { |c| c.name == name && c.namespace&.href == href }
    end

    # Returns all element children of +parent+ named +name+ in +parent+'s
    # namespace.
    def children(parent, name)
      href = parent.namespace&.href
      parent.element_children.select { |c| c.name == name && c.namespace&.href == href }
    end

    def adopt(parent, inserted)
      ns = parent.namespace
      nodes = inserted.is_a?(Nokogiri::XML::NodeSet) ? inserted.to_a : [inserted]
      nodes.each { |node| set_namespace_deep(node, ns) } if ns
      nodes.find(&:element?)
    end

    def set_namespace_deep(node, ns)
      return unless node.element?

      node.namespace = ns if node.namespace.nil?
      node.element_children.each { |child| set_namespace_deep(child, ns) }
    end
  end
end
