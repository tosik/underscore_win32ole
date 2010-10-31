require 'win32ole'
require 'active_support/inflector'

class WIN32OLE
  alias_method :old_method_missing, :method_missing
  def method_missing(name, *args, &block)
    begin
      old_method_missing(name.to_s.camelize(:upper), *args, &block)
    rescue NameError
      old_method_missing(name, *args, &block)
    end
  end
end
