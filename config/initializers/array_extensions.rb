# config/initializers/string_extensions.rb

# Open the String class to add custom functionality
class Array
  # Override the flatten method to remove nil values and flatten nested arrays
  def self.flatten
    result = []
    self.each do |element|
      if element.is_a?(Array)
        result += element.flatten
      elsif !element.nil?
        result << element
      end
    end
    result
  end
end