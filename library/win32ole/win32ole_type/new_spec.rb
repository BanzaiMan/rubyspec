platform_is :windows do
  require 'win32ole'

  describe 'WIN32OLE_TYPE.new' do
    before :each do
      @ole = WIN32OLE.new('InternetExplorer.Application')
      @event = ''
    end

    after :each do
      @ole = nil
    end
    
    it 'raises ArgumentError with no argument' do
      lambda { WIN32OLE_TYPE.new }.should raise_error ArgumentError
    end
    
    it 'raises ArgumentError with invalid string' do
      lambda { WIN32OLE_TYPE.new("foo") }.should raise_error ArgumentError
    end
    
    it 'raises TypeError if second argument is not a String' do
      lambda { WIN32OLE_TYPE.new(1,2) }.should raise_error TypeError
      lambda { WIN32OLE_TYPE.new('Microsoft Shell Controls And Automation',2) }.
        should raise_error TypeError
    end
    
    it 'raise WIN32OLERuntimeError if OLE object specified is not found' do
      lambda { WIN32OLE_TYPE.new('Microsoft Shell Controls And Automation','foo') }.
        should raise_error WIN32OLERuntimeError
      lambda { WIN32OLE_TYPE.new('Microsoft Shell Controls And Automation','Application') }.
        should raise_error WIN32OLERuntimeError
    end
    
    it 'creates a WIN32OLE_TYPE object with valid arguments' do
      ole_type = WIN32OLE_TYPE.new("Microsoft Shell Controls And Automation", "Shell")
      ole_type.should be_kind_of WIN32OLE_TYPE
    end

  end
end
