platform_is :windows do
  require 'win32ole'
  
  describe 'WIN32OLE_METHOD#offset_vtbl' do
    before :each do
      ole_type = WIN32OLE_TYPE.new("Microsoft Scripting Runtime", "File")
      @m_file_name = WIN32OLE_METHOD.new(ole_type, "name")
    end
    
    it 'raises ArgumentError if argument is given' do
      lambda { @m_file_name.offset_vtbl(1) }.should raise_error ArgumentError
    end
    
    it 'returns expected value for Scripting Runtime\'s "File" method' do
      @m_file_name.offset_vtbl.should == 32
    end
    
  end
  
end