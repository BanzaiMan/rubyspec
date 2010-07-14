platform_is :windows do
  require 'win32ole'
  
  describe 'WIN32OLE.const_load with one argument' do
    before :each do
      @win32ole = WIN32OLE.new 'InternetExplorer.application'
    end
    
    it 'loads constants into WIN32OLE namespace' do
      WIN32OLE.const_load @win32ole
      WIN32OLE.constants.should_not be_empty
    end
  end
  
  describe 'WIN32OLE.const_load with two arguments' do
    before :each do
      module WIN32OLE_RUBYSPEC; end
      @win32ole = WIN32OLE.new 'InternetExplorer.application'
    end
    
    it 'loads constants into given namespace' do
      WIN32OLE.const_load @win32ole, WIN32OLE_RUBYSPEC
      WIN32OLE_RUBYSPEC.constants.should_not be_empty
    end
  end

end