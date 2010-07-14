platform_is :windows do
  require 'win32ole'
  
  describe 'WIN32OLE_EVENT.new' do
    before :each do
      @ole = WIN32OLE.new('InternetExplorer.Application')
      @event = ''
    end
    
    after :each do
      @ole = nil
    end
  
    it 'raises TypeError given invalid argument' do
      lambda { WIN32OLE_EVENT.new "A" }.should raise_error TypeError
    end
    
    it 'creates WIN32OLE_EVENT object' do
      ev = WIN32OLE_EVENT.new(@ole, 'DWebBrowserEvents')
      ev.should be_kind_of WIN32OLE_EVENT
    end
    
  end
  
end