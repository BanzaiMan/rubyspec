platform_is :windows do
  
  describe "WIN32OLE#ole_get_methods" do
    
    before :each do
      @win32ole = WIN32OLE.new('InternetExplorer.application')
    end
    
    it "returns an array of WIN32OLE_METHOD objects" do
      @win32ole.ole_get_methods.each do |m|
        m.should be_kind_of WIN32OLE_METHOD
      end
    end
    
  end
  
end