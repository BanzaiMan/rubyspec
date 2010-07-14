platform_is :windows do
  require 'win32ole'
  
  describe 'WIN32OLE.codepage' do
    it 'returns the current codepage' do
      WIN32OLE.codepage.should == WIN32OLE::CP_ACP # default value
    end
  end
  
  describe 'WIN32OLE.codepage=' do
    it 'sets codepage' do
      WIN32OLE.codepage = WIN32OLE::CP_UTF8
      WIN32OLE.codepage.should == WIN32OLE::CP_UTF8
    end
  end
  
end