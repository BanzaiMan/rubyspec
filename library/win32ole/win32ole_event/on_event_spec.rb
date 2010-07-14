platform_is :windows do
  require 'win32ole'
  
  def default_handler(event, *args)
    @event += event
  end
  
  
  describe 'WIN32OLE_EVENT#on_event' do
    before :each do
      @ole = WIN32OLE.new('InternetExplorer.Application')
      @ev  = WIN32OLE_EVENT.new(@ole, 'DWebBrowserEvents')
      @event = ''
    end
    
    after :each do
      @ole = nil
    end
  
    it 'sets event handler properly, and the handler is invoked by event loop' do
      @ev.on_event {|*args| default_handler(*args)}
      @ole.StatusText='hello'
      WIN32OLE_EVENT.message_loop
      @event.should =~ /StatusTextChange/
    end
    
  end
  
end