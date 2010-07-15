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

    it 'raises RuntimeError if event does not exist' do
      lambda { WIN32OLE_EVENT.new(@ole, 'A') }.should raise_error RuntimeError
    end

    ruby_version_is "1.9" do
      # this appears to cause seg fault in MRI 1.8.7
      it 'raises RuntimeError if OLE object has no events' do
        dict = WIN32OLE.new('Scripting.Dictionary')
        lambda { WIN32OLE_EVENT.new(dict) }.should raise_error RuntimeError
      end
    end

    it 'creates WIN32OLE_EVENT object' do
      ev = WIN32OLE_EVENT.new(@ole, 'DWebBrowserEvents')
      ev.should be_kind_of WIN32OLE_EVENT
    end

  end

end