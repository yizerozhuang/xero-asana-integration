import tkinter
class TextExtension( tkinter.Frame ):
    """Extends Frame.  Intended as a container for a Text field.  Better related data handling
    and has Y scrollbar now."""
    def __init__( self, master, textvariable = None, *args, **kwargs ):
        self.textvariable = textvariable
        if ( textvariable is not None ):
            if not ( isinstance( textvariable, tkinter.Variable ) ):
                raise TypeError( "tkinter.Variable type expected, {} given.".format( type( textvariable ) ) )
            self.textvariable.get = self.GetText
            self.textvariable.set = self.SetText

        # build
        # self.YScrollbar = None
        # self.Text = None

        super().__init__( master )

        # self.YScrollbar = tkinter.Scrollbar( self, orient = tkinter.VERTICAL )

        self.Text = tkinter.Text( self, *args, **kwargs )
        # self.YScrollbar.config( command = self.Text.yview )
        # self.YScrollbar.pack( side = tkinter.RIGHT, fill = tkinter.Y )

        self.Text.pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=1)


    def Clear( self ):
        self.Text.delete( 1.0, tkinter.END )


    def GetText(self):
        text = self.Text.get(1.0, tkinter.END )
        if (text is not None ):
            text = text.strip()
        if (text == "" ):
            text = None
        return text


    def SetText( self, value ):
        self.Clear()
        if ( value is not None ):
            self.Text.insert( tkinter.END, value.strip() )