# debug for OpenOffice and LibreOffice

The BASIC programming language is embedded in each of the OpenOffice/LibreOffice suite of applications. BASIC includes three debugging facilities that reveal an objects properties, methods or supported interfaces. For example, with an object of *oDocument*, issue the commands:

```
msgbox oDocument.Dbg_Properties
msgbox oDocument.Dbg_Methods
msgbox oDocument.Dbg_SupportedInterfaces
```

Each command places their returned data into a Message dialog window. There are some issues with the Message dialog window that include:
* Unable to cut and paste its contents.
* Window may extend off the bottom of the display
* Returned data is not alphabetically sorted.

The file *debug.bas* contains a debug subroutine and sorting function. When writing a BASIC macro script for an OpenOffice/LibreOffice application add the two routines that are in *debug.bas*. Within the BASIC script place a call to the *debug()* routine. E.g.
```
debug(oDocument)
```
This *debug* sub-routine will:
* Call the Dbg_properties, Dbg_Method and Dbg_supportInterfaces routines and collect their returned data.
* Process the data to clean up its formatting. E.g. Remove excess spaces and newlines. Change case.
* Sort the data alphabetically.
* Write the data to a text file.
* In the case of a Linux based system with *pluma* or *gedit* it will open one of these text editors and display the data in the text file.

For example, when a push button is clicked, then it may call the subroutine:
```
sub ButtonPushEvent(ev as com.sun.star.awt.ActionEvent)
    ...
    debug(ev.source)
```
The processing by *debug* will produce an output file that displays in an editor like this:
```
Properties of object  "com.sun.star.form.OButtonControl":
ARRAY        ImplementationId
ARRAY        SupportedServiceNames
ARRAY        Types
BOOL         DesignMode
BOOL         Enable
BOOL         Visible
OBJECT       AccessibleContext
OBJECT       Context
OBJECT       Delegator
OBJECT       Graphics
OBJECT       MinimumSize
OBJECT       Model
OBJECT       OutputSize
OBJECT       Peer
OBJECT       PosSize
OBJECT       PreferredSize
OBJECT       Size
OBJECT       StyleSettings
OBJECT       View
STRING       ActionCommand
STRING       Dbg_Methods
STRING       Dbg_Properties
STRING       Dbg_SupportedInterfaces
STRING       ImplementationName
STRING       Label

Methods of object  "com.sun.star.form.OButtonControl":
ARRAY        getImplementationId (void)
ARRAY        getSupportedServiceNames (void)
ARRAY        getTypes (void)
BOOL         hasFocus (void)
BOOL         isActive (void)
BOOL         isDesignMode (void)
BOOL         isEnabled (void)
BOOL         isTransparent (void)
BOOL         isVisible (void)
BOOL         setGraphics (object)
BOOL         setModel (object)
BOOL         supportsService (string)
EMPTY        queryAggregation (object)
EMPTY        queryInterface (object)
OBJECT       calcAdjustedSize (object)
OBJECT       convertPointToLogic (object, integer)
OBJECT       convertPointToPixel (object, integer)
OBJECT       convertSizeToLogic (object, integer)
OBJECT       convertSizeToPixel (object, integer)
OBJECT       getAccessibleContext (void)
OBJECT       getContext (void)
OBJECT       getGraphics (void)
OBJECT       getMinimumSize (void)
OBJECT       getModel (void)
OBJECT       getOutputSize (void)
OBJECT       getPeer (void)
OBJECT       getPosSize (void)
OBJECT       getPreferredSize (void)
OBJECT       getSize (void)
OBJECT       getView (void)
OBJECT       queryAdapter (void)
STRING       getImplementationName (void)
VOID         actionPerformed (object)
VOID         addActionListener (object)
VOID         addApproveActionListener (object)
VOID         addEventListener (object)
VOID         addFocusListener (object)
VOID         addItemListener (object)
VOID         addKeyListener (object)
VOID         addModeChangeApproveListener (object)
VOID         addModeChangeListener (object)
VOID         addMouseListener (object)
VOID         addMouseMotionListener (object)
VOID         addPaintListener (object)
VOID         addSubmissionVetoListener (object)
VOID         addWindowListener (object)
VOID         createPeer (object, object)
VOID         dispose (void)
VOID         disposing (object)
VOID         draw (long, long)
VOID         itemStateChanged (object)
VOID         propertiesChange (array)
VOID         propertyChange (object)
VOID         registerDispatchProviderInterceptor (object)
VOID         releaseDispatchProviderInterceptor (object)
VOID         removeActionListener (object)
VOID         removeApproveActionListener (object)
VOID         removeEventListener (object)
VOID         removeFocusListener (object)
VOID         removeItemListener (object)
VOID         removeKeyListener (object)
VOID         removeModeChangeApproveListener (object)
VOID         removeModeChangeListener (object)
VOID         removeMouseListener (object)
VOID         removeMouseMotionListener (object)
VOID         removePaintListener (object)
VOID         removeSubmissionVetoListener (object)
VOID         removeWindowListener (object)
VOID         setActionCommand (string)
VOID         setContext (object)
VOID         setDelegator (object)
VOID         setDesignMode (bool)
VOID         setEnable (bool)
VOID         setFocus (void)
VOID         setLabel (string)
VOID         setOutputSize (object)
VOID         setPosSize (long, long, long, long, integer)
VOID         setVisible (bool)
VOID         setZoom (single, single)
VOID         statusChanged (object)
VOID         submit (void)
VOID         submitWithInteraction (object)

Supported interfaces by object  "com.sun.star.form.OButtonControl":
com.sun.star.accessibility.XAccessible
com.sun.star.awt.XActionListener
com.sun.star.awt.XButton
com.sun.star.awt.XControl
com.sun.star.awt.XItemEventBroadcaster
com.sun.star.awt.XItemListener
com.sun.star.awt.XLayoutConstrains
com.sun.star.awt.XStyleSettingsSupplier
com.sun.star.awt.XToggleButton
com.sun.star.awt.XUnitConversion
com.sun.star.awt.XView
com.sun.star.awt.XWindow
com.sun.star.awt.XWindow2
com.sun.star.beans.XPropertiesChangeListener
com.sun.star.beans.XPropertyChangeListener
com.sun.star.form.submission.XSubmission
com.sun.star.form.XApproveActionBroadcaster
com.sun.star.frame.XDispatchProviderInterception
com.sun.star.frame.XStatusListener
com.sun.star.lang.XComponent
com.sun.star.lang.XComponent
com.sun.star.lang.XComponent
com.sun.star.lang.XEventListener
com.sun.star.lang.XEventListener
com.sun.star.lang.XEventListener
com.sun.star.lang.XEventListener
com.sun.star.lang.XEventListener
com.sun.star.lang.XEventListener
com.sun.star.lang.XServiceInfo
com.sun.star.lang.XTypeProvider
com.sun.star.uno.XAggregation
com.sun.star.uno.XWeak
com.sun.star.util.XModeChangeBroadcaster
```


Note:

Tested with LibreOffice Version: 6.4.7.2 on Ubuntu Mate 20.04.2
