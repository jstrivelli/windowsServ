#---------------------------------------------------
#                Use Overview
#
# In a separate module, define a derived 
# PipedService class. Then, upon the module entry
# point, instantiate a WinServiceManager object, 
# passing it a reference to your service class,
# and invoke the manager's dispatch() function.
# That MUST be included for Windows to properly
# start the service when installed in a standalone
# context (i.e. without pythonservice.exe involved).
# Optionally, manage the service elsewhere in 
# your module via a WinServiceManager object.
#
# Example:
#
# class MyService( PipedService ):  
#   ... FILL IN THE DETAILS ...
#
# mgr = WinServiceManager( MyService, "MyService.exe" )
# mgr.start( autoInstall=True )
#
# if __name__ == "__main__":   
#    WinServiceManager( MyService, "MyService.exe" ).dispatch()
#
#---------------------------------------------------
#
# Here's a great source for ways of expanding and 
# improving upon this:
# http://www.icodeguru.com/WebServer/Python-Programming-on-Win32/ch18.htm
#
#---------------------------------------------------

# this project
from pipeServer import PipeServer
# pip mods
import win32api
import win32service  
import win32serviceutil  
import win32event  
import win32pipe
import win32file
import pywintypes
import winerror 
import servicemanager
# standard mods
import os, sys, time

SUCCESS = winerror.ERROR_SUCCESS
FAILURE = -1

MAX_STATUS_CHANGE_CHECKS  = 20
STATUS_CHANGE_CHECK_DELAY = 0.5

WORKING_DIR_OPT_NAME = "workingDir"           

# ABSTRACT CLASS
# =============================================================
class PipedService( win32serviceutil.ServiceFramework, PipeServer ):  

# Parts to override...
# -------------------------------------------------------------  
    # Basic properties (Required by ServiceFramework)
    """ OVERRIDE THESE """
    _svc_name_          = "MyService"  
    _svc_display_name_  = "My Custom Service"  
    _svc_description_   = "The purpose of this service is unknown."  
    _svc_is_auto_start_ = False
    
    def __init__( self, *args ):
        """ OVERRIDE THIS AS NEEDED (BUT CALL IT FIRST!) """
        win32serviceutil.ServiceFramework.__init__( self, *args )
        PipeServer.__init__( self
            , id=PipedService._svc_name_
            , friendlyPipeName=None
            , isDebugLogging=False
            , timeoutMillis=None
            , maxConsecutiveCommErrors=None
            , shutdownRequest=None               
        )
        # get basic pipe event
        self._hEvent_ = PipeServer._hEvent( self )       
        # create an event to listen for stop requests         
        self._hWaitStop_ = win32event.CreateEvent( None, 0, 0, None )          
        # Read working directory from registry and change to it
        servicename = args[0][0]
        workingDir = win32serviceutil.GetServiceCustomOption( 
            servicename, WORKING_DIR_OPT_NAME )
        PipeServer.taggedLog( self, "workingDir: %s" % workingDir )             
        os.chdir( workingDir )
        self.taggedLog( "service initialized" )             
    
    """ OVERRIDE THESE """    
    #def onStart( self ): 
        
    #def onStop( self ): 
         
    #def onRequest( self, request ): 
        
    def log( self, msg ):
        """ OVERRIDE THIS AS NEEDED """
        self._serviceLog( msg )

    def onUnCaughtException( self, exception ):        
        """ OVERRIDE THIS AS NEEDED """
        PipeServer.exceptionLog( self, exception )
        self.stop() # might want to ignore or restart instead?
        
# Helpers to call in your derived class        
# -------------------------------------------------------------  

    def enableDebugLog( self, isEnabled=True ) : 
        PipeServer.enableDebugLog( self, isEnabled )
    
    def _serviceLog( self, msg ):        
        servicemanager.LogInfoMsg( str(msg) )
        # Note service log messages can be seen most easily via:
        # pythonservice.exe -debug [service name] 
        # Note: service name = _svc_name_, and the service must 
        # be installed as a script (not exe). Alternatively, in 
        # a standalone exe context these messages can seen in 
        # the Windows Event Viewer within Windows Logs...Application.
    
# PipedService engine core   
# -------------------------------------------------------------  
    
    # Required by ServiceFramework 
    def SvcDoRun( self ): self.start() 
    def SvcStop( self ): self.stop()
    
    def start( self ):
        try:
            self.ReportServiceStatus( win32service.SERVICE_START_PENDING )         
            if not self._prepareToServe() : return            
            self.ReportServiceStatus( win32service.SERVICE_RUNNING )     
            PipeServer._enterRequestLoop( self )            
        except Exception as e: self.onUnCaughtException( e )
        
    def stop( self ) :
        self.ReportServiceStatus( win32service.SERVICE_STOP_PENDING )        
        win32event.SetEvent( self._hWaitStop_ )
        PipeServer._shutdown( self )        
        self.ReportServiceStatus( win32service.SERVICE_STOPPED )

# Overridden Protected Functions
# -------------------------------------------------------------          
    def _waitForEvents( self ) :
        waitResult = win32event.WaitForMultipleObjects( 
            [self._hWaitStop_, self._hEvent_], 
            0, win32event.INFINITE )            
        PipeServer.debugLog( self, "event detected" )            
        if waitResult == win32event.WAIT_OBJECT_0: 
            PipeServer.taggedLog( self, "encountered service stop event" )
            return False
        return True

# SERVICE MANAGEMENT       
# =============================================================
class WinServiceManager():  

    # pass the class, not an instance of it!
    def __init__( self, serviceClass, serviceExeName=None ):
        self.serviceClass_ = serviceClass
        # Added for pyInstaller v3
        self.serviceExeName_ = serviceExeName

    def isStandAloneContext( self ) : 
        # Changed for pyInstaller v3
        #return sys.argv[0].endswith( ".exe" ) 
        return not( sys.argv[0].endswith( ".py" ) )

    def dispatch( self ):
        if self.isStandAloneContext() :
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle( self.serviceClass_ )
            servicemanager.Initialize( self.serviceClass_._svc_name_, 
                os.path.abspath( servicemanager.__file__ ) )
            servicemanager.StartServiceCtrlDispatcher()        
        else :
            win32api.SetConsoleCtrlHandler(lambda x: True, True)  
            win32serviceutil.HandleCommandLine( self.serviceClass_ )        

# Service management functions
#            
# Note: all of these functions return:
# SUCCESS when explicitly successful
# FAILURE when explicitly not successful at their specific purpose
# winerror.XXXXXX when win32service (or related class) 
# throws an error of that nature
#------------------------------------------------------------------------

    # Note: an "auto start" service is not auto started upon installation!
    # To install and start simultaneously, use start( autoInstall=True ).
    # That performs both actions for manual start services as well.
    def install( self ):
        win32api.SetConsoleCtrlHandler(lambda x: True, True)        
        result = self.verifyInstall()
        if result == SUCCESS or result != FAILURE: return result
        thisExePath = os.path.realpath( sys.argv[0] )
        thisExeDir  = os.path.dirname( thisExePath )                
        # Changed for pyInstaller v3 - which now incorrectly reports the calling exe
        # as the serviceModPath (v2 worked correctly!)
        if self.isStandAloneContext() :
            serviceModPath = self.serviceExeName_
        else :
            serviceModPath = sys.modules[ self.serviceClass_.__module__ ].__file__        
        serviceModPath = os.path.splitext(os.path.abspath( serviceModPath ))[0] 
        serviceClassPath = "%s.%s" % ( serviceModPath, self.serviceClass_.__name__ )
        self.serviceClass_._svc_reg_class_ = serviceClassPath
        # Note: in a "stand alone context", a dedicated service exe is expected 
        # within this directory (important for cases where a separate master exe 
        # is managing services).  
        serviceExePath = (serviceModPath + ".exe") if self.isStandAloneContext() else None        
        isAutoStart = self.serviceClass_._svc_is_auto_start_
        startOpt = (win32service.SERVICE_AUTO_START if isAutoStart else 
                    win32service.SERVICE_DEMAND_START)        
        try :      
            win32serviceutil.InstallService(
                pythonClassString = self.serviceClass_._svc_reg_class_,
                serviceName       = self.serviceClass_._svc_name_,
                displayName       = self.serviceClass_._svc_display_name_,
                description       = self.serviceClass_._svc_description_,
                exeName           = serviceExePath,
                startType         = startOpt
            ) 
        except win32service.error as e: return e[0]
        except Exception as e: raise e        
        win32serviceutil.SetServiceCustomOption( 
            self.serviceClass_._svc_name_, WORKING_DIR_OPT_NAME, thisExeDir )
        for i in range( 0, MAX_STATUS_CHANGE_CHECKS ) :
            result = self.verifyInstall()
            if result == SUCCESS: return SUCCESS
            time.sleep( STATUS_CHANGE_CHECK_DELAY )            
        return result       

    def remove( self ):
        result = self.verifyInstall()
        if result == FAILURE : return SUCCESS
        if result != SUCCESS : return result
        result = self.verifyRunning()
        if result != SUCCESS and result != FAILURE: return result
        if result == SUCCESS :
            result = self.stop()
            if result != SUCCESS : return result    
        try : win32serviceutil.RemoveService( self.serviceClass_._svc_name_ )  
        except win32service.error as e: return e[0]
        except Exception as e: raise e        
        for i in range( 0, MAX_STATUS_CHANGE_CHECKS ) :
            result = self.verifyInstall()
            if result == FAILURE: return SUCCESS
            time.sleep( STATUS_CHANGE_CHECK_DELAY )            
        return result
        
    def start( self, autoInstall=False ):    
        if self.verifyRunning() == SUCCESS : return SUCCESS
        if self.verifyInstall() != SUCCESS :
            if not autoInstall : return FAILURE      
            result = self.install() 
            if result != SUCCESS: return result
        try : win32serviceutil.StartService( self.serviceClass_._svc_name_ )    
        except win32service.error as e: return e[0]
        except Exception as e: raise e       
        for i in range( 0, MAX_STATUS_CHANGE_CHECKS ) :
            result = self.verifyRunning()
            if result == SUCCESS : return SUCCESS
            time.sleep( STATUS_CHANGE_CHECK_DELAY )            
        return result

    def stop( self ):  
        if self.verifyRunning() == FAILURE: return SUCCESS
        result = self.verifyInstall()
        if result != SUCCESS : return result    
        try : win32serviceutil.StopService( self.serviceClass_._svc_name_ )  
        except win32service.error as e: return e[0]
        except Exception as e: raise e       
        for i in range( 0, MAX_STATUS_CHANGE_CHECKS ) :
            result = self.verifyRunning()
            if result == FAILURE: return SUCCESS
            time.sleep( STATUS_CHANGE_CHECK_DELAY )            
        return result

    def restart( self, autoInstall=False ):   
        result = self.stop()
        if result != SUCCESS and result != FAILURE: return result
        return self.start( autoInstall )

# Service status functions
#------------------------------------------------------------------------
        
    def verifyInstall( self ):  
        try: win32serviceutil.QueryServiceStatus( self.serviceClass_._svc_name_ )
        except win32service.error as e:
            return FAILURE if e[0] == winerror.ERROR_SERVICE_DOES_NOT_EXIST else e[0]    
        except Exception as e: raise e
        return SUCCESS
        
    def verifyRunning( self ): 
        svcState = None
        try :
            scvType, svcState, svcControls, err, svcErr, svcCP, svcWH = (
                win32serviceutil.QueryServiceStatus( self.serviceClass_._svc_name_) )
        except win32service.error as e: return e[0]
        except Exception as e: raise e                   
        return SUCCESS if (svcState == win32service.SERVICE_RUNNING) else FAILURE

# Describe service status, request results, and errors      
#------------------------------------------------------------------------

    def isInstalled( self ) : 
        try : return self.verifyInstall() == SUCCESS
        except: return False

    def isRunning( self ) : 
        try : return self.verifyRunning() == SUCCESS
        except: return False
        
    def statusDescr( self ):
        name = self.serviceClass_._svc_name_
        result = self.verifyInstall()
        if result == FAILURE: return "%s is not installed" % name
        if result == SUCCESS: 
            result = self.verifyRunning()            
            if result == SUCCESS: return "%s is running" % name
            if result == FAILURE: return "%s is installed, but not running" % name
        return "%s status is UNKNOWN" % name
        
    def resultDescr( self, requestDescr, result ) :
        if   result == SUCCESS: resultDescr = "succeeded" 
        elif result == FAILURE: resultDescr = "failed (no further details)"
        else : resultDescr = ("failed. error: %s (%d)" % 
                ( self.errorName( result ), result ) )
        return "PipedService request '%s' %s" % (requestDescr, resultDescr)
        
    def errorName( self, errorCode ) :
        if errorCode == SUCCESS: return "SUCCESS"
        if errorCode == FAILURE: return "FAILURE"
        import winerrorNames
        return winerrorNames.getName( errorCode )
