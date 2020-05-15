import sys
import time #sleep
import http.server
import threading
import tempfile
import os

import win32com.client
from io import BytesIO
import pythoncom

import urllib.request


from http.server import HTTPServer, BaseHTTPRequestHandler, SimpleHTTPRequestHandler
import logging

## Global Variables
callbackInfo = None  # create a global instance

class DispatchEnsurer(object):
    @staticmethod
    def EnsureDispatch(comObj):
        """ Sometimes we get a PyIDispatch so we'll need to wrap it, this class takes care of that contingency """
        try:
            dispApp = None
            apptypename = str(type(comObj))
            if apptypename == "<class 'win32com.client.CDispatch'>":
                # this call from GetObject so no need to Dispatch()
                dispApp = comObj
            elif apptypename == "<class 'PyIDispatch'>":
                # this was passed in from VBA so wrap in Dispatch
                dispApp = win32com.client.Dispatch(comObj)
            else:
                # other cases just attempt to wrap
                dispApp = win32com.client.Dispatch(comObj)
            return dispApp 
        except Exception as ex:
            if hasattr(ex,"message"):
                return "Error:" + ex.message 
            else:
                return "Error:" + str(ex)



#https://stackoverflow.com/questions/22863646/marshaling-com-objects-between-python-processes-using-pythoncom
class CallbackInfo(object):
    """ One global object to contain many (otherwise separately global) variables relating to the callback"""

    def __init__(self,excelApplication, appRunGet: str, appRunPost: str):
        try:
            self.excelApplication = excelApplication
            self.appRunGet = appRunGet
            self.appRunPost = appRunPost
        except Exception as ex:
            logging.info("CallbackInfo.__init__   error   : " + 
                 LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))

    def __del__(self):
        try:
            self.excelApplication = None
            delattr(self,'excelApplication')
            delattr(self,'appRunGet')
            delattr(self,'appRunPost')
        except Exception as ex:
            logging.info("CallbackInfo.__del__  error : " + 
                 LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))


    def GetExcelApplication(self):
        try:
            return DispatchEnsurer.EnsureDispatch(self.excelApplication)
        except Exception as ex:
            logging.info("CallbackInfo.GetExcelApplication  error    : " + 
                 LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))


    def MakeCallBackGet(self,*args, **kwargs):
        try:
            dispXlApp = self.GetExcelApplication()
            
            if len(args) == 1:
                logging.info("CallbackInfo.MakeCallBackGet         : about to ApplicationRun '" + self.appRunGet + "' with arg '" + args[0] + "'")
                return dispXlApp.Run(self.appRunGet,args[0])
            if len(args) == 2:
                logging.info("CallbackInfo.MakeCallBackGet         : about to ApplicationRun '" + self.appRunGet + "' with arg '" + args[0] + "','" + str(args[1]) + "'")
                return dispXlApp.Run(self.appRunGet,args[0],args[1])
        except Exception as ex:
            logging.info("CallbackInfo.MakeCallBackGet  error    : " + 
                 LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))

    def MakeCallBackPost(self,*args, **kwargs):
        try:
            dispXlApp = self.GetExcelApplication()

            if len(args) == 1:
                logging.info("CallbackInfo.MakeCallBackPost        : about to ApplicationRun '" + self.appRunPost + "' with arg '" + args[0] + "'")
                return dispXlApp.Run(self.appRunPost,args[0])
            if len(args) == 2:
                logging.info("CallbackInfo.MakeCallBackPost        : about to ApplicationRun '" + self.appRunPost + "' with arg '" + args[0] + "','" + str(args[1]) + "'")
                return dispXlApp.Run(self.appRunPost,args[0],args[1])
        except Exception as ex:
            logging.info("CallbackInfo.MakeCallBackPost error : " + 
                 LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))



class MarshalledCallbackInfo(CallbackInfo):
    """ One global object to contain many (otherwise separately global) variables relating to the callback"""

    def __init__(self,excelApplication, appRunGet: str, appRunPost: str):
        try:
            dispXlApp = DispatchEnsurer.EnsureDispatch(excelApplication)
            logging.info("MarshalledCallbackInfo.__init__      : creating Marshalled interface pointer" )

            self.myStream = pythoncom.CreateStreamOnHGlobal()    
            pythoncom.CoMarshalInterface(self.myStream, 
                             pythoncom.IID_IDispatch, 
                             dispXlApp._oleobj_, 
                             pythoncom.MSHCTX_LOCAL, 
                             pythoncom.MSHLFLAGS_TABLESTRONG) 
            self.excelApplication = excelApplication
            self.appRunGet = appRunGet
            self.appRunPost = appRunPost
        except Exception as ex:
            logging.info("MarshalledCallbackInfo.__init__ error : " + 
                 LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))
            raise 

    def __del__(self):
        try:
            super(MarshalledCallbackInfo,self).__del__()

            logging.info("MarshalledCallbackInfo.__del__       : releasing Marshalled interface pointer" )

            # Clear the stream now that we have finished
            self.myStream.Seek(0,0)
            pythoncom.CoReleaseMarshalData(self.myStream)
            self.myStream = None
            delattr(self,'myStream')
        except Exception as ex:
            logging.info("MarshalledCallbackInfo.__del__   error : " + 
                 LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))
            raise 


    def GetExcelApplication(self):
        try:
            self.myStream.Seek(0,0)
            myUnmarshaledInterface = pythoncom.CoUnmarshalInterface(self.myStream, pythoncom.IID_IDispatch)    
            unmarshalledExcelApp = win32com.client.Dispatch(myUnmarshaledInterface)
            return unmarshalledExcelApp
        except Exception as ex:
            logging.info("MarshalledCallbackInfo.GetExcelApplication  error    : " + 
                 LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))
            raise 


class VBACallbackRequestHandler(SimpleHTTPRequestHandler):

    def do_OPTIONS(self):
        try:
            logging.info("MyRequestHandler.do_OPTIONS          : entered.  path=" + self.path)
            self.send_response(200)
            self.send_header("Access-Control-Allow-Origin", "*");
            self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS, GET');
            self.send_header("Access-Control-Allow-Headers", "Origin, X-Requested-With, content-type, Accept");
            self.send_header('Content-type', 'text/html')
            self.end_headers()
        except Exception as ex:
            logging.info("MyRequestHandler.do_OPTIONS error : " + 
                LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))


    def do_GET(self):
        try:
            logging.info("VBACallbackRequestHandler.do_GET     : entered.  path=" + self.path)
            quitting = True if self.path.find("quit")>-1 else False 
            if (quitting):
                logging.info("VBACallbackRequestHandler.do_GET     : quit found in path, quitting after this request ")
                self.server.stop = True

            self.send_response(200)
            self.send_header("Access-Control-Allow-Origin", "*");
            self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS, GET');
            self.send_header("Access-Control-Allow-Headers", "Origin, X-Requested-With, content-type, Accept");
            self.send_header('Content-type', 'text/html')
            self.end_headers()

            if (self.path != r"/favicon.ico"):
                self.wfile.write("GET request for {}".format(self.path).encode('utf-8'))
                global callbackInfo
                userAgent = "" if self.headers["User-Agent"] is None else self.headers["User-Agent"]
                get = "default response"
                if not quitting:
                    logging.info("VBACallbackRequestHandler.do_GET     : about call callbackInfo.MakeCallBackGet ")
                    get = callbackInfo.MakeCallBackGet(self.path, userAgent)
                if get is not None:
                    self.wfile.write((" " + get).encode('utf-8'))
                else:
                    self.wfile.write(" get callback returned none".encode('utf-8'))
        except Exception as ex:
            logging.info("VBACallbackRequestHandler.do_GET   error   : " + 
                LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))

    def do_POST(self):
        try:
            logging.info("VBACallbackRequestHandler.do_POST    : entered ")
            content_length = int(self.headers['Content-Length']) # <--- Gets the size of data
            post_data = self.rfile.read(content_length) # <--- Gets the data itself

            self.send_response(200)
            self.send_header("Access-Control-Allow-Origin", "*");
            self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS, GET');
            self.send_header("Access-Control-Allow-Headers", "Origin, X-Requested-With, content-type, Accept");
            self.send_header('Content-type', 'text/html')
            self.end_headers()

            msgBytesReceived = "POST body:" + str(len(post_data)) + " bytes received" 

            response = BytesIO()
            response.write(msgBytesReceived.encode('utf-8'))

            self.wfile.write(response.getvalue())
            self.wfile.flush()

            print(msgBytesReceived)

            logging.info("VBACallbackRequestHandler.do_POST    : about call callbackInfo.MakeCallBackPost ")
            global callbackInfo
            callbackInfo.MakeCallBackPost(self.path, post_data)
            logging.info("VBACallbackRequestHandler.do_POST    : " + msgBytesReceived)

        except Exception as ex:
            logging.info("VBACallbackRequestHandler.do_POST  error    : " + 
                LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))

    def do_QUIT (self):
        try:
            logging.info("VBACallbackRequestHandler.do_QUIT    : entered")
            """send 200 OK response, and set server.stop to True"""
            self.send_response(200)
            self.end_headers()
            logging.info("VBACallbackRequestHandler.do_QUIT             : setting self.server.stop = True")
            self.server.stop = True
            self.wfile.write("quit called".encode('utf-8'))
        except Exception as ex:
            logging.info("VBACallbackRequestHandler.do_QUIT  error    : " + 
                LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))

class LocalsEnhancedErrorMessager(object):
    @staticmethod
    def Enhance(ex, localsString):
        locals2 = "\n Locals:{ " + (",\n".join(localsString[1:-1].split(","))) + " }"
        if hasattr(ex,"message"):
            return "Error:" + ex.message + locals2
        else:
            return "Error:" + str(ex) + locals2

def thread_function(webserver):
    try:
        pythoncom.CoInitialize() # need this to tell the COM runtime that a new thread exists
        webserver.running = True 

        ## we need to pipe output to a file because whilst running as COM server there is no longer a console window to print to
        buffer = 1
        sys.stderr = open((os.path.dirname(os.path.realpath(__file__))) + '\\logfile.txt', 'w', buffer)
        sys.stdout = open((os.path.dirname(os.path.realpath(__file__))) + '\\logfile.txt', 'w', buffer)

        logging.info("thread_function                      : about to enter webserver.httpd.serve_forever")
        webserver.httpd.serve_forever()  #code enters into the subclass's implementation, an almost infinite loop

        logging.info("thread_function                      : returned from webserver.httpd.serve_forever")
        
        logging.info("thread_function                      : finished")

    except Exception as ex:
        logging.info("thread_function   error   : " + 
            LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))

class StoppableHttpServer(HTTPServer):
    # http://code.activestate.com/recipes/336012-stoppable-http-server/ 
    """http server that reacts to self.stop flag"""

    def serve_forever (self):
        try:
            logging.info("StoppableHttpServer.serve_forever    : entered")
            """Handle one request at a time until stopped."""
            self.stop = False
            while not self.stop:
                self.handle_request()
                logging.info("StoppableHttpServer.serve_forever    : request successfully handled self.stop=" + str(self.stop))

            logging.info("StoppableHttpServer.serve_forever    : dropped out of the loop")
        except Exception as ex:
            logging.info("StoppableHttpServer.serve_forever  error   : " + 
                LocalsEnhancedErrorMessager.Enhance(ex,str(locals())))


class PythonVBAWebserver(object):
    import logging
    import threading
    import time

    _reg_clsid_ = "{7719C990-6300-47B1-AEC7-F50F341199BE}"
    _reg_progid_ = 'PythonInVBA.PythonVBAWebserver'
    _public_methods_ = ['StartWebServer','StopWebServer','CheckThreadStatus','StopLogging']
    _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER

    def StopLogging(self):
        try:
            logging.shutdown()
            return "logging.shutdown() ran"
        except Exception as ex:
            msg = "PythonVBAWebserver.StopLogging error:" +  LocalsEnhancedErrorMessager.Enhance(ex,str(locals()))
            logging.info(msg)
            return msg

    def StartWebServer(self,excelApplication, appRunGet: str, appRunPost: str, server_name:str, server_port: int):
        try:
            self.server_name = server_name
            self.server_port = server_port

            logging.basicConfig(filename =  (os.path.dirname(os.path.realpath(__file__))) + '\\app2.log', format="%(asctime)s: %(message)s", level=logging.INFO,
                        datefmt="%H:%M:%S")
            logging.info("PythonVBAWebserver.StartWebServer    : before creating thread")

            global callbackInfo # this indicates that we are dealing with the global instance
            callbackInfo = MarshalledCallbackInfo(excelApplication, appRunGet, appRunPost)
            

            logging.info("PythonVBAWebserver.StartWebServer    : server_name: " + server_name + ", server_port:" + str(server_port))

            self.running = False 
            
            self.httpd = StoppableHttpServer((server_name, server_port), VBACallbackRequestHandler)
            #self.httpd = HTTPServer((server_name, server_port), VBACallbackRequestHandler)

            logging.info("PythonVBAWebserver.StartWebServer    : about to create thread")

            self.serverthread = threading.Thread(name="webserver", target=thread_function, args=(self,))
            #self.serverthread.setDaemon(True)
            logging.info("PythonVBAWebserver.StartWebServer    : about to start thread")

            self.serverthread.start()
            logging.info("PythonVBAWebserver.StartWebServer    : after call to start thread")
            
            return "StartWebServer ran ok ( server_name: " + server_name + ", server_port:" + str(server_port) + ")"

        except Exception as ex:
            msg = "PythonVBAWebserver.StartWebServer error:" +  LocalsEnhancedErrorMessager.Enhance(ex,str(locals()))
            logging.info(msg)
            return msg

    def CheckThreadStatus(self):
        try:
            # Clear the stream now that we have finished
            global callbackInfo

            if self.running:
                if hasattr(self,'httpd') :
                    logging.info("PythonVBAWebserver.CheckThreadStatus    : checking thread status")
                    return self.serverthread.is_alive()
                else:
                    return "StopWebServer ran ok, nothing to stop"
            else:
                return "StopWebServer ran ok, nothing to stop"

        except Exception as ex:
            msg = "PythonVBAWebserver.CheckThreadStatus error:" +  LocalsEnhancedErrorMessager.Enhance(ex,str(locals()))
            logging.info(msg)
            return msg

    def StopWebServer(self):
        try:
            retMsg = "StopWebServer ran (default)"
            logging.info("PythonVBAWebserver.StopWebServer     : entered")
            # Clear the stream now that we have finished
            global callbackInfo
            callbackInfo = None # release stream
            logging.info("PythonVBAWebserver.StopWebServer     : callbackInfo released")

            if self.running:
                if hasattr(self,'httpd') :


                    logging.info("PythonVBAWebserver.StopWebServer     : call quit on own web server")
                    ## make a quit request to our own server 
                    quitRequest  = urllib.request.Request("http://" + self.server_name + ":" + str(self.server_port) + "/quit",
                                                      method="QUIT")
                    with urllib.request.urlopen(quitRequest ) as resp:
                        logging.info("PythonVBAWebserver.StopWebServer     : quit response '" + resp.read().decode("utf-8") + "'")

                    # web server should have exited loop and its thread should be ready to terminate
                    logging.info("PythonVBAWebserver.StopWebServer     : about to join thread")
                    self.serverthread.join()    # get the server thread to die and join this thread
                    self.running = False 
                    
                    logging.info("StarterAndStopper.StopWebServer      : thread joined")

                    logging.info("StarterAndStopper.StopWebServer      : about to call httpd.server_close()")
                    self.httpd.server_close()  #now we can close the server cleanly

                    logging.info("StarterAndStopper.StopWebServer      : completed")
                    
                    retMsg = "StopWebServer ran ok, web server stopped"
                else:
                    retMsg = "StopWebServer ran ok, nothing to stop"
            else:
                retMsg = "StopWebServer ran ok, nothing to stop"
            return retMsg

        except Exception as ex:
            msg = "PythonVBAWebserver.StopWebServer error:" +  LocalsEnhancedErrorMessager.Enhance(ex,str(locals()))
            print(msg)
            logging.info(msg)
            return msg


def run():
    try:

        print("Executing run")
        print((os.path.dirname(os.path.realpath(__file__))))

        logging.basicConfig(filename = (os.path.dirname(os.path.realpath(__file__))) + '\\app2.log', format="%(asctime)s: %(message)s", level=logging.INFO,
                    datefmt="%H:%M:%S")


        # Get an Excel Application by requesting an excel workbook and taking
        # its parent

        #wbFullPath = r"C:\Users\Simon\Downloads\My recent code\PythonWebServerCallsbackToVBA_20200503_2105.xlsm" 
        wbFullPath = r"C:\Users\Simon\Downloads\My recent code\JobLeadsCache.xlsm" 
        

        logging.info("run()                                : attempt to GetObject " + wbFullPath)

        wb = win32com.client.GetObject(wbFullPath)

        logging.info("run()                                : attempt to get Excel Application object from " + wbFullPath)
        xlApp = wb.Parent
        xlApp.Visible = True 

        global callbackInfo
        callbackInfo = MarshalledCallbackInfo(xlApp,wb.Name + "!VBA_DO_GET", wb.Name + "!VBA_DO_POST")


        print("VBA_DO_GET:  " + xlApp.Run(wb.Name + "!VBA_DO_GET","",""))
        print("VBA_DO_POST: " + xlApp.Run(wb.Name + "!VBA_DO_POST","",""))
        

        ws = PythonVBAWebserver()
        ws.StartWebServer(xlApp,wb.Name + "!VBA_DO_GET", wb.Name + "!VBA_DO_POST",'localhost',80)

        logging.info('called PythonVBAWebserver.StartWebServer ...\n')

        if False:

            logging.info('what next? ...\n')
            ws.StopWebServer()

            logging.info('finishing run()\n')

    except Exception as ex:
        print(ex)


def RegisterCOMServers():
    print("Registering COM servers...")
    import win32com.server.register
    win32com.server.register.UseCommandLine(PythonVBAWebserver)

if __name__ == '__main__':
    #run()
    RegisterCOMServers()

