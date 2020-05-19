I developed the Python file in this directory using Microsoft Visual Studio and so you are advised to open the solution file in the directory above.

### Registering the COM servers
The code needs to be execute COM registration code; to do this ensure the final line of code RegisterCOMServers() is uncommented and run the code.  You will need administrator privileges to register the COM servers.

Once registered the code is run via COM activation logic such as `VBA.CreateObject("PythonInVBA.PythonVBAWebserver")`

Once registered you may comment the line RegisterCOMServers() as it is not required anymore.

### Code explanation
A blog post and video demonstration of this little project are planned.
