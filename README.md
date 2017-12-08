# Leaptest - VSTS and TFS Integration  
This is Leaptest Integration extension for Visual Studio Team Services and Team Foundation Server 2015 Update 1 or later (minimum agent version 1.83.0)  
# More Details
  
Leaptest is a mighty automation testing system and now it can be used for running [smoke, functional, acceptance] tests, generating reports and a lot more in VSTS and TFS. 

# Features:
  
- Run automated tests in your VSTS/TFS build tasks
  
- Automatically receive test results
 
- Build status based tests results
  
- Generate a xml report file in JUnit format
 
- Write tests trace to build output log
 
# Installing  
Simply install the extension using VSIX-file or download the extension from Visual Studio Marketplace (available soon)
# Instruction
  
1. Add build step "Leaptest Inegration" to your build definition.

2. Enter Leaptest controller URL address something like http://{HOSTNAME where LEAPTEST controller is installed}:{port number (9000 is by default)}" or "http://localhost:9000".
3. Enter time delay in seconds. When schedule is run, extension will wait this time before trying to get schedule state. If schedule is still running, extension will wait this time again. By default this value is 3 seconds.  
4. Select how extension should set "Done" status value: to Success or Failed.  
5. Enter JUnit report file name. This file will be created at $(SourceDir) directory. If there is an xml file with the same name, it will be overwritten. By default it is "report.xml".  
6. (Optional) Enter schedule ids, each one must be entered from a new line. Using id is more reliable, because schedule title can be changed.
7. Enter schedule titles you want to run,  each one must be entered from a new line. 
8. Add Test build step "Publish Test Results" to your build definition. Choose JUnit report format. Enter JUnit report file name. It MUST be the same you've entered before!  
9. Run your job and get results.Enjoy!  
10. Look at build results in Detailed report in Build summary  
11. Download logs if required  


# Screenshots
  
![image](http://customatics.com/wp-content/uploads/2017/09/screen1.png)

![image](http://customatics.com/wp-content/uploads/2017/09/screen2.png)

![image](http://customatics.com/wp-content/uploads/2017/09/screen3.png)

