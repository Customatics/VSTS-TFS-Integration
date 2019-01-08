# Leapwork - VSTS and TFS Integration  

# Joining forces of Leapwork and VSTS/TFS  
For over 15 years [Customatics](http://customatics.com) team have been delivering high-quality software development and IT service solutions. We think outside the box, sharing fresh ideas to make our products work for you. We are proud to say that the solution we developed [Leapwork automation platform](https://www.leapwork.com/) is now trusted by numerous customers around the word. Now we give you a chance to combine the power of Leapwork automation platform and VSTS/ TFS management systems. Customatics team have created Leapwork integration extension for VSTS/ TFS and you are welcome to benefit from our solution.

# Streamline your test automation  
A significant part of running automated cases is to monitor, inspect and react on the results. Integration with test management and bug tracking systems enables to you streamline your test automation and development cycles so that not a single flaw of the system is left unnoticed.  
Leapwork Automation Platform is a mighty automation tool which requires no coding skills or in-depth system knowledge, letting both specialists and management build automation for practically any application.  
Now it can be used for running smoke, functional, acceptance tests, generating reports and a lot more in VSTS and TFS. 

# What’s the deal?  
The extension makes integration as easy as pie. Here is what you can do:  
1.	Run automated cases using TFS build tasks or VSTS Agent phases  
2.	Automatically receive case execution log  
3.	Build status-based results  
4.	Generate xml report file in JUnit format  
5.	Write case execution log into file and attach it to build results  

# A couple of requirements
There are just a few requirements to ensure seamless integration:  
1.	.NET framework 4.5  
2.	Team Foundation Server 2015 Update 1 minimum required  
3.	Agent version 1.83.0 minimum required  

# Install it in a matter of minutes
You are just one click away from seamless automation! Simply install the extension from Visual Studio Marketplace.  
If you have already installed previous version, delete it first!  
Alternatively, you can install it using gallery extension manager. Here is a small guide for you:
1.	Open extensions manager {hostname}/tfs/_gallery/manage 
2.	Click "Upload new extension" button
3.	Choose downloaded VSIX-file 
4.	Install extension
5.	Choose Collection
6.	Click Confirm button
7.	Now you are ready to get started!

# Update 3.1.0  
- For LEAPWORK version 2018.2.262  
- Fixed bug when JUnit parser interpreted all flows as "Failed"  
- Uses new Leapwork v3 API, API v2 is not supported

# Let’s get you started!
1.	 TFS: Add build step "Leapwork Integration" to your build definition.   
VSTS: Add an agent phase and then add task "Leapwork Integration" to this agent phase.  
2.	Enter Leapwork controller hostname or IP-address. 
3.  Enter your LEAPWORK controller API port, by default it is 9001.  
4.	Enter time delay (in seconds, by default 5 seconds). While schedule is running extension will check schedule state with a specified delay before the schedule is finished.  
5.	Select how the extension should process "Done" status: “Success” or “Failed”.  
6.	Enter JUnit report file name. The default name is "report.xml".  
7.	Enter schedule titles you want to run, each one must be entered from a new line.  
8.	Enter schedule ids, each one must be entered from a new line. Using id is preferable, because schedule title can be changed.  
9.	TFS: Add Test build step "Publish Test Results" to your build definition.  
VSTS: Add a task " Publish Test Results " to the previously added agent phase.  
10.	Choose JUnit report format. Enter JUnit report file name. It MUST be the same you've entered in point 5. 
11.	Run your schedule and get results. Enjoy! 
12.	Check build results in a detailed report, which can be found in Build summary  
13.	Download logs if needed.  

# Troubleshooting
- If you catch an error "No such run [runId]!" after schedule starting, increase time delay parameter.

# Do you need a perfectly tailored solution? 
Visit our [website](http://customatics.com) or drop us a line by email info@customatics.com

# Screenshots
  
![image](http://customatics.com/wp-content/uploads/2017/09/screen1.png)

![image](http://customatics.com/wp-content/uploads/2017/09/screen2.png)

![image](http://customatics.com/wp-content/uploads/2017/09/screen3.png)
